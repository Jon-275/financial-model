"""Set up the saved scenarios and arrange the workbook around common tasks."""

from lxml import etree

from example_inputs import example_cells
from plan_report import APPENDIX, DECISION_TABS, DETAIL_GROUPS, INDEX
from workbook_metadata import make_sheets_visible

_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_PROPS = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
_VT = '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}'
_TAB_ORDER = (
    'Start Here', 'Inputs', 'Summary', *DECISION_TABS, APPENDIX, INDEX,
    *(name for _, schedules in DETAIL_GROUPS for name, _, _ in schedules),
)
_INPUT_GROUPS = ((17, 24), (55, 57), (59, 60), (102, 124))


def _tab_name(name):
    return 'Start Here' if name == 'README' else name


def sheet_order(names):
    """Sort sheet names without renaming the compiler's source worksheet."""
    return sorted(names, key=lambda name: _TAB_ORDER.index(_tab_name(name))
                  if _tab_name(name) in _TAB_ORDER else len(_TAB_ORDER))


def write_scenarios(put):
    """Seed the release workbook with fictional household inputs."""
    for address, value in example_cells().items():
        put('Inputs', address, value)


def _freeze(view, columns, rows):
    for tag in ('pane', 'selection'):
        for child in view.findall(_MAIN + tag):
            view.remove(child)
    if columns:
        first_col = chr(ord('A') + columns)
        etree.SubElement(view, _MAIN + 'pane', xSplit=str(columns), ySplit=str(rows),
                         topLeftCell=f'{first_col}{rows + 1}', activePane='bottomRight', state='frozen')
        for pane, cell in (
            ('topRight', f'{first_col}1'), ('bottomLeft', f'A{rows + 1}'),
            ('bottomRight', f'{first_col}{rows + 1}'),
        ):
            etree.SubElement(view, _MAIN + 'selection', pane=pane, activeCell=cell, sqref=cell)
    else:
        cell = f'A{rows + 1}'
        etree.SubElement(view, _MAIN + 'pane', ySplit=str(rows),
                         topLeftCell=cell, activePane='bottomLeft', state='frozen')
        etree.SubElement(view, _MAIN + 'selection', pane='bottomLeft', activeCell=cell, sqref=cell)


def _group_inputs(root):
    properties = root.find(_MAIN + 'sheetPr')
    if properties is None:
        properties = etree.Element(_MAIN + 'sheetPr')
        root.insert(0, properties)
    outline = properties.find(_MAIN + 'outlinePr')
    if outline is None:
        outline = etree.SubElement(properties, _MAIN + 'outlinePr')
    outline.set('summaryBelow', '0')
    outline.set('showOutlineSymbols', '1')
    root.find(_MAIN + 'sheetFormatPr').set('outlineLevelRow', '1')
    rows = {int(row.get('r')): row for row in root.find(_MAIN + 'sheetData')}
    for first, last in _INPUT_GROUPS:
        rows[first - 1].set('collapsed', '1')
        for number in range(first, last + 1):
            if number in rows:
                rows[number].set('outlineLevel', '1')
                rows[number].set('hidden', '1')


def _validate_inputs(root):
    rules = (
        ('B4', 'whole', 'between', 1, 3, 'Choose scenario 1, 2 or 3.'),
        ('D8:F10 D53:F53', 'whole', 'between', 0, 125, 'Enter a whole age from 0 to 125.'),
        ('D98:F98 D101:F101', 'whole', 'between', 0, 1, 'Enter 0 or 1.'),
        ('D55:F55', 'decimal', 'between', 0, 1, 'Enter a share from 0% to 100%.'),
        ('B130', 'whole', 'between', 2026, 9874, 'Enter a whole opening year from 2026 to 9874.'),
        ('B131', 'decimal', 'between', 0, 0.999999999999999, 'Enter a fee from 0% to less than 100%.'),
        ('B132', 'decimal', 'greaterThan', -1, None, 'Enter inflation above -100%, or leave this blank.'),
    )
    validations = root.find(_MAIN + 'dataValidations')
    if validations is None:
        validations = etree.Element(_MAIN + 'dataValidations')
        later = ('hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter',
                 'rowBreaks', 'colBreaks', 'drawing', 'legacyDrawing', 'tableParts', 'extLst')
        following = next((node for node in root if etree.QName(node).localname in later), None)
        if following is None:
            root.append(validations)
        else:
            following.addprevious(validations)
    for cells, kind, operator, first, last, message in rules:
        item = etree.SubElement(validations, _MAIN + 'dataValidation', sqref=cells, type=kind,
                                operator=operator, allowBlank='1' if cells == 'B132' else '0',
                                showErrorMessage='1', errorTitle='Check this input', error=message)
        etree.SubElement(item, _MAIN + 'formula1').text = str(first)
        if last is not None:
            etree.SubElement(item, _MAIN + 'formula2').text = str(last)
    validations.set('count', str(len(validations)))


def apply_sheet_layout(name, root):
    """Apply presentation metadata after values and formulas are assembled."""
    views = root.find(_MAIN + 'sheetViews')
    if views is None:
        return
    view = views.find(_MAIN + 'sheetView')
    if name == 'Inputs':
        _freeze(view, columns=2, rows=6)
        _group_inputs(root)
        _validate_inputs(root)
    if name in {'Cash_Flow', 'Roth_Access'}:
        _freeze(view, columns=2, rows=4)
    if name == 'README':
        view.set('tabSelected', '1')
        # The old introduction leaves formatted rows after its text is cleared.
        data = root.find(_MAIN + 'sheetData')
        for row in reversed(list(data)):
            if any(len(cell) for cell in row):
                break
            data.remove(row)
        root.find(_MAIN + 'dimension').set('ref', f"A1:A{data[-1].get('r')}")
    else:
        view.attrib.pop('tabSelected', None)


def _update_sheet_titles(app, names):
    """Keep the file properties in step with the worksheet list."""
    headings = app.find(f'{_PROPS}HeadingPairs/{_VT}vector')
    titles = app.find(f'{_PROPS}TitlesOfParts/{_VT}vector')
    if headings is None or titles is None:
        return
    offset = 0
    for label, count in zip(headings[::2], headings[1::2]):
        count_node = count.find(_VT + 'i4')
        size = int(count_node.text)
        if label.findtext(_VT + 'lpstr') == 'Worksheets':
            for title in titles[offset:offset + size]:
                titles.remove(title)
            for i, name in enumerate(names):
                title = etree.Element(_VT + 'lpstr')
                title.text = name
                titles.insert(offset + i, title)
            count_node.text = str(len(names))
            titles.set('size', str(len(titles)))
            return
        offset += size


def apply_tab_order(workbook, app=None):
    """Reorder tabs and preserve worksheet IDs and relationships."""
    sheets = workbook.find(_MAIN + 'sheets')
    original = [sheet.get('name') for sheet in sheets]
    by_name = dict(zip(original, sheets))
    ordered = sheet_order(original)
    indices = {name: index for index, name in enumerate(ordered)}

    for item in workbook.findall(f'{_MAIN}definedNames/{_MAIN}definedName'):
        if (local := item.get('localSheetId')) is not None:
            item.set('localSheetId', str(indices[original[int(local)]]))
    sheets[:] = [by_name[name] for name in ordered]
    for sheet in sheets:
        name = sheet.get('name')
        sheet.set('name', _tab_name(name))
    make_sheets_visible(workbook)

    start = str(indices.get('README', indices.get('Start Here', 0)))
    for view in workbook.findall(f'{_MAIN}bookViews/{_MAIN}workbookView'):
        view.set('activeTab', start)
        view.set('firstSheet', start)
    if app is not None:
        _update_sheet_titles(app, [_tab_name(name) for name in ordered])
