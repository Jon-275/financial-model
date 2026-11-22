"""Presentation for the working tabs; cell addresses and calculations stay fixed."""

from lxml import etree as E
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter

from report_layout import X, _columns, _rows, _merges, _link


NAVIGATION = {
    'Asset_Allocation': ('H', (
        ('Choose profiles on Inputs', 'Inputs!A15'),
        ('Returns and income', 'Asset_Allocation!A4'),
        ('Return correlations', 'Asset_Allocation!A14'),
        ('Edit account mixes', 'Asset_Allocation!A23'),
        ('Compare profile returns', 'Asset_Allocation!A52'),
        ('Selected portfolio', 'Asset_Allocation!A68'),
        ('Check selected profiles', 'Asset_Allocation!A93'),
        ('Returns used by the plan', 'Asset_Allocation!A97'),
    )),
    'Income_Targeting': ('H', (
        ('Choose age and amount', 'Income_Targeting!A4'),
        ('Compare results', 'Income_Targeting!A22'),
        ('Set the chart range', 'Income_Targeting!A33'),
        ('View conversion charts', 'Income_Targeting!A44'),
        ('Open conversion detail', "'Conversion Detail'!A1"),
    )),
    'Risk_Settings': ('E', (
        ('Set spending cuts', 'Risk_Settings!A4'),
        ('Edit early-loss tests', 'Risk_Settings!A12'),
        ('Set market simulation', 'Risk_Settings!A28'),
        ('Review last run', 'Risk_Settings!A42'),
        ('Run market risk on Summary', 'Summary!K20'),
        ('View risk results', "'Risk Results'!B2"),
    )),
    'Tax_Tables': ('G', (
        ('Federal income tax', 'Tax_Tables!A4'),
        ('California income tax', 'Tax_Tables!A28'),
        ('Payroll and saving limits', 'Tax_Tables!A49'),
        ('Marketplace credits', 'Tax_Tables!A202'),
        ('Medicare surcharges', 'Tax_Tables!A87'),
        ('Single-filer tables', 'Tax_Tables!A141'),
        ('Published tax-year updates', 'Tax_Tables!A230'),
        ('Roth IRA limits', 'Tax_Tables!A253'),
    )),
}


def configure(put):
    for sheet, (column, links) in NAVIGATION.items():
        put(sheet, f'{column}1', 'On this sheet', 'report_section')
        put(sheet, f'{column}2', 'Back to Summary', 'report_link')
        context = 'Shared by all scenarios.' if sheet == 'Tax_Tables' else '= "Scenario: "&Inputs!B7'
        put(sheet, f'{column}3', context, 'report_caption')
        for row, (label, _) in enumerate(links, 4):
            put(sheet, f'{column}{row}', label, 'report_link')
        row = len(links) + 6
        put(sheet, f'{column}{row}',
            'Blue cells are editable.' if sheet != 'Tax_Tables' else 'Reference tables used by the model.',
            'report_caption')
    # The tax index now sits beside the table rather than at the far edge.
    for row in (1, *range(3, 10)):
        put('Tax_Tables', f'J{row}', '')
    put('Asset_Allocation', 'A2', 'Set returns and account mixes here. Choose working and retirement profiles on Inputs.', 'report_caption')
    put('Asset_Allocation', 'A3', 'Edit blue cells. Each account must total 100%. Returns include income and price changes.', 'report_caption')
    put('Income_Targeting', 'A2', 'Choose an age and a conversion amount. Compare the result with your active plan.', 'report_caption')
    put('Income_Targeting', 'A3', 'This comparison does not change your plan. Amounts use dollars for the selected year.', 'report_caption')
    put('Income_Targeting', 'A5', '1. Age to compare (0 = first retirement year)', 'label')
    put('Income_Targeting', 'A21', '2. Roth conversion to test', 'label')
    put('Income_Targeting', 'A36', '3. Largest amount to show on charts', 'label')
    put('Risk_Settings', 'A2', 'Set the spending cuts you could accept, then run market risk on Summary.', 'report_caption')
    put('Risk_Settings', 'A3', 'Blue cells are editable. Rerun market risk after changing your plan or these settings.', 'report_caption')
    put('Risk_Settings', 'A12', 'Early-loss tests', 'detail_band')
    put('Risk_Settings', 'A28', 'Monte Carlo simulation', 'detail_band')
    put('Tax_Tables', 'A2', 'Reference values for federal and California calculations. Source years and links appear below each table.', 'report_caption')
    for row, title in {
        56: 'Contribution limits · 2026', 66: 'Marketplace premium credit · 2026',
        73: 'Net investment income tax', 79: 'Alternative minimum tax · Joint · 2026',
        87: 'Medicare income surcharges · 2026',
        134: 'Social Security taxable income thresholds · Joint',
        142: 'Federal income tax · Single · 2026',
        152: 'Federal long-term gains · Single · 2026',
        158: 'California income tax · Single · 2025', 170: 'Single-filer deductions and limits',
        181: 'Medicare income thresholds · Single · 2026',
        202: 'Marketplace premium credit · 2026',
        212: 'Poverty guidelines for Marketplace coverage',
    }.items():
        put('Tax_Tables', f'A{row}', title, 'detail_band')
    put('Inputs', 'J25', 'How to edit', 'report_section')
    put('Inputs', 'J26', 'Blue cells: enter your assumptions.', 'report_caption')
    put('Inputs', 'J27', 'Active value: the scenario used by the plan.', 'report_caption')
    put('Inputs', 'J28', 'Changing a scenario updates Summary. Rerun market risk to refresh simulations.', 'report_caption')


def _content(cell):
    inline = cell.find(X + 'is')
    return (cell.find(X + 'f') is not None or cell.findtext(X + 'v') not in (None, '')
            or (inline is not None and bool(''.join(inline.itertext()))))


def _cell(root, address):
    data = root.find(X + 'sheetData')
    _, number = coordinate_from_string(address)
    row = data.find(f"{X}row[@r='{number}']")
    if row is None:
        row = E.SubElement(data, X + 'row', r=str(number))
        data[:] = sorted(data, key=lambda item: int(item.get('r')))
    cell = row.find(f"{X}c[@r='{address}']")
    if cell is None:
        cell = E.SubElement(row, X + 'c', r=address)
        row[:] = sorted(row, key=lambda item: column_index_from_string(coordinate_from_string(item.get('r'))[0]))
    return cell


def _span(root, first, last, row, style, merges):
    """Only span unused cells: never merge away a value or formula."""
    start, end = map(column_index_from_string, (first, last))
    for index in range(start + 1, end + 1):
        if _content(_cell(root, f'{get_column_letter(index)}{row}')):
            raise ValueError(f'Presentation would cover {get_column_letter(index)}{row}')
    area = f'{first}{row}:{last}{row}'
    if area not in merges:
        merges.append(area)
    for index in range(start, end + 1):
        _cell(root, f'{get_column_letter(index)}{row}').set('s', style)


def _freeze(root, rows):
    view = root.find(f'{X}sheetViews/{X}sheetView')
    for item in view.findall(X + 'pane') + view.findall(X + 'selection'):
        view.remove(item)
    E.SubElement(view, X + 'pane', ySplit=str(rows), topLeftCell=f'A{rows+1}', state='frozen', activePane='bottomLeft')
    E.SubElement(view, X + 'selection', pane='bottomLeft', activeCell=f'A{rows+1}', sqref=f'A{rows+1}')


def apply(name, root, styles, numbers):
    if name not in {*NAVIGATION, 'Planning Metrics', 'Inputs'}:
        return
    if name == 'Inputs':
        _cell(root, 'B4').set('s', styles['input_value', '1'])
        for index, height in ((25, 30), (26, 36), (27, 42), (28, 60)):
            row = root.find(f"{X}sheetData/{X}row[@r='{index}']")
            row.set('ht', str(height))
        return
    widths = {
        'Planning Metrics': (56, 24, 76, 4, 29),
        'Asset_Allocation': (45, 20, 20, 20, 20, 20, 4, 33),
        'Income_Targeting': (54, 21, 68, 21, 22, 30, 4, 34),
        'Risk_Settings': (46, 20, 72, 4, 35),
        'Tax_Tables': (59, 20, 25, 37, 22, 4, 35),
    }
    _columns(root, widths[name])
    data = root.find(X + 'sheetData')
    merges = [item.get('ref') for item in root.iter(X + 'mergeCell')]
    if name in NAVIGATION:
        column, links = NAVIGATION[name]
        _link(root, {f'{column}2': 'Summary!B3',
                     **{f'{column}{row}': target for row, (_, target) in enumerate(links, 4)}})
        for row in (2, *range(4, len(links) + 4)):
            _cell(root, f'{column}{row}').set('s', styles['report_link'])
    # Remove the old cell grid and standardize body type without changing formats.
    for cell in data.iter(X + 'c'):
        column, row = coordinate_from_string(cell.get('r'))
        old = cell.get('s', '0')
        advanced = styles.get(('advanced_input', old))
        if not _content(cell):
            cell.set('s', styles['schedule', '0'])
        elif advanced and cell.find(X + 'f') is None:
            cell.set('s', advanced)
        elif old not in {styles[key] for key in ('title', 'band', 'detail_title', 'detail_band', 'report_table_head', 'header', 'report_link', 'report_section', 'report_caption', 'report_active')}:
            if _content(cell):
                cell.set('s', styles['schedule', numbers[int(old)]])
    from workbook_protection import _entry_cells
    for address in _entry_cells(name, root):
        cell = _cell(root, address)
        code = numbers[int(cell.get('s', '0'))]
        cell.set('s', styles['input_value', code])
    # Preserve existing merged ranges, then extend titles and section bars.
    last = {'Asset_Allocation': 'F', 'Income_Targeting': 'D', 'Risk_Settings': 'C', 'Tax_Tables': 'E'}.get(name)
    if last:
        for row in (1, 2, 3):
            # Tax row 3 is blank; the others contain the short usage instruction.
            for area in list(merges):
                if area.startswith(f'A{row}:'):
                    merges.remove(area)
            _span(root, 'A', last, row, styles['detail_title' if row == 1 else 'report_caption'], merges)
    sections = {
        'Asset_Allocation': (4, 14, 23, 24, 33, 42, 52, 68, 82, 93, 97),
        'Income_Targeting': (4, 22, 33),
        'Risk_Settings': (12, 28, 38, 42),
        'Tax_Tables': (),
        'Planning Metrics': (),
    }[name]
    for row in sections:
        for area in list(merges):
            if area.startswith(f'A{row}:'):
                merges.remove(area)
        _span(root, 'A', last, row, styles['detail_band'], merges)
    if name == 'Asset_Allocation':
        for column in 'BCDEF':
            _cell(root, f'{column}15').set('s', styles['report_table_head'])
        for row in (24, 33, 42):
            _cell(root, f'A{row}').set('s', styles['input_value', '0'])
        for row in (12, 22, 91):
            _span(root, 'A', 'F', row, styles['report_caption'], merges)
        _rows(root, 107, 25, {1: 38, 2: 28, 3: 28, 5: 44, 12: 44, 15: 44, 22: 30,
                             25: 34, 34: 34, 43: 34, 53: 42, 70: 38, 82: 38,
                             91: 30, 98: 38, 106: 36, 107: 36,
                             **dict.fromkeys((11, 13, 21, 32, 41, 50, 51, 67, 77, 81, 88, 92, 96), 9)})
        _freeze(root, 3)
    elif name == 'Income_Targeting':
        for row in (*range(5, 22), *range(24, 41)):
            if row == 33:
                continue
            _cell(root, f'A{row}').set('s', styles['report_label'])
            if row != 23:
                cell = _cell(root, f'C{row}')
                if cell.get('t') in ('s', 'inlineStr'):
                    cell.set('s', styles['report_caption'])
        _cell(root, 'B5').set('s', styles['input_value', '1'])
        _cell(root, 'B21').set('s', styles['input_value', '1'])
        _rows(root, 41, 30, {1: 38, 2: 28, 3: 28, 4: 30, 5: 42, 21: 36, 23: 40, 30: 44, 32: 40, 36: 42, 41: 32})
        _freeze(root, 3)
    elif name == 'Risk_Settings':
        for row in (*range(5, 11), *range(29, 37)):
            _cell(root, f'A{row}').set('s', styles['report_label'])
            _cell(root, f'C{row}').set('s', styles['report_caption'])
        _rows(root, 44, 29, {1: 38, 2: 28, 3: 28, 4: 34, 6: 38, 10: 40,
                            11: 10, 13: 34, 27: 10, 29: 38, 30: 38, 32: 38,
                            37: 10, 39: 36, 40: 36, 41: 10, 43: 34, 44: 55})
        _freeze(root, 3)
    elif name == 'Tax_Tables':
        links = root.find(X + 'hyperlinks')
        if links is not None:
            for link in list(links):
                if link.get('ref', '').startswith('J'):
                    links.remove(link)
        _rows(root, 265, 27, {1: 38, 2: 36, 3: 10})
        for row in data:
            i = int(row.get('r'))
            if i <= 3:
                continue
            a = row.find(f"{X}c[@r='A{i}']")
            populated = [cell for cell in row if _content(cell)
                         and column_index_from_string(coordinate_from_string(cell.get('r'))[0]) <= 5]
            if a is not None and len(populated) == 1:
                text = ''.join(a.itertext())
                heading = a.get('s') in (styles['detail_band'], styles['band']) or (len(text) < 100 and text.isupper())
                _span(root, 'A', 'E', i, styles['detail_band' if heading else 'report_caption'], merges)
                row.set('ht', '38' if len(text) > 145 else '30')
            elif any(cell.get('s') in (styles['header'], styles['report_table_head']) for cell in row):
                row.set('ht', '46')
            elif a is not None:
                row.set('ht', '34' if a.get('t') in ('inlineStr', 's') else '25')
                if a.get('t') in ('inlineStr', 's'):
                    a.set('s', styles['report_label'])
            # Right-hand source notes wrap within a predictable width.
            d = row.find(f"{X}c[@r='D{i}']")
            if d is not None and d.get('t') in ('inlineStr', 's') and d.get('s') not in (styles['header'], styles['report_table_head']):
                if not _content(_cell(root, f'E{i}')):
                    _span(root, 'D', 'E', i, styles['report_caption'], merges)
                    row.set('ht', '42' if len(''.join(d.itertext())) > 100 else '32')
        _freeze(root, 3)
    elif name == 'Planning Metrics':
        _cell(root, 'A1').set('s', styles['detail_title'])
        # Leave notes readable and retain the established result striping.
        for row in data:
            i = int(row.get('r'))
            if i not in (1, 2, 7, 14, 22, 37, 44, 49, 13, 21, 36, 43, 48):
                row.set('ht', '38' if i in (8, 18, 20, 38) else '30')
                _cell(root, f'A{i}').set('s', styles['report_label'])
                _cell(root, f'C{i}').set('s', styles['report_caption'])
        _cell(root, 'A3').set('s', styles['report_section'])
        _freeze(root, 2)
    _merges(root, merges)


def conversion_charts(drawing):
    """Stack native charts below the comparison instead of off the right edge."""
    d = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
    for index, anchor in enumerate(drawing):
        for endpoint, row in (('from', 43 + index * 22), ('to', 63 + index * 22)):
            node = anchor.find(d + endpoint)
            if node is None:
                continue
            node.find(d + 'col').text = '0' if endpoint == 'from' else '4'
            node.find(d + 'row').text = str(row)
            for key in ('colOff', 'rowOff'):
                node.find(d + key).text = '0'


def conversion_chart_style(root):
    c = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
    a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    for font in root.iter(a + 'latin'):
        font.set('typeface', 'Arial')
    for axis in root.findall('.//' + c + 'catAx'):
        for key in ('tickLblSkip', 'tickMarkSkip'):
            node = axis.find(c + key)
            if node is None:
                node = E.Element(c + key, val='5')
                following = axis.find(c + 'noMultiLvlLbl')
                if following is None:
                    following = axis.find(c + 'extLst')
                if following is None:
                    axis.append(node)
                else:
                    following.addprevious(node)
            else:
                node.set('val', '5')
        for body in axis.findall(f'{c}txPr/{a}bodyPr'):
            body.set('rot', '0')
