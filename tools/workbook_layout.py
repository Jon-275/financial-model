"""Set up the saved scenarios and arrange the workbook around common tasks."""

from lxml import etree

_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_PROPS = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
_VT = '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}'
_TAB_ORDER = (
    'Start Here', 'Inputs', 'Summary', 'Cash_Flow', 'Projection', 'Roth_Access',
    'Healthcare', 'Asset_Allocation', 'Income_Targeting', 'Risk_Settings',
    'Tax_Tables', 'Checks',
)
_HIDDEN = ('Report_Data', 'Funding_Calc', 'Targeting_Calc')
_INPUT_GROUPS = ((17, 24), (55, 57), (59, 60), (75, 82), (102, 124))


def _tab_name(name):
    return 'Start Here' if name == 'README' else name


def sheet_order(names):
    """Sort sheet names without renaming the compiler's source worksheet."""
    ordered = sorted(names, key=lambda name: _TAB_ORDER.index(_tab_name(name))
                     if _tab_name(name) in _TAB_ORDER else len(_TAB_ORDER))
    return [name for name in ordered if name not in _HIDDEN] + [name for name in _HIDDEN if name in ordered]


def write_scenarios(put, workbook):
    """Preserve the base case and make each saved alternative change one assumption."""
    inputs = workbook['Inputs']
    later_retirement = inputs['D10'].value + 5
    lower_spending = round(inputs['D37'].value * 0.9, 2)
    for address, value in {
        'D7': 'Base plan',
        'E7': f'Retire at {later_retirement}',
        'E10': later_retirement,
        'F7': 'Spend 10% less',
        'F37': lower_spending,
        'G7': 'The saved alternatives each change one input.',
        'G16': 'Choose a retirement investment profile. Expand rows 17–24 to see the rates.',
        'G54': 'Enter 999 to keep the two-person plan. Expand the rows below to enter survivor costs.',
        'G58': 'Enter 999 to leave long-term care out. Expand the rows below to enter the cost and duration.',
        'G101': 'Use 0 for the household budget. Use 1 and expand the rows below to enter costs by person.',
        'G125': 'Expand rows 102–124 to enter Marketplace quotes and prior Medicare income.',
        'A62': 'Legacy Marketplace income reference',
        'G62': 'Used only for the older Projection income ratio. The Healthcare sheet calculates the credit limit separately.',
    }.items():
        put('Inputs', address, value)

    for address, value in {
        'A1': 'RETIREMENT PLAN',
        'A3': 'QUICK START',
        'A4': '1. On Inputs, choose 1 for Base plan, 2 to retire later or 3 to spend less in retirement.',
        'A5': '2. Review ages, wages, balances and spending. Set the last age to show in row 53.',
        'A6': '3. Read the funding answer at the top of Summary.',
        'A7': 'If the plan falls short, open Cash_Flow to see what went unpaid in the first shortfall year.',
        'A12': 'Use the + signs at the left of Inputs to open sections such as healthcare and one-time expenses.',
        'A16': 'Cash_Flow shows how bills were paid. Projection and Roth_Access show balances and which savings were available.',
        'A18': 'Income_Targeting compares Roth conversions for one retirement year when money is available to convert.',
        'A19': 'Asset_Allocation sets the investment mix. Tax_Tables lists tax rates and limits.',
        'A20': 'Risk_Settings holds the inputs for the separate portfolio risk report.',
        'A22': 'The risk report recalculates costs and withdrawals for each market path. Rerun it after changing inputs.',
        'A23': '',
        'A39': 'The separate risk report uses asset volatility and correlations from Asset_Allocation.',
        'A89': 'PORTFOLIO RISK',
        'A90': 'The separate risk report recalculates the plan under different market returns. It shows how often costs were paid in full.',
        'A91': 'Choose the return method and early-retirement stress settings on Risk_Settings.',
        'A92': 'Check how often a shortfall occurs and when it first happens.',
        'A93': '',
        'A94': '',
        'A95': '',
        'A96': '',
        'A97': '',
        'A98': '',
        'A109': 'EARLY RETIREMENT LOSSES',
        'A110': 'The risk report compares fixed spending with temporary cuts to flexible spending during poor markets.',
        'A111': 'Risk_Settings shows the stock returns tested during the first retirement years.',
        'A112': 'The report counts any year with unpaid spending or taxes as a shortfall, even if the portfolio recovers later.',
        'A113': '',
        'A114': '',
        'A115': '',
        'A119': 'The projection and risk report stop at the end age you choose on Inputs, up to age 125.',
        'A120': '',
    }.items():
        put('README', address, value)


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


def apply_sheet_layout(name, root):
    """Apply presentation metadata after values and formulas are assembled."""
    views = root.find(_MAIN + 'sheetViews')
    if views is None:
        return
    view = views.find(_MAIN + 'sheetView')
    if name in {'README', 'Inputs', 'Summary'}:
        view.set('zoomScale', '100')
        view.set('zoomScaleNormal', '100')
    if name in {'README', 'Summary'}:
        view.set('showGridLines', '0')
    if name == 'Inputs':
        _freeze(view, columns=3, rows=6)
        _group_inputs(root)
    if name in {'Cash_Flow', 'Roth_Access'}:
        _freeze(view, columns=2, rows=4)
    if name == 'README':
        view.set('tabSelected', '1')
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
    """Reorder visible tabs and preserve worksheet IDs and relationships."""
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
        if name in _HIDDEN:
            sheet.set('state', 'hidden')

    start = str(indices.get('README', indices.get('Start Here', 0)))
    for view in workbook.findall(f'{_MAIN}bookViews/{_MAIN}workbookView'):
        view.set('activeTab', start)
        view.set('firstSheet', start)
    if app is not None:
        _update_sheet_titles(app, [_tab_name(name) for name in ordered])
