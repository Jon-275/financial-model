"""Generate Excel styling, navigation and print settings for the planning views."""

from copy import deepcopy
import investment_results

from lxml import etree as E
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter

import input_record
from workbook_guide import TABLE_GUIDES, TAX_LINKS
from plan_report import APPENDIX, CONVERSION_DETAIL, DECISION_TABS, DETAIL_GROUPS, DOLLAR_VIEW_CELL, FUTURE_DOLLARS, INDEX, INDEX_LINK_ROW, METRICS, METRICS_REF, METRIC_SECTIONS, REPORT_PAGE_BREAKS, TODAY_DOLLARS, index_rows
import risk_presentation as risk

X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
INK = '182C42'
MUTED = '455A6C'
ACCENT = '126761'
BLUE = '135A9C'
PALE = 'E9F2F8'
LINE = 'C9D9E4'
TAB_COLORS = {
    'README': ACCENT, 'Inputs': BLUE, 'Summary': INK,
    APPENDIX: '48617A', INDEX: '718090', risk.SHEET: '48617A',
    **dict.fromkeys(DECISION_TABS, '3E686A'),
}
DETAIL_SHEETS = {*DECISION_TABS, *(name for _, schedules in DETAIL_GROUPS for name, _, _ in schedules)}
SOLVERS = {'Funding_Calc', 'Targeting_Calc'}
ANNUAL_TABLES = {'Cash_Flow': 'AW', 'Projection': 'CG', 'Roth_Access': 'AP',
                 'Healthcare': 'AF', 'Checks': 'M'}
FORMATS = {
    'detail_title': (17, INK, True, None, None),
    'detail_band': (11, INK, True, 'EDF4F7', None),
    'report_title': (25, INK, True, None, None),
    'report_result': (16, INK, True, 'EDF4F7', None),
    'report_section': (13, INK, True, None, None),
    'report_amount': (19, INK, True, None, "report_currency"),
    'report_percent_amount': (19, ACCENT, True, None, 10),
    'report_body': (11, INK, False, None, None),
    'report_caption': (10, MUTED, False, None, None),
    'report_kicker': (10, ACCENT, True, None, None),
    'report_link': (10, BLUE, True, None, None),
    'report_table_head': (10, 'FFFFFF', True, INK, None),
    'report_row': (11, INK, False, None, 1),
    'report_row_alt': (11, INK, False, 'F2F6F9', 1),
    'report_label': (11, INK, False, None, None),
    'report_label_alt': (11, INK, False, 'F2F6F9', None),
    'report_case_age': (11, INK, False, None, 1),
    'report_case_age_alt': (11, INK, False, 'F2F6F9', 1),
    'report_money': (11, INK, False, None, "report_currency"),
    'report_money_alt': (11, INK, False, 'F2F6F9', 'report_currency'),
    'report_cost': (11, INK, False, None, 'report_currency_cents'),
    'report_cost_alt': (11, INK, False, 'F2F6F9', 'report_currency_cents'),
    'report_percent': (11, INK, False, None, 10),
    'report_caption_gutter': (10, MUTED, False, None, None),
    'report_section_gutter': (13, INK, True, None, None),
    'report_write_in': (11, BLUE, False, None, None),
    'report_active': (11, INK, True, PALE, 1),
    'report_decimal': (11, INK, False, None, 2),
}


def _font(parent, size, color, bold, underline=False):
    font = E.SubElement(parent, X + 'font')
    if bold:
        E.SubElement(font, X + 'b')
    if underline:
        E.SubElement(font, X + 'u')
    E.SubElement(font, X + 'sz', val=str(size))
    E.SubElement(font, X + 'color', rgb='FF' + color)
    E.SubElement(font, X + 'name', val='Arial')
    E.SubElement(font, X + 'family', val='2')
    return len(parent) - 1


def _fill(parent, color):
    fmt = E.SubElement(E.SubElement(parent, X + 'fill'), X + 'patternFill', patternType='solid')
    E.SubElement(fmt, X + 'fgColor', rgb='FF' + color)
    E.SubElement(fmt, X + 'bgColor', indexed='64')
    return len(parent) - 1


def install_styles(root, style_ids):
    num_formats = root.find(X + 'numFmts')
    currency = max(163, *(int(item.get('numFmtId')) for item in num_formats)) + 1
    E.SubElement(num_formats, X + 'numFmt', numFmtId=str(currency), formatCode='"$"#,##0')
    E.SubElement(num_formats, X + 'numFmt', numFmtId=str(currency + 1), formatCode='"$"#,##0.00')
    num_formats.set('count', str(len(num_formats)))
    fonts, fills, borders, xfs = (root.find(X + key) for key in ('fonts', 'fills', 'borders', 'cellXfs'))
    # Keep the Normal style's font metrics for the template's column widths.
    # Use the shared typeface on the worksheet cells themselves.
    body_font_id = str(len(fonts))
    fonts.append(deepcopy(fonts[0]))
    for xf in xfs:
        if xf.get('fontId', '0') == '0':
            xf.set('fontId', body_font_id)
            xf.set('applyFont', '1')
    for font in list(fonts)[1:]:
        if (name := font.find(X + 'name')) is not None:
            name.set('val', 'Arial')
        scheme = font.find(X + 'scheme')
        if scheme is not None:
            font.remove(scheme)
    border = E.SubElement(borders, X + 'border')
    for name in ('left', 'right', 'top', 'bottom', 'diagonal'):
        side = E.SubElement(border, X + name)
        if name == 'bottom':
            side.set('style', 'thin')
            E.SubElement(side, X + 'color', rgb='FF' + LINE)
    rule_border = len(borders) - 1
    format_ids = {'report_currency': currency, 'report_currency_cents': currency + 1}
    for name, (size, color, bold, fill, code) in FORMATS.items():
        style_ids[name] = str(len(xfs))
        code = format_ids.get(code, code)
        xf = E.SubElement(xfs, X + 'xf', xfId='0', numFmtId=str(code or 0),
                          fontId=str(_font(fonts, size, color, bold, name == 'report_link')),
                          fillId=str(_fill(fills, fill) if fill else 0),
                          borderId=str(rule_border if name == 'report_write_in' else 0),
                          applyFont='1', applyFill='1', applyBorder='1',
                          applyNumberFormat='1', applyAlignment='1')
        align = E.SubElement(xf, X + 'alignment', vertical='center')
        if name in ('report_table_head', 'report_body', 'report_result', 'report_caption',
                    'report_label','report_label_alt'):
            align.set('wrapText', '1')
        if name.endswith('_gutter'):
            align.set('horizontal', 'left')
            align.set('indent', '1')
            align.set('wrapText', '1')
        if name in ('report_amount', 'report_percent_amount'):
            align.set('horizontal', 'left')
        if name in ('report_money', 'report_money_alt', 'report_cost', 'report_cost_alt'):
            align.set('horizontal', 'right')
        if name in ('report_table_head', 'report_case_age', 'report_case_age_alt'):
            align.set('horizontal', 'center')

    # Replace the template's shared presentation styles in one place.
    for previous, current in (('title', 'detail_title'), ('band', 'detail_band'),
                              ('header', 'report_table_head')):
        xfs[int(style_ids[previous])] = deepcopy(xfs[int(style_ids[current])])

    # Preserve native input number formats while removing grid borders.
    for index in range(len(xfs)):
        original = xfs[index]
        if original.get('xfId') is None:
            continue
        if index >= min(int(style_ids[name]) for name in FORMATS):
            break
        base_code = original.get('numFmtId', '0')
        for role, color, bold, fill in (
            ('input_value', BLUE, False, 'EAF3FB'),
            ('input_selected', INK, True, 'F2F6F9'),
            ('input_label', INK, False, None),
            ('input_note', MUTED, False, None),
            ('input_section', INK, True, 'EDF4F7'),
        ):
            key = (role, base_code)
            if key in style_ids:
                continue
            style_ids[key] = str(len(xfs))
            xf = E.SubElement(xfs, X + 'xf', xfId='0', numFmtId=base_code,
                              fontId=str(_font(fonts, 10 if role == 'input_note' else 11, color, bold)),
                              fillId=str(_fill(fills, fill) if fill else 0), borderId='0',
                              applyFont='1', applyFill='1', applyBorder='1',
                              applyAlignment='1', applyNumberFormat='1')
            E.SubElement(xf, X + 'alignment', vertical='center',
                         wrapText='1' if role in ('input_label', 'input_note', 'input_section') else '0')
        font = fonts[int(original.get('fontId', '0'))]
        color = font.find(X + 'color')
        if color is not None and color.get('rgb', '').upper() == 'FF0000FF':
            style_ids[('advanced_input', str(index))] = style_ids[('input_value', base_code)]
    for group in (fonts, fills, borders, xfs):
        group.set('count', str(len(group)))
    schedule_font = xfs[int(style_ids['report_row'])].get('fontId')
    for code in sorted({xf.get('numFmtId', '0') for xf in xfs}, key=int):
        style_ids['schedule', code] = str(len(xfs))
        xf = E.SubElement(xfs, X + 'xf', xfId='0', numFmtId=code, fontId=schedule_font,
                          fillId='0', borderId='0', applyNumberFormat='1', applyFont='1',
                          applyFill='1', applyBorder='1', applyAlignment='1')
        E.SubElement(xf, X + 'alignment', vertical='center')
    xfs.set('count', str(len(xfs)))
    dxfs = root.find(X + 'dxfs')
    if dxfs is None:
        dxfs = E.Element(X + 'dxfs', count='0')
        xfs.addnext(dxfs)
    active = len(dxfs)
    for color in ('DCEAF5', 'FAE7E3', 'F2F6F9'):
        differential = E.SubElement(dxfs, X + 'dxf')
        pattern = E.SubElement(E.SubElement(differential, X + 'fill'), X + 'patternFill')
        E.SubElement(pattern, X + 'bgColor', rgb='FF' + color)
    dxfs.set('count', str(len(dxfs)))
    return active, [xf.get('numFmtId', '0') for xf in xfs]


def _stripe(root, area, key, fill):
    """Shade populated alternate rows. Reserved future rows use the Normal fill."""
    rules = root.findall(X + 'conditionalFormatting')
    priority = max((int(rule.get('priority')) for block in rules for rule in block), default=0) + 1
    block = E.Element(X + 'conditionalFormatting', sqref=area)
    rule = E.SubElement(block, X + 'cfRule', type='expression', priority=str(priority), dxfId=str(fill))
    E.SubElement(rule, X + 'formula').text = f'AND({key}<>"",MOD(ROW(),2)=1)'
    _replace_or_insert_format(root, block)


def _replace_or_insert_format(root, block):
    following = next((node for node in root if E.QName(node).localname in _AFTER_MERGES[2:]), None)
    if following is None:
        root.append(block)
    else:
        following.addprevious(block)


def _style_table(root, first_row, last_column, styles, numbers, *, cents=None):
    last = column_index_from_string(last_column)
    for row in root.find(X + 'sheetData'):
        if int(row.get('r')) < first_row:
            continue
        for attribute in ('s', 'customFormat', 'ht', 'customHeight', 'thickTop', 'thickBot'):
            row.attrib.pop(attribute, None)
        for cell in row:
            column, _ = coordinate_from_string(cell.get('r'))
            if column_index_from_string(column) > last:
                continue
            if cents and column_index_from_string(column) >= 4 and column not in ('G', 'AC'):
                cell.set('s', cents)
            else:
                code = numbers[int(cell.get('s', '0'))]
                cell.set('s', styles['schedule', code])
    for column in root.find(X + 'cols'):
        column.attrib.pop('style', None)


def install_solver_filter(root, last_column, last_row):
    _replace(root, 'autoFilter', ('sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
                                  *_AFTER_MERGES), ref=f'A1:{last_column}{last_row}')


def _replace(root, key, before, **attributes):
    if (previous := root.find(X + key)) is not None:
        root.remove(previous)
    item = E.Element(X + key, **attributes)
    following = next((node for node in root if E.QName(node).localname in before), None)
    if following is None:
        root.append(item)
    else:
        following.addprevious(item)
    return item


_AFTER_COLS = ('sheetData',)
_AFTER_MERGES = ('phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks',
                 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks',
                 'colBreaks', 'drawing', 'legacyDrawing', 'tableParts', 'extLst')
_AFTER_LINKS = _AFTER_MERGES[4:]
_AFTER_PRINT = ('drawing', 'legacyDrawing', 'tableParts', 'extLst')


def _columns(root, widths):
    cols = _replace(root, 'cols', _AFTER_COLS)
    for index, width in enumerate(widths, 1):
        E.SubElement(cols, X + 'col', min=str(index), max=str(index), width=str(width), customWidth='1')


def _rows(root, end, default, custom, *, first=1):
    data = root.find(X + 'sheetData')
    rows = {int(row.get('r')): row for row in data}
    for index in range(first, end + 1):
        if index not in rows:
            rows[index] = E.SubElement(data, X + 'row', r=str(index))
        row = rows[index]
        row.set('ht', str(custom.get(index, default)))
        row.set('customHeight', '1')
        row.attrib.pop('hidden', None)
    data[:] = sorted(rows.values(), key=lambda row: int(row.get('r')))


def _merges(root, ranges):
    block = _replace(root, 'mergeCells', _AFTER_MERGES, count=str(len(ranges)))
    for address in ranges:
        E.SubElement(block, X + 'mergeCell', ref=address)


def _link(root, destinations):
    links = _replace(root, 'hyperlinks', _AFTER_LINKS)
    for origin, target in destinations.items():
        E.SubElement(links, X + 'hyperlink', ref=origin, location=target)


def _print(root, pages=(), *, appendix=False):
    _replace(root, 'printOptions', ('pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', *_AFTER_PRINT),
             gridLines='0', headings='0', horizontalCentered='1')
    _replace(root, 'pageMargins', ('pageSetup', 'headerFooter', 'rowBreaks', *_AFTER_PRINT),
             left='0.35', right='0.35', top='0.42', bottom='0.45', header='0.18', footer='0.20')
    properties = root.find(X + 'sheetPr')
    if properties is None:
        properties = E.Element(X + 'sheetPr')
        root.insert(0, properties)
    setup = _replace(properties, 'pageSetUpPr', (), fitToPage='1' if appendix else '0', autoPageBreaks='0')
    setup.tail = None
    settings = dict(fitToWidth='1', fitToHeight='0') if appendix else {'scale': '100'}
    _replace(root, 'pageSetup', ('headerFooter', 'rowBreaks', 'colBreaks', *_AFTER_PRINT),
             paperSize='1', orientation='portrait', **settings)
    footer = _replace(root, 'headerFooter', ('rowBreaks', 'colBreaks', *_AFTER_PRINT))
    E.SubElement(footer, X + 'oddFooter').text = '&L&"Arial,Regular"&9Planning estimate. US federal and California taxes.&R&"Arial,Regular"&9Page &P of &N'
    if pages:
        breaks = _replace(root, 'rowBreaks', ('colBreaks', *_AFTER_PRINT),
                          count=str(len(pages)), manualBreakCount=str(len(pages)))
        for row in pages:
            E.SubElement(breaks, X + 'brk', id=str(row), min='0', max='16383', man='1')


def _select_top(root, cell):
    view = root.find(f'{X}sheetViews/{X}sheetView')
    view.attrib.pop('topLeftCell', None)
    for child in list(view):
        if child.tag in (X + 'pane', X + 'selection'):
            view.remove(child)
    E.SubElement(view, X + 'selection', activeCell=cell, sqref=cell)


def _collapse_empty_return_rows(root):
    # Keep saved scenario addresses stable. Only entirely empty retired rows may be collapsed.
    for row in root.find(X + 'sheetData'):
        if int(row.get('r')) in investment_results.RETIRED_ROWS:
            if any(cell.find(X + 'f') is not None or cell.find(X + 'v') is not None or
                   ''.join(cell.itertext()).strip() for cell in row):
                raise ValueError(f"Retired input row {row.get('r')} still contains data")
            for cell in list(row):
                row.remove(cell)
            row.attrib.pop('outlineLevel', None)
            row.attrib.pop('collapsed', None)
            row.set('hidden', '1')
            row.set('ht', '0')
            row.set('customHeight', '1')


def _style_inputs(root, styles, number_formats, active_fill):
    view = root.find(f'{X}sheetViews/{X}sheetView')
    view.set('showGridLines', '0')
    _columns(root, (39, 16, 11, 15, 15, 17, 43, 10, 3, 29))
    all_entries = {
        'A1': 'report_title', 'A2': 'report_body', 'A3': 'report_caption', 'A4': 'report_section',
        'B4': 'report_active', 'C4': 'report_caption', 'G4': 'report_link',
    }
    headers = {6, 74}
    sections = {7, 15, 25, 32, 36, 42, 44, 46, 53, 58, 70, 73, 85, 100, 130}
    editable = []
    for row in root.find(X + 'sheetData'):
        row_index = int(row.get('r'))
        if row_index <= 135:
            row.set('ht', str({1: 38, 2: 28, 4: 25, 6: 30}.get(row_index, 30 if row_index in sections else 27)))
            row.set('customHeight', '1')
        for cell in row:
            coordinate = cell.get('r')
            column, _ = coordinate_from_string(coordinate)
            code = number_formats[int(cell.get('s', '0'))]
            style = all_entries.get(coordinate)
            if not style and column in 'ABCDEFGH' and row_index >= 6:
                formula = cell.find(X + 'f') is not None
                content = cell.find(X + 'v') is not None or cell.find(X + 'is') is not None
                if row_index in headers:
                    style = 'report_table_head'
                elif column in 'DEF' and not formula and (content or row_index in (8, 10, 12, 36, 37, 53)):
                    style = ('input_value', code)
                    editable.append(coordinate)
                elif not formula and ((column in 'ABC' and 75 <= row_index <= 82)
                                      or (column == 'B' and 130 <= row_index <= 132)):
                    style = ('input_value', code)
                elif column == 'B' and formula:
                    style = ('input_selected', code)
                elif column == 'G':
                    style = ('input_note', code)
                else:
                    style = ('input_section' if row_index in sections else 'input_label', code)
            if style:
                cell.set('s', styles[style])
    _collapse_empty_return_rows(root)
    _merges(root, ['A2:G2', 'A3:G3', 'C4:F4', 'G4:H4'])
    _link(root, {
        'G4': "'Summary'!B3", 'J5': 'Inputs!A8', 'J6': 'Inputs!A25',
        'J7': 'Inputs!A36', 'J8': 'Inputs!A44', 'J9': 'Inputs!A53',
        'J10': 'Inputs!A73', 'J11': 'Inputs!A85', 'J12': 'Inputs!A101',
        'J13': 'Inputs!A130', 'J15': "'Summary'!B3",
        'J16': f'Summary!B{input_record.FIRST_ROW + 1}',
    })
    existing = root.findall(X + 'conditionalFormatting')
    for block in existing:
        for rule in block:
            rule.set('priority', str(int(rule.get('priority', '1')) + 2))
    def formula_rule(addresses, priority, style, formula):
        cf = E.Element(X + 'conditionalFormatting', sqref=addresses)
        rule = E.SubElement(cf, X + 'cfRule', type='expression', priority=str(priority), dxfId=str(style))
        E.SubElement(rule, X + 'formula').text = formula
        following = next((child for child in root if E.QName(child).localname in _AFTER_MERGES[2:]), None)
        if following is not None:
            following.addprevious(cf)
        else:
            root.append(cf)
    formula_rule(' '.join(editable), 2, active_fill, '$B$4=COLUMN()-3')
    formula_rule('D8:F8 D10:F10 D53:F53', 1, active_fill + 1,
                 'AND($B$4=COLUMN()-3,ISBLANK(D8))')


def prepare_summary(root):
    """Replace the old result grid while retaining the template drawing relationship."""
    root.find(X + 'sheetData').clear()


def summary_merges():
    """Keep printable sections and side controls within their own columns."""
    single=(2,3,5,6,8,9,29,30,31,33,34,36,38,39,55,56,57,58,59,60,
            65,66,68,69,74,75,76,77,78,79,80,81,82,84,86,87,92,94,
            97,98,100,102,110,111,115,117,118,120,124)
    return ([f'B{row}:I{row}' for row in single]
            +[f'{left}{row}:{right}{row}' for row in (11,12,13)
              for left,right in (('B','E'),('F','I'))]
            +[f'{left}{row}:{right}{row}' for row in range(70,74)
              for left,right in (('B','D'),('E','F'),('H','I'))]
            +[f'{left}{row}:{right}{row}' for row in (*range(103,109),112,113,114)
              for left,right in (('B','D'),('H','I'))]
            +['B61:I62','B90:C90','D90:I90','B121:I122',
              'K22:K23','K25:K26','K27:K28','K31:K32',*input_record.merges()])


def apply(name: str, root: E._Element, style_ids: dict, numbers: list[str], active_fill: int,
          *, with_state_policy: bool = False) -> None:
    """Apply report presentation, including reserved caption cells on explicit candidates."""
    view = root.find(f'{X}sheetViews/{X}sheetView')
    view.set('showGridLines', '0')
    view.set('zoomScale', '100')
    view.set('zoomScaleNormal', '100')
    properties = root.find(X + 'sheetPr')
    if (color := TAB_COLORS.get(name)) is not None:
        if properties is None:
            properties = E.Element(X + 'sheetPr')
            root.insert(0, properties)
        _replace(properties, 'tabColor', ('outlinePr', 'pageSetUpPr'), rgb='FF' + color)
    elif properties is not None and (color_node := properties.find(X + 'tabColor')) is not None:
        properties.remove(color_node)
    if name == 'README':
        _columns(root, (3, 11, 11, 11, 11, 11, 11, 11, 11))
        _select_top(root, 'B3')
        from start_here import CONTENT, LINKS
        _merges(root, [f'{address}:I{address[1:]}' for address in CONTENT])
        _rows(root, 20, 18, {3: 38, 5: 30, 8: 30, 11: 34, 15: 36, 19: 30})
        _link(root, LINKS)
        root.find(X + 'dimension').set('ref', 'A1:I20')
    elif name == 'Inputs':
        _style_inputs(root, style_ids, numbers, active_fill)
    elif name == 'Summary':
        _columns(root, (3, 8, 7, 12.5, 12.5, 12.5, 12.5, 12.5, 12.5, 3, 28))
        _rows(root, input_record.LAST_ROW, 18, {3: 37, 8: 30, 9: 28, 12: 32, 34: 37, 40: 31, 55: 27,
                            58: 28, 66: 37, 70: 36, **dict.fromkeys(range(71, 74), 22),
                            31:25,69:27,74:25,75:30,78:27,79:27,80:25,82:31,84:22,87:31,92:27,94:27,
                            98:37,100:27,103:32,111:27,112:32,115:30,118:32,124:21,
                            **dict.fromkeys((*range(104,109),113,114),28),
                            **dict.fromkeys((67,83,85,88,89,91,93,95,96,99,101,109,116,119,123,125,126,127,128),6),
                            **{input_record.FIRST_ROW+1:37,input_record.FIRST_ROW+3:24,
                               input_record.SECOND_ROW+1:37,input_record.SECOND_ROW+3:24},
                            **dict.fromkeys((*range(input_record.FIRST_HEADERS+1,input_record.FIRST_HEADERS+13),
                                             *range(input_record.SECOND_HEADERS+1,input_record.SECOND_HEADERS+13)),23),
                            **dict.fromkeys(range(input_record.FIRST_ROW+25,input_record.FIRST_ROW+28),28)})
        policy_merge = ['B7:I7'] if with_state_policy else []
        _merges(root, summary_merges() + policy_merge)
        for row in range(71, 74):
            suffix = '_alt' if row % 2 else ''
            for column, style in (('G', 'report_case_age'), ('H', 'report_money')):
                cell = root.find(f"{X}sheetData/{X}row[@r='{row}']/{X}c[@r='{column}{row}']")
                cell.set('s', style_ids[style + suffix])
        _link(root, {
            'K5': 'Inputs!A1', 'K6': 'Summary!B34', 'K7': 'Summary!B66',
            'K8':'Summary!K29','K9':'Summary!B98','K10':"'Annual Detail'!B2",
            'K11': 'Appendix!B3', 'K12': "'Risk Results'!B2",
            'K13': 'Cash_Flow!A1', 'K14': 'Roth_Access!A1',
            'K15':'Risk_Settings!A1','K34':f'Summary!B{input_record.FIRST_ROW+1}',
            'K35':f'Summary!B{input_record.SECOND_ROW+1}',
            'K36':'Asset_Allocation!A1','K37':'Income_Targeting!A1',
            'K38':f'{METRICS_REF}A1','K39':'Tax_Tables!A1',
        })
        alert = E.Element(X + 'conditionalFormatting', sqref='B8:I8')
        rule = E.SubElement(alert, X + 'cfRule', type='expression', priority='1',
                            dxfId=str(active_fill + 1))
        E.SubElement(rule, X + 'formula').text = 'LEFT($B$8,LEN("First shortfall"))="First shortfall"'
        validations = root.find(X + 'dataValidations')
        if validations is None:
            validations = E.Element(X + 'dataValidations')
            root.find(X + 'hyperlinks').addprevious(validations)
        validations.addprevious(alert)
        dollar_view = E.SubElement(validations, X + 'dataValidation', type='list', sqref=DOLLAR_VIEW_CELL,
                                   showErrorMessage='1', errorTitle='Choose a dollar view',
                                   error='Select Today’s dollars or Future dollars.')
        E.SubElement(dollar_view, X + 'formula1').text = f'"{TODAY_DOLLARS},{FUTURE_DOLLARS}"'
        validations.set('count', str(len(validations)))
        _select_top(root, 'B3')
        _print(root, REPORT_PAGE_BREAKS)
    elif name == METRICS:
        _columns(root, (49, 23, 75, 4, 24))
        _rows(root, 50, 25, {1: 36, 2: 23, 3: 34, 6: 32, 8: 32, 17: 32, 18: 32,
                             20: 32, 25: 34, 27: 32, 34: 30,
                             **dict.fromkeys((13, 21, 36, 43, 48), 8)})
        _merges(root, [f'A{row}:C{row}' for row in (1, 2, *METRIC_SECTIONS)])
        for first, last in ((4, 6), (8, 12), (15, 20), (23, 35), (38, 42), (45, 47), (50, 50)):
            _stripe(root, f'A{first}:C{last}', f'$A{first}', active_fill + 2)
        _link(root, {'E2': 'Summary!B3', 'E3': "'Annual Detail'!B2", 'E4': 'Cash_Flow!A1'})
        _select_top(root, 'A1')
    elif name == APPENDIX:
        _columns(root, (3, 8, 7, 12, 12, 12, 12, 12, 12, 15, 3, 25))
        _rows(root, 5, 18, {2: 37, 4: 26, 5: 30})
        _merges(root, [f'B{row}:J{row}' for row in (2, 3, 4)])
        _stripe(root, 'B6:J131', '$B6', active_fill + 2)
        _link(root, {'L2': 'Summary!B3'})
        _print(root, appendix=True)
    elif name == INDEX:
        _columns(root, (3, 13, 13, 13, 13, 13, 13, 13, 13))
        _rows(root, INDEX_LINK_ROW, 23, {3: 39, 5: 27})
        sections, items = [], []
        for row, _, schedules in index_rows():
            sections.append(row)
            items.extend((row + offset, name) for offset, (name, _, _) in enumerate(schedules, 1))
        _merges(root, [f'B{row}:I{row}' for row in (3, 5, 6, *sections)]
                + [f'{left}{row}:{right}{row}' for row in (INDEX_LINK_ROW, *(row for row, _ in items))
                   for left, right in (('B', 'D'), ('E', 'I'))])
        targets = {sheet: guide[0]+'1' for sheet, guide in TABLE_GUIDES.items()}
        targets.update({APPENDIX: 'B2', risk.SHEET: 'B2'})
        _link(root, {**{f'B{row}': f"'{sheet}'!{targets.get(sheet, 'A1')}" for row, sheet in items},
                     f'B{INDEX_LINK_ROW}': 'Summary!B3'})
        _select_top(root, 'B3')
    elif name == risk.SHEET:
        _columns(root, (3, 15, 9, 15, 9, 15, 9, 15, 9, 15, 12, 4, 23, 20))
        _rows(root, 57, 21, {2: 39, 3: 25, 4: 27, 7: 32, 12: 27,
                                               25: 29, 27: 29, 34: 31, 39: 35, 57: 28})
        policy_merge = ['B9:I9'] if with_state_policy else []
        _merges(root, policy_merge + [f'B{row}:I{row}' for row in (2, 3, 4, 5, 10, 11, 30, 32, 33, 34, 36, 37, 54, 55)]
                + [f'{left}{row}:{right}{row}' for row in (6, 7, 8)
                   for left, right in (('B', 'E'), ('F', 'I'))]
                + [f'{left}{row}:{right}{row}' for row in range(12, 30)
                   for left, right in (('B', 'E'), ('F', 'G'), ('H', 'I'), ('J', 'K'))]
                + [f'{left}{row}:{right}{row}' for row in (*range(39, 53), *range(57, risk.LAST_BALANCE_ROW + 1))
                   for left, right in (('B', 'C'), ('D', 'E'), ('F', 'G'), ('H', 'I'), ('J', 'K'))])
        for row in root.find(X + 'sheetData'):
            number = int(row.get('r'))
            styles = (dict.fromkeys('FHJ', 'report_case_age') if number == 14 else
                      {'D': 'report_case_age', 'F': 'report_money',
                       'H': 'report_case_age', 'J': 'report_money'} if 40 <= number <= 52 else {})
            suffix = '_alt' if number % 2 else ''
            for cell in row:
                if (style := styles.get(cell.get('r').rstrip('0123456789'))) is not None:
                    cell.set('s', style_ids[style + suffix])
        _link(root, {'K2': 'Summary!B3', 'K3': 'Risk_Settings!A1', 'K4': 'Risk_Data!A1'})
        _stripe(root, f'B58:K{risk.LAST_BALANCE_ROW}', '$B58', active_fill + 2)
        _select_top(root, 'B2')
    elif name == risk.DATA_SHEET:
        widths = [18] * 54
        widths[:2] = (11, 9)
        widths[15:25] = [3] * 10
        widths[26:30] = (36, 24, 24, 12)
        widths[52:54] = (29, 59)
        _columns(root, widths)
        _merges(root, ['A1:I1', 'A2:I2', 'A3:C3', 'D3:E3', 'F3:G3', 'H3:I3', 'K1:O1', 'AF1:AX1'])
        _rows(root, 4, 20, {1: 33, 2: 26, 3: 24, 4: 37})
        _link(root, {'A3': "'Risk Results'!B2", 'D3': 'Risk_Data!K1',
                     'F3': 'Risk_Data!AA1', 'H3': 'Risk_Data!AF1'})
        for area, key in (('A5:I130', '$A5'), ('K5:O17', '$K5'),
                          ('AF5:AX7', '$AG5'), ('AA2:AD2476', '$AD2')):
            _stripe(root, area, key, active_fill + 2)
    elif name in DETAIL_SHEETS:
        if view.find(X + 'pane') is None:
            _select_top(root, 'A1')
        data = root.find(X + 'sheetData')
        # Keep the underlying numeric formats and give titles, notes and headers room.
        style_heights = {
            style_ids['title']: 32, style_ids['detail_title']: 32,
            style_ids['header']: 40, style_ids['report_table_head']: 40,
            style_ids['band']: 24, style_ids['detail_band']: 26,
        }
        tables_without_title = SOLVERS | {'Report_Data'}
        header_rows = {1} if name in tables_without_title else {
            'Projection': {4}, 'Cash_Flow': {4}, 'Roth_Access': {4}, 'Healthcare': {4},
            'Checks': {4}, 'Risk_Settings': {4, 13},
            'Asset_Allocation': {5, 15, 25, 34, 43},
            'Income_Targeting': {23},
            CONVERSION_DETAIL: {4},
            'Tax_Tables': {5, 16, 29, 143, 152, 159, 204, 220},
        }.get(name, set())
        for row in data[:8] if name in tables_without_title else data:
            index = int(row.get('r'))
            if index in header_rows:
                for cell in row:
                    if cell.get('s') == style_ids['header']:
                        cell.set('s', style_ids['report_table_head'])
            heights = [style_heights[cell.get('s')] for cell in row if cell.get('s') in style_heights]
            if index in (2, 3) and name not in tables_without_title:
                heights.append(24)
            if name in TABLE_GUIDES and 2 <= index <= 5:
                heights.append(24)
            if heights:
                row.set('ht', str(max(float(row.get('ht', '0')), *heights)))
                row.set('customHeight', '1')
        if name == 'Asset_Allocation':
            existing = [item.get('ref') for item in root.iter(X + 'mergeCell')]
            _merges(root, existing + ['A97:E97', 'A106:E106', 'A107:E107'])
            _rows(root, 107, 22, {97: 26, 98: 25, 106: 26, 107: 26}, first=97)
        if name in {'Asset_Allocation', 'Income_Targeting', 'Risk_Settings'}:
            for cell in data.iter(X + 'c'):
                style = style_ids.get(('advanced_input', cell.get('s', '0')))
                if style and cell.find(X + 'f') is None and cell.find(X + 'v') is not None:
                    cell.set('s', style)
        if name == 'Report_Data':
            root.find(X + 'cols')[15].set('width', '32')
            _stripe(root, 'A2:N127', '$A2', active_fill + 2)
            _stripe(root, 'AE2:AH127', '$AE2', active_fill + 2)
        if name in TABLE_GUIDES:
            first, _, notes, results = TABLE_GUIDES[name]
            start = column_index_from_string(first)
            end = get_column_letter(start + 4)
            _merges(root, [f'{first}{row}:{end}{row}' for row in (1, *range(2, len(notes) + 2), 7, 8)])
            _link(root, {f'{first}7': results, f'{first}8': f'{name}!A1'})
            root.find(X + 'dimension').set('ref', f'A1:{"AY" if name == "Targeting_Calc" else end}{int(data[-1].get("r"))}')
        if name in SOLVERS:
            cols = root.find(X + 'cols')
            for index, width in ((0, 11), (1, 9), (2, 24), (28, 28)):
                cols[index].set('width', str(width))
            # Local source values keep the large trial blocks shareable in Excel.
            # They remain visible beside the calculation guide.
            for index, width in ((56, 45), (57, 18)):
                E.SubElement(cols, X + 'col', min=str(index), max=str(index),
                             width=str(width), customWidth='1')
            last_column = max(column_index_from_string(coordinate_from_string(cell.get('r'))[0])
                              for cell in data.iter(X + 'c'))
            root.find(X + 'dimension').set('ref',
                f'A1:{get_column_letter(last_column)}{int(data[-1].get("r"))}')
            last = int(data[-1].get('r'))
            end_column = 'AP' if name == 'Targeting_Calc' else 'AO'
            _style_table(root, 2, end_column, style_ids, numbers, cents=style_ids['schedule', numbers[int(style_ids['cents'])]])
            _stripe(root, f'A2:{end_column}{last}', '$A2', active_fill + 2)
            install_solver_filter(root, end_column, last)
        if name in ANNUAL_TABLES:
            column = ANNUAL_TABLES[name]
            _style_table(root, 5, column, style_ids, numbers)
            _stripe(root, f'A5:{column}130', '$A5', active_fill + 2)
        if name == CONVERSION_DETAIL:
            _merges(root, ['A1:H1', 'A2:H2', 'A3:C3', 'D3:H3'])
            _link(root, {'A3': 'Income_Targeting!A1'})
            _style_table(root, 5, 'V', style_ids, numbers)
            _stripe(root, 'A5:V55', '$A5', active_fill + 2)
        if name == 'Income_Targeting':
            _link(root, {'A41': "'Conversion Detail'!A1"})
        if name == 'Cash_Flow':
            _merges(root, ['A2:I2', 'AN1:AQ1', 'AN2:AQ2', 'AR1:AW1', 'AR2:AW2'])
            _link(root, {'A3': 'Cash_Flow!AN1', 'AN3': 'Cash_Flow!A1'})
        if name == 'Risk_Settings':
            _merges(root, ['A2:C2', 'A3:C3'])
        if name == 'Tax_Tables':
            _link(root, {address: f'Tax_Tables!{target}' for address, (_, target) in TAX_LINKS.items()})


def install_print_names(workbook):
    sheets = [sheet.get('name') for sheet in workbook.find(X + 'sheets')]
    names = workbook.find(X + 'definedNames')
    print_metadata = {
        'Summary': {'_xlnm.Print_Area': f'Summary!$B$2:$I${input_record.LAST_ROW}'},
        APPENDIX: {
            '_xlnm.Print_Area': "'Annual Detail'!$B$2:INDEX('Annual Detail'!$J$5:$J$131,MAX(1,Report_Data!$Q$2+1))",
            '_xlnm.Print_Titles': "'Annual Detail'!$5:$5",
        },
    }
    input_index = str(sheets.index('Inputs'))
    for entry in list(names):
        if entry.get('localSheetId') == input_index and entry.get('name') in ('_xlnm.Print_Area', '_xlnm.Print_Titles'):
            names.remove(entry)
    for name, entries in print_metadata.items():
        index = str(sheets.index(name))
        for entry in list(names):
            if entry.get('localSheetId') == index and entry.get('name') in entries:
                names.remove(entry)
        for key, formula in entries.items():
            E.SubElement(names, X + 'definedName', name=key, localSheetId=index).text = formula
