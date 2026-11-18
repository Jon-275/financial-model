"""Embed a locally calculated risk snapshot without saving through a workbook library."""

from datetime import date
from math import isfinite
from posixpath import join, normpath
import re
from zipfile import BadZipFile, ZipFile

from lxml import etree as E
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils.exceptions import InvalidFileException

from dashboard_package import rebind_dashboard_workbook
from risk_presentation import DATA_SHEET, snapshot_address
from risk_engine import __version__
import investment_results
from risk_engine.scenario import (
    ASSET_ROWS, CACHED_INPUT_ROWS, CORRELATION_HEADER_ROW, CORRELATION_ROWS, HSA_RETURN_ROWS,
    ONE_TIME_EXPENSE_ROWS, RISK_CONTROL_ROWS, RISK_RETURN_BASIS_ROW, RISK_STRESS_ROWS,
    SCENARIO_INPUT_ROWS, TAX_TABLE_LAST_ROW, asset_profile_start,
)

SHEET = DATA_SHEET
SCHEMA = 'retirement-risk-workbook-v4'
STATUS = 'Z1'
SCHEMA_CELL = 'Z2'
SNAPSHOT_CELL = 'Z3'
ENGINE_CELL = 'Z4'
FIRST_ANNUAL_ROW = 46
LAST_ANNUAL_ROW = 171
_X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
_RELATIONSHIP_TYPE = _R[1:-1] + '/worksheet'
_FUNDING_KEYS = (
    'unpaid_spending_real', 'unpaid_tax_real', 'required_spending_real', 'tax_due_real',
    'income_available_real', 'income_used_real', 'hsa_used_real', 'hsa_eligible_expense_real',
    'brokerage_available_real', 'brokerage_draw_real', 'traditional_available_real',
    'traditional_draw_real', 'roth_available_real', 'roth_draw_real',
)
_FUNDING_COLUMNS = (
    'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW',
)


def _manifest():
    """Source references used by both the live Excel comparison and the snapshot writer."""
    yield 'Active scenario', 'Inputs!$B$4', False
    yield 'Wage growth definition', 'Inputs!$A$13', False
    for row in SCENARIO_INPUT_ROWS:
        yield (f'Selected scenario, row {row}',
               f'INDEX(Inputs!$D{row}:$F{row},1,Inputs!$B$4)', False)
    for row in (*SCENARIO_INPUT_ROWS, *CACHED_INPUT_ROWS, 84, *range(130, 134), *HSA_RETURN_ROWS):
        if row in investment_results.CELLS:
            yield investment_results.LABELS[row], investment_results.reference(row), row <= 129
        else:
            yield f'Effective input B{row}', f'Inputs!$B${row}', row <= 129 and row != 84
    for row in (*ONE_TIME_EXPENSE_ROWS, ONE_TIME_EXPENSE_ROWS.stop):
        for col in 'BC':
            yield f'One-time expense {col}{row}', f'Inputs!${col}${row}', False

    for rows, columns in ((ASSET_ROWS, 'ABCDE'),
                          ((CORRELATION_HEADER_ROW, *CORRELATION_ROWS), 'ABCDEF')):
        for row in rows:
            for col in columns:
                yield f'Asset_Allocation!{col}{row}', f'Asset_Allocation!${col}${row}', False
    first_profile = asset_profile_start(1)
    last_profile = asset_profile_start(3) + len(ASSET_ROWS) - 1
    stride = asset_profile_start(2) - first_profile
    for input_row, stage in ((15, 'Working'), (16, 'Retirement')):
        for offset in range(len(ASSET_ROWS)):
            for col in 'ABCDE':
                reference = (f'INDEX(Asset_Allocation!${col}${first_profile}:'
                             f'${col}${last_profile},'
                             f'{offset + 1}+{stride}*(Inputs!$B${input_row}-1),1)')
                yield f'{stage} investment mix, asset {offset + 1}, column {col}', reference, False

    for row in (*RISK_CONTROL_ROWS, RISK_RETURN_BASIS_ROW, *RISK_STRESS_ROWS):
        yield f'Risk_Settings!B{row}', f'Risk_Settings!$B${row}', False
    for row in range(1, TAX_TABLE_LAST_ROW + 1):
        for col in 'ABCDEFGH':
            yield f'Tax_Tables!{col}{row}', f'Tax_Tables!${col}${row}', False


_MANIFEST = tuple(_manifest())
_LEGACY_RATE_LABELS = {label: f'Effective input B{row}' for row, label in investment_results.LABELS.items()}
_FIRST_CHECK_ROW = 2
_LAST_CHECK_ROW = _FIRST_CHECK_ROW + len(_MANIFEST) - 1


def configure_freshness(put, *, is_tax_formula=None):
    """Author the visible snapshot audit and the status used by Summary formulas."""
    for address, text, style in (
        ('A1', 'Saved market balances', 'detail_title'),
        ('A2', 'Balances use opening-year dollars from the saved run. See Risk Results for the active dollar view.', 'report_caption'),
        ('A3', 'Back to risk results', 'report_link'),
        ('D3', 'Open early-loss tests', 'report_link'),
        ('F3', 'Open saved inputs', 'report_link'),
        ('H3', 'Open funding examples', 'report_link'),
        ('K1', 'Saved early-loss tests', 'detail_title'),
        ('AF1', 'Saved funding examples', 'detail_title'),
        ('BA1', 'Run field', 'header'),
        ('BB1', 'Saved value', 'header'),
    ):
        put(SHEET, address, text, style)
    for columns, labels in (
        ('ABCDEFGHI', ('Year', 'Age', 'Base balance', 'Fixed 10th percentile', 'Fixed median',
                       'Fixed 90th percentile', 'Flexible 10th percentile', 'Flexible median',
                       'Flexible 90th percentile')),
        ('KLMNO', ('Stock return', 'Fixed shortfall age', 'Fixed unpaid costs',
                   'Flexible shortfall age', 'Flexible unpaid costs')),
        (('AF', 'AG', 'AH', 'AI', *_FUNDING_COLUMNS, 'AX'),
         ('Path', 'Ending balance', 'Shortfall age', 'Total unpaid',
          *(key.removesuffix('_real').replace('_', ' ').capitalize() for key in _FUNDING_KEYS), 'Trial')),
    ):
        for column, label in zip(columns, labels, strict=True):
            put(SHEET, f'{column}4', label, 'header')
    for columns, first, last in (
        (dict.fromkeys('ABCDEFGHI', 'report_money') | {'A': 'whole', 'B': 'whole'}, 5, 130),
        ({'K': 'report_percent', 'L': 'whole', 'M': 'report_money', 'N': 'whole', 'O': 'report_money'}, 5, 17),
        (dict.fromkeys(('AG', 'AI', *_FUNDING_COLUMNS), 'report_money') | {'AH': 'whole', 'AX': 'whole'}, 5, 7),
    ):
        for row in range(first, last + 1):
            for column, style in columns.items():
                put(SHEET, f'{column}{row}', '', style)
    for row, label in enumerate((
        'Run date', 'Scenario', 'End age', 'Trials', 'Seed', 'Engine version',
        'Fixed success rate', 'Flexible success rate', 'Runs with cuts', 'Median years with cuts',
        'Spending cut trigger', 'Cut of flexible budget', 'Flexible share', 'Opening year',
        'Market variation multiplier', 'Early-loss years', 'Median first shortfall age',
        'Source checksum', 'Input checksum',
    ), 2):
        put(SHEET, f'BA{row}', label, 'report_caption')
    put(SHEET, SCHEMA_CELL, SCHEMA)
    put(SHEET, ENGINE_CELL, __version__)
    for column, title in zip(('AA', 'AB', 'AC', 'AD'),
                             ('Assumption checked', 'Current value', 'Value used in run', 'Matches')):
        put(SHEET, f'{column}1', title, 'report_table_head')
    for row, (label, ref, derived) in enumerate(_MANIFEST, _FIRST_CHECK_ROW):
        current, saved = f'AB{row}', f'AC{row}'
        if is_tax_formula is not None and label.startswith('Tax_Tables!'):
            derived = is_tax_formula(label.split('!', 1)[1])
        numeric = (f'ABS({current}-{saved})<=MAX(0.0000001,'
                   f'0.000000000002*MAX(ABS({current}),ABS({saved})))'
                   if derived else f'{current}={saved}')
        put(SHEET, f'AA{row}', label)
        put(SHEET, current, f'=IF(ISBLANK({ref}),"",{ref})')
        put(SHEET, saved, '')
        put(SHEET, f'AD{row}', f'=IFERROR(IF({current}="",--({saved}=""),'
            f'IF(ISNUMBER({current}),--AND(ISNUMBER({saved}),{numeric}),'
            f'--AND(ISTEXT({current}),ISTEXT({saved}),EXACT({current},{saved})))),0)')
    valid = (f'AND($Z$2=$Z$3,$Z$4=$BB$7,Report_Data!$Q$2>=2,'
             f'COUNTIF($AD$2:$AD${_LAST_CHECK_ROW},1)={len(_MANIFEST)},'
             f'ISNUMBER($BB$5),$BB$5>=1,COUNT($BB$8:$BB$9)=2)')
    put(SHEET, STATUS,
        f'=IF($Z$3="","missing",IF(IFERROR({valid},FALSE),"current","stale"))')


def _worksheet(source, *, calculated=False):
    try:
        book = load_workbook(source, read_only=True, data_only=calculated, keep_links=False)
    except (BadZipFile, InvalidFileException, KeyError, E.XMLSyntaxError):
        raise ValueError('The risk workbook could not be read.') from None
    try:
        if SHEET not in book.sheetnames:
            raise ValueError('Use a workbook release with a Risk_Data sheet.')
        sheet = book[SHEET]
        version = sheet[SCHEMA_CELL].value
        if version != SCHEMA or sheet[ENGINE_CELL].value != __version__:
            raise ValueError('Use the workbook and risk runner from the same release.')
        return book, sheet
    except BaseException:
        book.close()
        raise


def validate_source(source):
    """Reject older files and incompatible versions before doing an expensive run."""
    book, sheet = _worksheet(source)
    try:
        if sheet[STATUS].data_type != 'f':
            raise ValueError('The workbook is missing its risk report status formula.')
    finally:
        book.close()


def _assumption_values(calculated_source):
    book, sheet = _worksheet(calculated_source, calculated=True)
    try:
        values = {}
        cells = sheet.iter_rows(min_row=_FIRST_CHECK_ROW, max_row=_LAST_CHECK_ROW,
                                min_col=27, max_col=28)
        for row, ((label, _, _), (name, value)) in enumerate(zip(_MANIFEST, cells, strict=True), _FIRST_CHECK_ROW):
            if name.value not in (label, _LEGACY_RATE_LABELS.get(label, label)) or value.data_type == 'e':
                raise ValueError('Risk assumptions could not be read from the calculated workbook.')
            values[f'AC{row}'] = value.value
        return values
    finally:
        book.close()


def runtime_addresses():
    """List the snapshot fields that are replaced on every refresh."""
    yield SNAPSHOT_CELL
    for row in range(_FIRST_CHECK_ROW, _LAST_CHECK_ROW + 1):
        yield f'AC{row}'
    for row in range(2, 21):
        yield f'AF{row}'
    for row in range(21, 24):
        for col in ('AF', 'AG', 'AH', 'AI', *_FUNDING_COLUMNS, 'AX'):
            yield f'{col}{row}'
    for row in range(28, 41):
        for col in ('AF', 'AG', 'AH', 'AI', 'AJ'):
            yield f'{col}{row}'
    for row in range(FIRST_ANNUAL_ROW, LAST_ANNUAL_ROW + 1):
        for col in ('AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN'):
            yield f'{col}{row}'


def _report_values(report, scenario_name, run_date):
    provenance, timeline, controls = (report[key] for key in ('provenance', 'timeline', 'controls'))
    fixed, flexible, cuts = (report[key] for key in ('fixed', 'flexible', 'spending_cuts'))
    values = dict.fromkeys(runtime_addresses())
    values[SNAPSHOT_CELL] = SCHEMA
    fields = (run_date.isoformat(), scenario_name, timeline['end_age'], provenance['trials'],
              provenance['seed'], provenance['engine_version'], fixed['success_rate'],
              flexible['success_rate'], cuts['trials_with_cuts'], cuts['median_years_when_cut'],
              controls['trigger'], controls['cut'], controls['flexible_share'],
              timeline['opening_year'], provenance['volatility_multiplier'],
              controls['stress_years_applied'], fixed['median_failure_age'],
              provenance['workbook_sha256'], provenance['inputs_sha256'])
    for row, value in enumerate(fields, 2):
        values[f'AF{row}'] = value
    for row, path in zip(range(21, 24), report['representative_paths'], strict=True):
        policy = path['fixed']
        funding = policy['first_shortfall'] or {}
        cells = ('AF', 'AG', 'AH', 'AI', *_FUNDING_COLUMNS, 'AX')
        details = (path['label'], policy['ending_balance_real'], policy['first_shortfall_age'],
                   funding.get('total_real'), *(funding.get(key) for key in _FUNDING_KEYS), path['trial'] + 1)
        for col, detail in zip(cells, details, strict=True):
            values[f'{col}{row}'] = detail
    for row, case in enumerate(report['sequence'], 28):
        if row > 40:
            raise ValueError('Too many early retirement cases for the workbook.')
        fixed, flexible = case['fixed'], case['flexible']
        fields = (case['equity_return'], fixed['first_shortfall_age'],
                  (fixed['first_shortfall'] or {}).get('total_real'), flexible['first_shortfall_age'],
                  (flexible['first_shortfall'] or {}).get('total_real'))
        for col, value in zip(('AF', 'AG', 'AH', 'AI', 'AJ'), fields, strict=True):
            values[f'{col}{row}'] = value

    opening_balance = report['baseline']['opening_balance_real']
    opening = (timeline['opening_year'], timeline['opening_age'], *([opening_balance] * 7))
    annual = [(year['year'], year['age'], year['base'], *year['fixed_p10_p50_p90'],
               *year['flexible_p10_p50_p90']) for year in report['annual']]
    for row, fields in enumerate((opening, *annual), FIRST_ANNUAL_ROW):
        if row > LAST_ANNUAL_ROW:
            raise ValueError('Risk horizon exceeds the workbook layout.')
        columns = ('AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN')
        for col, value in zip(columns, fields, strict=True):
            values[f'{col}{row}'] = value
    return {snapshot_address(address): value for address, value in values.items()}


def _parse(data):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))


def _sheet_part(archive):
    book = _parse(archive.read('xl/workbook.xml'))
    rels = _parse(archive.read('xl/_rels/workbook.xml.rels'))
    sheet = next(item for item in book.find(_X + 'sheets') if item.get('name') == SHEET)
    identifier = sheet.get(_R + 'id')
    rel = next(item for item in rels if item.get('Id') == identifier)
    if rel.get('Type') != _RELATIONSHIP_TYPE or rel.get('TargetMode') == 'External':
        raise ValueError('Invalid risk worksheet package relationship.')
    target = rel.get('Target')
    return normpath(target.lstrip('/') if target.startswith('/') else join('xl', target))


def _set_cell(node, data):
    if node.find(_X + 'f') is not None:
        raise ValueError('Risk results overlap a workbook formula.')
    for tag in ('is', 'v'):
        child = node.find(_X + tag)
        if child is not None:
            node.remove(child)
    node.attrib.pop('t', None)
    if data in (None, ''):
        return
    if type(data) in (int, float) and isfinite(data):
        E.SubElement(node, _X + 'v').text = repr(data)
    elif isinstance(data, str):
        node.set('t', 'inlineStr')
        text = E.SubElement(E.SubElement(node, _X + 'is'), _X + 't')
        if data.strip() != data:
            text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        text.text = data
    else:
        raise ValueError('Risk results contain an unsupported cell value.')


def _replace_snapshot(sheet, values):
    sheet_data = sheet.find(_X + 'sheetData')
    rows = {int(node.get('r')): node for node in sheet_data.findall(_X + 'row')}
    cells = {node.get('r'): node for node in sheet_data.iter(_X + 'c')}
    for address, data in values.items():
        if address in cells:
            _set_cell(cells[address], data)
        elif data not in (None, ''):
            _, row = coordinate_from_string(address)
            if row not in rows:
                rows[row] = E.SubElement(sheet_data, _X + 'row', r=str(row))
            node = E.SubElement(rows[row], _X + 'c', r=address)
            cells[address] = node
            _set_cell(node, data)
            # Recalculate row spans for writers that originally saved this sheet in Excel.
            rows[row].attrib.pop('spans', None)
    for row in rows.values():
        ordered = sorted(row.findall(_X + 'c'),
                         key=lambda cell: column_index_from_string(coordinate_from_string(cell.get('r'))[0]))
        row[:] = ordered + [element for element in row if element.tag != _X + 'c']
    sheet_data[:] = [rows[row] for row in sorted(rows)]
    dimension = sheet.find(_X + 'dimension')
    if dimension is not None:
        furthest_col = max((coordinate_from_string(address)[0] for address in cells),
                           key=column_index_from_string)
        dimension.set('ref', f'A1:{furthest_col}{max(rows)}')


def write_snapshot(source, destination, calculated_source, scenario, report, *, workbook_name=None, run_date=None):
    """Update the risk sheet and local chart names in an unpublished copy of the source."""
    values = _report_values(report, scenario.inputs[7], run_date or date.today())
    values.update(_assumption_values(calculated_source))
    try:
        with ZipFile(source) as original:
            part = _sheet_part(original)
            sheet = _parse(original.read(part))
            _replace_snapshot(sheet, values)
            risk_xml = E.tostring(sheet, encoding='UTF-8', xml_declaration=True, standalone=True)
            with ZipFile(destination, 'x') as output:
                for entry in original.infolist():
                    payload = risk_xml if entry.filename == part else original.read(entry)
                    if workbook_name and re.fullmatch(r'xl/charts/chart\d+\.xml', entry.filename):
                        chart = _parse(payload)
                        if rebind_dashboard_workbook(chart, workbook_name):
                            payload = E.tostring(chart, encoding='UTF-8', xml_declaration=True, standalone=True)
                    output.writestr(entry, payload, compresslevel=6)
    except (BadZipFile, KeyError, StopIteration, E.XMLSyntaxError) as exc:
        raise ValueError('Could not write the risk snapshot into the workbook package.') from exc


def verify_snapshot(calculated_source, report):
    """Require fresh workbook formulas to accept the new snapshot before publishing it."""
    book, sheet = _worksheet(calculated_source, calculated=True)
    try:
        if sheet[STATUS].value != 'current':
            matches = sheet.iter_rows(min_row=_FIRST_CHECK_ROW, max_row=_LAST_CHECK_ROW,
                                      min_col=30, max_col=30, values_only=True)
            for (label, _, _), (match,) in zip(_MANIFEST, matches):
                if match != 1:
                    raise ValueError(f'The risk assumption check disagreed at {label}.')
            raise ValueError('The calculated workbook did not validate its new risk results.')
        if sheet['BB5'].value != report['provenance']['trials']:
            raise ValueError('The calculated workbook contains an unexpected risk trial count.')
        if (sheet['BB19'].value != report['provenance']['workbook_sha256']
                or sheet['BB20'].value != report['provenance']['inputs_sha256']):
            raise ValueError('The calculated workbook contains unexpected risk run provenance.')
    finally:
        book.close()
