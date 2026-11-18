"""Check which spending and pay inputs the base plan and extended market tests require."""
import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from openpyxl.worksheet.cell_range import MultiCellRange

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from fixtures import COMMON
from native_bootstrap import _read_plan
from native_capture_acceptance import prepare_trial_copy
from risk_engine.scenario import ScenarioError, load_scenario
from workbook_recalc import recalculate
from xlsx_acceptance import prepare_xlsx
from xlsx_read import Workbook, q, sheet_parts


LONGER = 'Enter retirement spending on Inputs for the longer market tests.'
BASE_REQUIRED = 'Enter retirement spending on Inputs'
EXTEND = 'Extend or correct the plan'
# Opening age 54; the three projected base years cover ages 55–57.
# Blank cells are intentional. Zero is a valid selected budget.
CASES = (
    ('retired', 55, (12, 36), 'OK', 'OK'),
    ('explicit zero retirement budget', 65, (), 'OK', 'OK'),
    ('retirement at 100', 100, (37,), 'OK', LONGER),
    ('retirement at 105', 105, (37,), 'OK', LONGER),
    ('retirement at 106', 106, (37,), 'OK', 'OK'),
    ('retirement during base plan', 56, (37,), BASE_REQUIRED, EXTEND),
)


def equal(label, actual, expected):
    if actual != expected:
        raise AssertionError(f'{label}: expected {expected!r}, found {actual!r}')


def rejected(action, path, error_type, expected):
    try:
        action(path)
    except error_type as error:
        if expected not in str(error):
            raise AssertionError(f'{path.name}: expected {expected!r}, found {error}') from error
    else:
        raise AssertionError(f'{path.name}: accepted invalid input, expected {expected!r}')


def replace_input(source, destination, address, value, *, cache=False):
    """Edit one source literal or its selected-value cache while preserving all other saved results."""
    with ZipFile(source) as original, ZipFile(destination, 'x', ZIP_DEFLATED) as output:
        inputs = sheet_parts(original)['Inputs']
        for entry in original.infolist():
            data = original.read(entry)
            if entry.filename == inputs:
                root = etree.fromstring(data)
                cell = root.find(f".//{q('c')}[@r='{address}']")
                if cell is None or (cell.find(q('f')) is not None) != cache:
                    raise AssertionError(f'Expected {"cached formula" if cache else "literal"} in Inputs!{address}')
                for child in list(cell):
                    if child.tag != q('f'):
                        cell.remove(child)
                cell.attrib.pop('t', None)
                if isinstance(value, str):
                    cell.set('t', 'inlineStr')
                    etree.SubElement(etree.SubElement(cell, q('is')), q('t')).text = value
                elif value is not None:
                    etree.SubElement(cell, q('v')).text = str(value)
                data = etree.tostring(root, encoding='UTF-8', xml_declaration=True)
            output.writestr(entry, data)


def check_case(source, folder, case, office):
    label, retirement, blanks, base, market = case
    inputs = {**COMMON['Inputs'], 'D10': retirement, 'D12': 62000, 'D36': 37000, 'D37': 0 if label == 'explicit zero retirement budget' else 40000,
              'D53': 57, **{f'D{row}': None for row in blanks}}
    in_file = folder / ('scenario' + source.suffix)
    # Clear the previous run's metadata as well as its saved trial rows before testing bootstrap.
    changes = {'Inputs': inputs, 'Risk_Settings': {'B29': 1},
               'Risk_Data': {f'Z{row}': None for row in range(6, 12)}}
    prepare_xlsx(source, in_file, changes, compact=True)
    directory = folder / 'calculated'
    directory.mkdir()
    result = recalculate(in_file, directory, folder / 'profile', office, 600)
    with Workbook(result) as book:
        equal('Base input check', book['Inputs']['B71'].value, base)
        equal('Base row count', book['Inputs']['B70'].value, 4 if base == 'OK' else 0)
        equal('Market input check', book['Market_Paths']['AL3'].value, market)
        status = book['Risk_Settings']['A44'].value
        if market != 'OK':
            equal('User-facing market message', status, market)
    if market == 'OK':
        _read_plan(result)
    else:
        rejected(_read_plan, result, ValueError, market)
    if base == 'OK':
        scenario = load_scenario(result)
        for row in (12, 36, 37):
            equal(f'Python selected input {row}', scenario.inputs[row], inputs[f'D{row}'] or 0)
    else:
        rejected(load_scenario, result, ScenarioError, 'Inputs!D37')
    if label == 'retired':
        for cache in (None, 0):
            changed = folder / f'blank-selector-{cache}.xlsx'
            replace_input(result, changed, 'B12', cache, cache=True)
            equal('Blank or zero saved selector', load_scenario(changed).inputs[12], 0)
        for pos, (address, value, saved, message) in enumerate((
            ('D12', '0', False, 'Inputs!D12'), ('D36', 'pasted text', False, 'Inputs!D36'),
            ('B12', 125, True, 'Inputs!B12'), ('B12', 0.00000001, True, 'Inputs!B12'),
            ('D37', None, False, 'Inputs!D37'),
        )):
            changed = folder / f'invalid-{pos}.xlsx'
            replace_input(result, changed, address, value, cache=saved)
            rejected(load_scenario, changed, ScenarioError, message)
    if label == 'retirement at 106':
        for address in ('D12', 'D36'):
            changed = folder / f'active-{address}.xlsx'
            replace_input(result, changed, address, None)
            rejected(load_scenario, changed, ScenarioError, f'Inputs!{address}')
    print(f'PASS {label}: base check {base!r}, market check {market!r}.', flush=True)


def check_entry_rules(workbook):
    # Check all three scenario columns, including adjacent amounts that still require a number.
    expected = {f'{col}{row}': row in (12, 36, 37)
                for col in 'DEF' for row in (12, 25, 35, 36, 37, 38)}
    observed = {}
    with ZipFile(workbook) as archive:
        inputs = etree.fromstring(archive.read(sheet_parts(archive)['Inputs']))
    for rule in inputs.iter(q('dataValidation')):
        areas = MultiCellRange(rule.get('sqref'))
        for address in expected:
            if address in areas:
                if address in observed:
                    raise AssertionError(f'Multiple Excel entry rules at Inputs!{address}')
                equal(f'Inputs!{address} entry amount', (rule.get('type'), rule.get('operator'), rule.findtext(q('formula1'))),
                      ('decimal', 'greaterThanOrEqual', '0'))
                observed[address] = rule.get('allowBlank') in ('1', 'true')
    equal('Excel entry rules for pay and spending', observed, expected)


def verify(workbook, office):
    check_entry_rules(workbook)
    with TemporaryDirectory() as tmp:
        root = Path(tmp)
        cleared = root / ('source' + workbook.suffix)
        prepare_trial_copy(workbook, cleared)
        for index, case in enumerate(CASES):
            directory = root / str(index)
            directory.mkdir()
            check_case(cleared, directory, case, office)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--soffice', default='soffice')
    args = parser.parse_args()
    verify(args.workbook, args.soffice)
