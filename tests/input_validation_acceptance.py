"""Recalculate the production input checks with valid, missing and pasted values."""

import argparse
from collections import defaultdict
from pathlib import Path
import subprocess
import sys
import tempfile

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))
from example_inputs import example_cells
import model_plan
from risk_engine.scenario import SCENARIO_INPUT_ROWS
import tax_policy


def add_state_policy_cases(book: Workbook, horizon: str) -> dict[str, tuple[str, str]]:
    """Use production formulas to exercise pasted values and active scenario changes."""
    edits = defaultdict(dict)

    def put(sheet: str, cell: str, value: str, style: str) -> None:
        edits[sheet][cell] = value

    tax_policy.configure_workbook_state_policy(put)
    invalid = 'Check state tax jurisdiction'
    cases = (
        ('California', {}, {}, 'OK'),
        ('Blank jurisdiction', {'D136': None}, {}, invalid),
        ('Lowercase jurisdiction', {'D136': 'ca'}, {}, invalid),
        ('Unsupported jurisdiction', {'D136': 'FL'}, {}, invalid),
        ('Pasted jurisdiction formula', {'D136': '="CA"'}, {}, invalid),
        ('Inactive unsupported jurisdiction', {'E136': 'FL'}, {}, 'OK'),
        ('Select unsupported jurisdiction', {'B4': 2, 'E136': 'FL'}, {}, invalid),
        ('Select third California scenario', {'B4': 3, 'D136': 'FL'}, {}, 'OK'),
        ('Replaced selector formula', {'B136': 'CA'}, {}, invalid),
        ('Missing workbook schema', {}, {'B266': None}, invalid),
        ('Wrong policy version case', {}, {'B267': 'CA-INDIVIDUAL-V1'}, invalid),
        ('Pasted policy version formula', {}, {'B267': '="ca-individual-v1"'}, invalid),
        ('Invalid opening basis', {'B51': -1}, {}, 'Check tax inputs'),
    )
    valid_inputs = {'B4': 1, 'B8': 65, 'B11': 1, 'B25': 100, 'B51': 0, 'B53': 67, 'B70': horizon,
                    **{f'B{row}': 'OK' for row in (71, 84, 99, 125, 133)},
                    **{f'{column}{row}': value for column in 'BDEF'
                       for row, value in ((126, .5), (127, 0), (128, 0))}}
    checks = {}
    for index, (name, inputs, metadata, status) in enumerate(cases):
        sheet = book.create_sheet(f'State{index}')
        tax = book.create_sheet(f'Tax{index}')
        for cell, value in (edits['Tax_Tables'] | metadata).items():
            tax[cell] = value
        for cell, value in (valid_inputs | edits['Inputs'] | inputs).items():
            sheet[cell] = value.replace('Tax_Tables!', f'{tax.title}!') if isinstance(value, str) else value
        checks[sheet.title] = (name, status)
    return checks


def check(soffice):
    template = load_workbook(ROOT / 'workbook/build-template.xlsx', read_only=True)
    model_plan.inputs(template)
    edits = model_plan.cells['Inputs']
    sample = example_cells()
    book = Workbook()
    book.remove(book.active)
    cases = (
        ('Unchanged inputs', {}, 'OK', 'OK'),
        ('Zero wages and spending', {'D12': 0, 'D36': 0, 'D37': 0}, 'OK', 'OK'),
        ('Missing current spending', {'D36': None}, 'Enter annual spending now on Inputs', 'OK'),
        ('Retired with no working inputs', {'D8':54,'D10':55,'D53':57,'D12':None,'D36':None}, 'OK', 'OK'),
        ('Working with no retirement spending', {'D8':54,'D10':65,'D53':57,'D37':None}, 'OK', 'OK'),
        ('Missing active pay', {'D12':None}, 'Enter annual pay on Inputs', 'OK'),
        ('Missing active retirement spending', {'D8':54,'D10':55,'D53':57,'D37':None}, 'Enter retirement spending on Inputs', 'OK'),
        ('Inactive phase with pasted text', {'D8':54,'D10':65,'D53':57,'D37':'unknown'}, 'Check for text or a negative amount', 'OK'),
        ('Inactive phase with negative spending', {'D8':54,'D10':55,'D53':57,'D36':-1}, 'Check for text or a negative amount', 'OK'),
        ('Text opening balance', {'D25': 'amount'}, 'Check for text or a negative amount', 'OK'),
        ('Negative wages', {'D12': -1}, 'Check for text or a negative amount', 'OK'),
        ('Inactive bad scenario', {'E12': -1}, 'OK', 'OK'),
        ('Select invalid scenario', {'E12': -1, 'B4': 2}, 'Check for text or a negative amount', 'OK'),
        ('Invalid pasted Roth type', {'D28': 3}, 'Check the number of earners', 'OK'),
        ('Horizon sentinel', {'D58': 999}, 'OK', 'OK'),
        ('Unsupported age', {'D54': 150}, 'Check the age for survivor', 'OK'),
        ('Return can fall', {'D13': -1, 'D14': -.10}, 'OK', 'OK'),
        ('Inflation must exceed -100%', {'D14': -1}, 'Check Inputs row 14', 'OK'),
        ('Missing shared bill year', {'C78': 150}, 'OK', 'Check the year and cost on Inputs row 78'),
        ('Negative shared bill', {'B80': 2040, 'C80': -1}, 'OK', 'Check the year and cost on Inputs row 80'),
        ('Zero shared bill', {'C81': 0}, 'OK', 'OK'),
        ('Valid shared bill', {'B77': 2041, 'C77': 5500}, 'OK', 'OK'),
        ('Bill in retired row', {'C83': 5500}, 'OK', 'Move one-time expenses'),
        ('Pasted asset return floor', {'asset': ('B6', -1)}, 'Check asset returns', 'OK'),
        ('Pasted asset text', {'asset': ('C8', 'rate')}, 'Check asset returns', 'OK'),
        ('Pasted negative yield', {'asset': ('D9', -.01)}, 'Check asset returns', 'OK'),
    )
    assets = book.create_sheet('Asset_Allocation')
    for row in template['Asset_Allocation'].iter_rows(min_row=6, max_row=10, min_col=2, max_col=5):
        for cell in row:
            assets[cell.coordinate] = cell.value
    assets['B94'] = assets['B95'] = 'OK'
    for i, (_, values, _, _) in enumerate(cases):
        sheet = book.create_sheet(f'Case{i}')
        assumptions = dict(values)
        invalid_asset = assumptions.pop('asset', None)
        asset_sheet = assets
        if invalid_asset is not None:
            asset_sheet = book.copy_worksheet(assets)
            asset_sheet.title = f'Assets{i}'
            asset_sheet[invalid_asset[0]] = invalid_asset[1]
        for address, value in (sample | assumptions).items():
            sheet[address] = value
        for row in SCENARIO_INPUT_ROWS:
            address = f'B{row}'
            sheet[address] = edits.get(address) or template['Inputs'][address].value
        sheet['B71'] = edits['B71'].replace('Asset_Allocation!', asset_sheet.title + '!')
        sheet['B84'] = edits['B84']
    template.close()
    state_cases = add_state_policy_cases(book, edits['B70'])

    # Use a dedicated profile so an unrelated office session cannot intercept the run.
    with tempfile.TemporaryDirectory() as directory:
        folder = Path(directory)
        output = folder / 'calculated'
        output.mkdir()
        book.save(folder / 'cases.xlsx')
        subprocess.run([soffice, f'-env:UserInstallation={(folder / "profile").as_uri()}', '--headless',
                        '--convert-to', 'xlsx', '--outdir', str(output), str(folder / 'cases.xlsx')],
                       check=True, capture_output=True, text=True, timeout=90)
        result = load_workbook(output / 'cases.xlsx', read_only=True, data_only=True)
        for index, (name, _, plan_status, expense_status) in enumerate(cases):
            sheet = result[f'Case{index}']
            for address, expected in [('B71',plan_status), ('B84',expense_status)]:
                actual = sheet[address].value
                if not isinstance(actual, str) or not actual.startswith(expected):
                    raise AssertionError((name, address, actual, expected))
        for sheet, (name, expected) in state_cases.items():
            status, horizon = (result[sheet][cell].value for cell in ('B129', 'B70'))
            if (status, horizon) != (expected, 3 if expected == 'OK' else 0):
                raise AssertionError((name, status, horizon, expected))
        result.close()
    print(f'PASS {len(cases)} calculated input-validation cases and {len(state_cases)} state-policy cases')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--soffice', default='soffice')
    check(parser.parse_args().soffice)
