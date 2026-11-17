"""Recalculate the production input checks with valid, missing and pasted values."""

import argparse
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
        result.close()
    print(f'PASS {len(cases)} calculated input-validation cases')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--soffice', default='soffice')
    check(parser.parse_args().soffice)
