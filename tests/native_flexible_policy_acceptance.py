"""Recalculate the release workbook's flexible-spending decision formulas at their boundaries."""

import argparse
from math import isclose
from pathlib import Path
from tempfile import TemporaryDirectory
import sys

from openpyxl import Workbook as NewWorkbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from workbook_recalc import recalculate
from xlsx_read import Workbook


# Stage, fixed opening balance, base opening balance, year, filing, living costs,
# expected opening-balance threshold, expected flexible cut.
CASES = (
    ('Working', 0, 100, 2029, 'MFJ', 2000, 0, 0),
    ('Retirement', 75, 100, 2029, 'MFJ', 2000, 75, 0),
    ('Retirement', 74.99, 100, 2029, 'MFJ', 2000, 75, 80),
    ('Retirement', 76, 100, 2029, 'MFJ', 2000, 75, 0),
    ('Retirement', 50, 100, 2030, 'MFJ', 2000, 75, 140),
    ('Retirement', 50, 100, 2030, 'Single', 1700, 75, 140),
    ('Retirement', 0, 100, 2031, 'Single', 500, 75, 50),
    ('Retirement', 50, 100, 2029, 'MFJ', 600, 75, 0),
    ('Retirement', 0, 0, 2030, 'MFJ', 2000, 0, 0),
    ('Retirement', '=1/0', 100, 2031, 'MFJ', 2000, 75, '#DIV/0!'),
    ('Retirement', 75, '=1/0', 2031, 'MFJ', 2000, '#DIV/0!', '#DIV/0!'),
)


def check(path, office):
    model = NewWorkbook()
    model.remove(model.active)
    for name in ('Trial_Inputs', 'Risk_Settings', 'Projection', 'Trial_Projection', 'Fixed', 'Flexible'):
        model.create_sheet(name)
    inputs, settings, base, trial, fixed, flexible = model.worksheets
    inputs['B38'], inputs['B39'], inputs['B40'], inputs['B56'] = 1200, 2030, 6, .5
    settings['B6'], settings['B7'], settings['B8'] = .75, .25, .40
    for policy in (fixed, flexible):
        policy['D3'], policy['H3'], policy['J3'] = len(CASES) + 1, policy is flexible, 'OK'
    with Workbook(path) as release:
        source = release['Market_Paths']
        for index, values in enumerate(CASES, 7):
            stage, opening, planned, year, filing, living, _, _ = values
            trial[f'BH{index-2}'], base[f'BH{index-2}'] = opening, planned
            trial[f'B{index-1}'], trial[f'F{index-1}'], trial[f'CH{index-1}'] = year, filing, living
            for policy in (fixed, flexible):
                policy[f'D{index}'] = stage
        for index in range(6, 8 + len(CASES)):
            for col in ('AE', 'AH', 'AI'):
                address = f'{col}{index}'
                formula = source[address].formula
                if not formula:
                    raise AssertionError(f'Market_Paths!{address} has no formula.')
                for policy in (fixed, flexible):
                    policy[address] = '=' + formula
    with TemporaryDirectory() as temp:
        root = Path(temp)
        original = root / 'policy.xlsx'
        model.save(original)
        output = root / 'calculated'
        output.mkdir()
        result = recalculate(original, output, root / 'profile', office, 90)
        with Workbook(result) as saved:
            for name in ('Fixed', 'Flexible'):
                sheet = saved[name]
                for index, case in enumerate(CASES, 7):
                    threshold, cut = case[-2:]
                    for col, expected in (('AI', threshold), ('AH', cut), ('AE', cut if name == 'Flexible' else 0)):
                        actual = sheet[f'{col}{index}'].value
                        if type(expected) in (float, int):
                            ok = type(actual) in (int, float) and isclose(actual, expected, abs_tol=1e-10)
                        else:
                            ok = actual == expected
                        if not ok:
                            raise AssertionError(f'{name}!{col}{index}: expected {expected!r}; got {actual!r}')
                for col in ('AE', 'AH', 'AI'):
                    if sheet[f'{col}6'].value != 0 or sheet[f'{col}{7 + len(CASES)}'].value not in ('', None):
                        raise AssertionError(f'{name}!{col} must be zero in the opening year and blank after the end.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--soffice', default='soffice')
    args = parser.parse_args()
    check(args.workbook, args.soffice)
    print(f'PASS {len(CASES)} annual boundary cases for both fixed and flexible worksheet formulas.')
