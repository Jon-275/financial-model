"""Check scheduled bills from household entry through both market policies and the saved report."""
import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from dashboard_acceptance import validate_fixed_headline
from fixtures import COMMON
from native_bootstrap import bootstrap
import native_capture
from native_capture_acceptance import compare_rows, prepare_trial_copy
from native_model_acceptance import validate_trial
from native_random import normal_draws
from workbook_recalc import recalculate
from xlsx_acceptance import prepare_xlsx
from xlsx_read import Workbook
from xlsx_values import write_values


def close(actual, expected, label):
    if type(actual) not in (int, float) or abs(actual - expected) > 0.000001:
        raise AssertionError(f'{label}: expected {expected}, found {actual}')


def calculated(source, folder, changes, office):
    folder.mkdir()
    output = folder / 'out'
    output.mkdir()
    fixture = folder / 'case.xlsx'
    prepare_xlsx(source, fixture, changes, compact=True)
    return recalculate(fixture, output, folder / 'profile', office, 600)


def fixture_settings():
    # Two bills share the first year, an explicit zero follows, the last entry
    # falls in a survivor year, and the age-58 base plan ends before the 2033 bill.
    inputs = {**COMMON['Inputs'], 'D53': 58, 'D54': 55, 'D14': .025, 'D15': 1, 'D16': 1,
              'D25': 90000, 'D51': 90000, 'D37': 7500, 'D56': .8, 'D101': 0,
              'B75': 2027, 'C75': 600, 'B77': 2027, 'C77': 1400, 'B79': 2029, 'C79': 0,
              'B81': 2033, 'C81': 8700, 'B82': 2028, 'C82': 3000}
    assets = {**COMMON['Asset_Allocation'], 'C6': .40,
              **{f'{col}{row}': int(row == 26) for row in range(26, 31) for col in 'BCDE'}}
    return {'Inputs': inputs, 'Asset_Allocation': assets,
            'Risk_Settings': {'B6': 1, 'B7': .4, 'B8': .5, 'B29': 1, 'B30': 3, 'B31': 1,
                              'B32': 1, 'B33': .03, 'B34': 0, 'B35': -.05, 'B36': .12},
            'Risk_Data': {f'Z{row}': None for row in range(6, 12)}}


def check_trial(path, saved, code):
    # Independently reconstruct the price path, including its lower and upper bounds.
    shocks = list(normal_draws(3, 1, 4))
    trial_price = 1
    with Workbook(path) as book, Workbook(saved) as stored:
        capture, original = book[native_capture.SHEET], stored[native_capture.SHEET]
        record_row = native_capture.FIRST + int(code == 1001)
        compare_rows(capture, original, 7, record_row, native_capture.SIGNATURE_COLUMN,
                     code, refresh_date=True)
        for row, dollars in enumerate((2000, 3000, 0, 0), start=6):
            elapsed = row - 5
            trial_price *= 1 + min(.12, max(-.05, .025 + .03 * shocks[elapsed - 1][-1]))
            for sheet, price in (('Projection', 1.025 ** elapsed), ('Trial_Projection', trial_price)):
                table = book[sheet]
                close(table[f'G{row}'].value, price, f'{sheet} price index, year {elapsed}')
                close(table[f'AX{row}'].value, dollars * price, f'{sheet} nominal bill, year {elapsed}')
                if table[f'F{row}'].value != ('MFJ' if elapsed == 1 else 'Single'):
                    raise AssertionError('The fixture must include the transition to survivor spending.')
        if code == 1001 and not book['Market_Paths']['AE8'].value > 0:
            raise AssertionError('The flexible case must cut living costs in the year of the survivor bill.')
        close(book['Trial_Inputs']['C79'].value, 0, 'Explicit zero expense')
        if book['Trial_Inputs']['C76'].value not in (None, ''):
            raise AssertionError('An unused amount must stay blank.')
    validate_trial(path, saved)


def check_funding_change(entered, saved, root, office):
    revision = {'Inputs': {'C77': 100000}}
    changed = calculated(saved, root / 'edited-bill', revision, office)
    with Workbook(changed) as book:
        if book['Risk_Data']['Z1'].value != 'stale' or 'update' not in book['Summary']['K22'].value.lower():
            raise AssertionError('Changing only a bill amount must mark the saved Summary for refresh.')
        if book['Summary']['F12'].value not in (None, ''):
            raise AssertionError('The old fixed-spending probability is still visible.')
        close(book['Projection']['AX6'].value, 100600 * 1.025, 'Revised first-year bills')
    revised, refreshed = root / 'revised.xlsx', root / 'refreshed.xlsx'
    write_values(entered, revised, revision)
    bootstrap(revised, refreshed, office, 600)
    with Workbook(saved) as before, Workbook(refreshed) as after:
        for book, probability in ((before, 1), (after, 0)):
            if book['Risk_Data']['Z1'].value != 'current':
                raise AssertionError('Both independently populated reports must show current results.')
            validate_fixed_headline(book['Summary'], probability, book['Risk_Data']['BB5'].value)
            for row in (8, 9):
                close(book['Risk_Trials'][f'D{row}'].value, probability, 'Fixed or flexible stored success')
        if not after['Summary']['B8'].value.startswith('First shortfall at age 55'):
            raise AssertionError('The Summary must show the age of the first unfunded bill.')
        for row in (8, 9):
            close(after['Risk_Trials'][f'E{row}'].value, 55, 'First unfunded year in the market trial')


def verify(source, root, office):
    clean, entered, saved = (root / name for name in ('clean.xlsx', 'entered.xlsx', 'saved.xlsx'))
    prepare_trial_copy(source, clean)
    prepare_xlsx(clean, entered, fixture_settings())
    bootstrap(entered, saved, office, 600)
    with Workbook(saved) as book:
        if book['Risk_Data']['Z1'].value != 'current':
            raise AssertionError('The saved run must be current before testing an amount change.')
    for code in (1, 1001):
        trial = calculated(entered, root / str(code),
                           {'Risk_Trials': {'B3': code}, 'Risk_Data': {'Z6': 'running'}}, office)
        check_trial(trial, saved, code)
    check_funding_change(entered, saved, root, office)
    print('PASS scheduled bills in base, fixed and flexible years; a larger bill changes the report’s funding result.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--soffice', default='soffice')
    options = parser.parse_args()
    with TemporaryDirectory() as directory:
        verify(options.workbook, Path(directory), options.soffice)
