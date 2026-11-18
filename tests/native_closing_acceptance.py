"""Check annual account rollforwards using the formulas in a generated workbook."""
import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

from openpyxl import Workbook as TestBook
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from workbook_recalc import recalculate
from xlsx_read import Workbook


PROJECTION, ROTH = 'Trial_Projection', 'Trial_Roth_Access'
CHECKS, CASH, INPUT, MARKET = 'Trial_Checks', 'Trial_Cash_Flow', 'Trial_Inputs', 'Market_Paths'
TABS = {PROJECTION: 'p', ROTH: 'r', CHECKS: 'c', CASH: 'f', INPUT: 'i', MARKET: 'm'}
# Independent public contract for N:Y. Its order also verifies the description above each field.
CLOSINGS = ((PROJECTION, 'BC'), (PROJECTION, 'BE'), (PROJECTION, 'BG'), (ROTH, 'AI'), (ROTH, 'AJ'),
            (ROTH, 'AK'), (ROTH, 'AL'), (PROJECTION, 'BD'), (ROTH, 'AM'), (PROJECTION, 'BF'),
            (PROJECTION, 'BH'), (PROJECTION, 'BN'))
RESIDUALS = tuple(get_column_letter(index) for index in range(14, 26))
ACTIVE = range(5, 10)


def tab(case, name):
    return f'{case}_{TABS[name]}'


def equal(label, actual, expected):
    if actual != expected:
        raise AssertionError(f'{label}: expected {expected!r}, found {actual!r}')


def add_case(test, source, case, *, ira=True, depleted=False, rates=()):
    """Use the public close and review formulas. Supply literal balances, flows and market rates."""
    for name in TABS:
        test.create_sheet(tab(case, name))
    def get(name):
        return test[tab(case, name)]
    def copy(name, address):
        formula = source[name][address].formula
        if not formula:
            raise AssertionError(f'Missing source formula at {name}!{address}')
        for original in TABS:
            formula = formula.replace(original + '!', tab(case, original) + '!')
        get(name)[address] = '=' + formula

    inputs = get(INPUT)
    inputs['B70'], inputs['B28'], inputs['B90'] = len(ACTIVE), 1 if ira else 2, .25
    for address, value in {'B25': 400, 'B26': 800, 'B27': 1200, 'B31': 160, 'B51': 200}.items():
        inputs[address] = 0 if depleted else value
    for row in (*ACTIVE, 10):
        get(PROJECTION)[f'G{row}'] = 1
        get(CASH)[f'AM{row}'] = 'OK'
        for column in ('AA', 'AB', 'AC', 'AD'):
            get(MARKET)[f'{column}{row+1}'] = dict(rates).get(row, 0)
        for column in 'DEFG':
            copy(ROTH, column + str(row))
        for sheet, col in CLOSINGS:
            copy(sheet, col + str(row))
        for col in ('J', *RESIDUALS):
            copy(CHECKS, col + str(row))


def calculate(test, folder, label, office):
    source = folder / (label + '.xlsx')
    test.save(source)
    out = folder / (label + '-output')
    out.mkdir()
    return recalculate(source, out, folder / (label + '-profile'), office, 90)


def test_book():
    test = TestBook()
    test.remove(test.active)
    return test


def verify(workbook, office):
    with TemporaryDirectory() as temp, Workbook(workbook) as original:
        folder = Path(temp)
        # One prior trial returns to the same closing amount after a temporary market change.
        # The other finishes with a different balance. Both tails are frozen below.
        old = test_book()
        add_case(old, original, 'returns', rates=((7, .25), (8, -.2)))
        add_case(old, original, 'gains', rates=((7, .25),))
        prior_path = calculate(old, folder, 'prior', office)
        old.close()

        current = test_book()
        clean = ('ira', 'workplace', 'flat', 'zero')
        add_case(current, original, 'ira', rates=((7, .25), (9, -.2)))
        add_case(current, original, 'workplace', ira=False, rates=((7, .25),))
        add_case(current, original, 'flat')
        add_case(current, original, 'zero', depleted=True, rates=((7, .25),))
        # Mutate one observed closing at a time, including zero-valued accounts and subtotals.
        bad = {}
        for index, (owner, col) in enumerate(CLOSINGS):
            case = f'field{index}'
            add_case(current, original, case)
            # The brokerage case also guards against rounding away a small nonzero discrepancy.
            current[tab(case, owner)][col + '7'] = 400.000001 if index == 0 else -123
            bad[case] = (RESIDUALS[index], 7)
        for case, invalid in (('missing', None), ('error', '=NA()')):
            add_case(current, original, case, depleted=True)
            current[tab(case, PROJECTION)]['BG7'] = invalid
            bad[case] = ('P', 7)
        add_case(current, original, 'cash_error')
        current[tab('cash_error', CASH)]['AM7'] = '=1/0'
        current[tab('cash_error', CASH)]['AM10'] = '=1/0'
        add_case(current, original, 'stale_roth')
        # A cached Roth opening and closing can agree while both are stale.
        current[tab('stale_roth', ROTH)]['D7'] = 450
        current[tab('stale_roth', ROTH)]['AI7'] = 450
        bad['stale_roth'] = ('Q', 7)

        with Workbook(prior_path) as old:
            for case in ('returns', 'gains'):
                add_case(current, original, case)
                for row in (7, 8, 9):
                    for sn, col in CLOSINGS:
                        addr = col + str(row)
                        current[tab(case, sn)][addr] = old[tab(case, sn)][addr].value
                bad[case] = ('N', 7)
            prior_ending = old[tab('gains', PROJECTION)]['BH9'].value

        result = calculate(current, folder, 'test', office)
        current.close()
        with Workbook(result) as book:
            for case in clean:
                checks = book[tab(case, CHECKS)]
                for row in ACTIVE:
                    equal(f'{case} annual check {row}', checks[f'J{row}'].value, 'OK')
                    for col in RESIDUALS:
                        actual = checks[f'{col}{row}'].value
                        if type(actual) not in (int, float) or actual != 0:
                            raise AssertionError(f'{case} {col}{row}: expected numeric zero, found {actual!r}')
                for col in ('J', *RESIDUALS):
                    if checks[col+'10'].value not in ('', None):
                        raise AssertionError(f'{case} {col}10: expected an empty cell beyond the horizon')
            for case, (col, row) in bad.items():
                checks = book[tab(case, CHECKS)]
                value = checks[f'{col}{row}'].value
                if case in ('missing', 'error'):
                    equal(f'{case} invalid actual', value, '#N/A')
                elif type(value) not in (int, float) or value == 0:
                    raise AssertionError(f'{case} {col}{row}: expected a nonzero closing difference, found {value!r}')
                equal(f'{case} annual status', checks[f'J{row}'].value, 'Review')
            cash_error = book[tab('cash_error', CHECKS)]
            equal('Cash-flow error remains visible at its source', book[tab('cash_error', CASH)]['AM7'].value, '#DIV/0!')
            for col in RESIDUALS:
                equal(f'Cash-flow error leaves {col}7 reconciled', cash_error[col+'7'].value, 0)
            equal('Active year with a cash-flow error', cash_error['J7'].value, 'Review')
            if cash_error['J10'].value not in ('', None):
                raise AssertionError('An inactive year showed a cash-flow status.')
            base_ending = book[tab('flat', PROJECTION)]['BH9'].value
            equal('Interior change with unchanged terminal balance', book[tab('returns', PROJECTION)]['BH9'].value, base_ending)
            equal('Stale suffix with prior terminal balance', book[tab('gains', PROJECTION)]['BH9'].value, prior_ending)
            if prior_ending == base_ending:
                raise AssertionError('The second prior trial must finish with a different balance.')
    print(f'PASS {len(clean)} valid paths, {len(bad)} corrupt closing paths and a cash-flow error, including all 12 closing checks.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--soffice', default='soffice')
    args = parser.parse_args()
    verify(args.workbook, args.soffice)
