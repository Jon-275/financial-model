"""Reconcile a saved desktop Excel run with an independent example using the same assumptions."""

import argparse
from collections import Counter
from math import isclose, ulp
from pathlib import Path
import re
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from native_capture_acceptance import compare_rows, same
from xlsx_read import Workbook, col_letter

from funding_search import ROWS as FUNDING_ROWS_PER_YEAR
import native_capture as trials
import native_funding_paths as funding
from native_bootstrap import _funding_row
from risk_engine.native_annual import NativeAnnualModel, price_path
from risk_engine.native_paths import NativeReturnModel
from risk_engine.paths import RiskInputs, TrialOutcome
from risk_engine.report import _representative_trials
from risk_engine.scenario import load_scenario


def value(sheet, row, column):
    cell = sheet.rows.get(row, {}).get(column)
    return cell.value if cell else None


def assumptions(book):
    market = book['Market_Paths']
    return {value(market, row, 53): value(market, row, 54) for row in market.rows
            if row >= 5 and value(market, row, 53) not in (None, 'Risk input 29')}


def positions(book):
    data = book['Risk_Data']
    n = book['Risk_Settings']['B29'].value
    codes = [*range(1, n + 1), *range(1001, 1001 + n), *range(2001, 2014),
             *range(2101, 2114), *range(2201, 2208)]
    metadata = [data[key].value for key in ('Z1', 'Z6', 'Z10', 'Z11', 'AB3')]
    if metadata != ['current', 'complete', n, book['Risk_Settings']['B30'].value, len(codes)]:
        raise AssertionError('A workbook has an incomplete or stale risk result.')
    if type(data['Z7'].value) not in (int, float) or data['Z7'].value <= 0:
        raise AssertionError('A workbook has no valid run date.')
    table = book[trials.SHEET]
    rows = {code: trials.FIRST + i for i, code in enumerate(codes)}
    for code, row in rows.items():
        if value(table, row, 1) != code or value(table, row, 3) != 1:
            raise AssertionError(f'Missing or invalid capture for trial {code}.')
    return rows


def check_reuse_note(message, visible_message, no_cut_pairs, expected=None):
    if not isinstance(message, str) or message != visible_message:
        raise AssertionError('The run details on Risk_Settings differ from the saved result.')
    number = r'[0-9]+(?:[, .\u00a0\u202f][0-9]{3})?'
    match = re.fullmatch(rf'Updated {number} paired paths? and 33 stress cases\. ({number}) '
                         r'(?:flexible path used its verified fixed result|flexible paths used their verified fixed results)\.',
                         message)
    if not match:
        raise AssertionError('The desktop run has no count of verified flexible paths.')
    reused = int(''.join(ch for ch in match[1] if ch.isdigit()))
    if reused > no_cut_pairs:
        raise AssertionError('The run reports more reused paths than its saved no-cut flexible records.')
    if expected is not None and reused != expected:
        raise AssertionError(f'Expected {expected} verified flexible paths in this benchmark, found {reused}.')
    return reused


def same_saved_value(before, after):
    """Excel can write 15 significant digits or shift the last binary digit when it saves."""
    if type(before) in (int, float) and type(after) in (int, float):
        return (before == after or format(before, '.15g') == format(after, '.15g')
                or abs(before - after) <= max(ulp(before), ulp(after)))
    return type(before) is type(after) and before == after


def restored_controls(reference, saved):
    status = saved['Risk_Data']['Z6'].value
    interrupted = status in ('cancelled', 'failed')
    if interrupted:
        message = saved['Risk_Data']['Z9'].value
        if (saved['Risk_Data']['Z1'].value != status or not isinstance(message, str)
                or not message.strip() or message == reference['Risk_Data']['Z9'].value):
            raise AssertionError('The interrupted run has no matching report status and explanation.')
    controls = {
        trials.SHEET: ('B3',),
        'Trial_Inputs': ('B70',),
        'Risk_Settings': ('B6', 'B29', 'B30'),
        'Risk_Data': ('Z6', 'Z7', 'Z8', 'Z9', 'Z10', 'Z11', 'AB3'),
    }
    for sheet, addresses in controls.items():
        for address in addresses:
            # Excel replaces the status and explanation when an interruption retains the prior batch.
            if interrupted and sheet == 'Risk_Data' and address in ('Z6', 'Z9'):
                continue
            if not same_saved_value(reference[sheet][address].value, saved[sheet][address].value):
                raise AssertionError(f'{sheet}!{address} was not restored.')
    selector = saved[trials.SHEET]['B3'].value
    years = saved['Trial_Inputs']['B70'].value
    if type(years) not in (int, float) or not 0 <= years <= trials.ANNUAL_CAPACITY or years != int(years):
        raise AssertionError('The saved inspection has an invalid horizon.')
    if selector is None or (type(selector) in (int, float) and selector == 0):
        if years != 0:
            raise AssertionError('The idle inspection has active years.')
    elif type(selector) not in (int, float) or not float(selector).is_integer() or selector < 1 or years < 2:
        raise AssertionError('The selected trial or its active horizon is invalid.')
    return selector, int(years)


def check_saved_funding_detail(before_rows, after_rows, years):
    """Inspect the cached results and formula locations in the visible search table."""
    before_number, before_header = next(before_rows, (None, {}))
    after_number, after_header = next(after_rows, (None, {}))
    widths = [col for col, cell in before_header.items() if cell.value == 'Best available draw']
    if before_number != 1 or after_number != 1 or len(widths) != 1:
        raise AssertionError('The withdrawal-search table has no recognized header.')
    width, = widths
    if width not in after_header or after_header[width].value != before_header[width].value:
        raise AssertionError('The saved withdrawal-search table has a different width.')

    first_inactive = 2 + max(0, years - 1) * FUNDING_ROWS_PER_YEAR
    last = 1 + (trials.ANNUAL_CAPACITY - 1) * FUNDING_ROWS_PER_YEAR
    for expected_row in range(2, last + 1):
        old_number, old = next(before_rows, (None, {}))
        new_number, new = next(after_rows, (None, {}))
        if old_number != expected_row or new_number != expected_row:
            raise AssertionError(f'Trial_Funding_Calc row {expected_row} is missing.')
        old_formulas = {col for col, cell in old.items() if col <= width and cell.formula is not None}
        new_formulas = {col for col, cell in new.items() if col <= width and cell.formula is not None}
        if old_formulas != new_formulas:
            raise AssertionError(f'Trial_Funding_Calc row {expected_row} lost or gained a search formula.')
        if (expected_row - 2) % FUNDING_ROWS_PER_YEAR == 0:
            if 1 not in old_formulas or old[1].formula != new[1].formula:
                raise AssertionError(f'Trial_Funding_Calc!A{expected_row} has a different year formula.')
        if expected_row >= first_inactive:
            for col, cell in new.items():
                if col <= width and cell.value not in (None, ''):
                    raise AssertionError(f'Inactive funding cell Trial_Funding_Calc!{cell.coordinate} still displays a value.')


def compare_restoration(reference, desktop, pass_cell=None):
    """Check that the prior financial batch and inspection survived a recovery exercise."""
    if Path(reference).samefile(desktop):
        raise AssertionError('Use separate workbooks saved before and after the recovery exercise.')
    with Workbook(reference) as before, Workbook(desktop) as after:
        if pass_cell is not None:
            sheet, separator, address = pass_cell.rpartition('!')
            if not separator or sheet not in after.parts or not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*', address):
                raise AssertionError('Specify the saved exercise verdict as Sheet!A1.')
            if after[sheet][address].value != 'PASS':
                raise AssertionError(f'The saved Excel exercise did not pass at {pass_cell}.')
        selector, years = restored_controls(before, after)
        stored_rows = positions(before)
        if after['Risk_Data']['Z6'].value not in ('cancelled', 'failed'):
            positions(after)
        count = 0
        for name, first, last, width in (
            (trials.SHEET, trials.FIRST, trials.LAST, trials.SIGNATURE_COLUMN),
            (funding.DATA_SHEET, funding.FIRST_ROW, funding.LAST_ROW, funding.SIGNATURE_COLUMN),
        ):
            original, saved = before[name], after[name]
            for row in range(first, last + 1):
                for col in range(1, width + 1):
                    if not same_saved_value(value(original, row, col), value(saved, row, col)):
                        raise AssertionError(f'Stored result {name}!{col_letter(col)}{row} was not restored.')
                count += width
        if selector:
            original, saved = before[trials.SHEET], after[trials.SHEET]
            current = trials.FIRST - 1
            if value(original, current, 1) != selector or value(saved, current, 1) != selector:
                raise AssertionError('The displayed trial differs from the restored inspection selector.')
            if selector not in stored_rows:
                raise AssertionError('The reference has no valid result for the selected trial.')
            compare_rows(saved, original, current, stored_rows[selector], trials.SIGNATURE_COLUMN,
                         selector, refresh_date=True)
        check_saved_funding_detail(before.iter_sheet_rows('Trial_Funding_Calc'),
                                   after.iter_sheet_rows('Trial_Funding_Calc'), years)
    print(f'PASS restored inspection, inactive funding detail and {count:,} stored result fields.')
    return True


def reduced_funding_rows(reference, expected, ref_rows, count):
    """Rank the smaller batch independently and rebuild its three full funding histories."""
    table = expected[trials.SHEET]
    outcomes = []
    for code in range(1, count + 1):
        row = ref_rows[code]
        age = value(table, row, 5)
        outcomes.append(TrialOutcome(trial=code, first_failure_age=None if age == '' else age,
                                     ending_balance_real=value(table, row, 7),
                                     max_cash_shortfall=value(table, row, 6)))
    calendar = NativeAnnualModel(load_scenario(reference))
    scenario = calendar.scenario
    end_age = int(scenario.inputs[53])
    inflation = scenario.inputs[14]
    baseline = price_path(calendar, (inflation,) * (end_age - calendar.opening_age), end_age=end_age)
    assets = scenario.allocation
    returns = NativeReturnModel(RiskInputs(assets.means, assets.volatilities, assets.correlation,
                                           assets.working, assets.retirement))
    settings = expected['Risk_Settings']
    seed, volatility, vary, sd, persistence, lower, upper = (settings[f'B{r}'].value for r in range(30, 37))
    ledgers = {}
    rows = []
    for _, _, code in _representative_trials(outcomes):
        # A one-pair run selects the same path for all three examples.
        if code not in ledgers:
            market = returns.market_path(baseline.ages, baseline.retirement_age, seed, code - 1, volatility,
                                         scenario.return_basis, baseline.baseline_returns, inflation=inflation,
                                         vary_inflation=bool(vary), inflation_sd=sd, persistence=persistence,
                                         lower=lower, upper=upper)
            annual = price_path(calendar, market.inflation, end_age=end_age) if vary else baseline
            ledgers[code] = _funding_row(code, annual, annual.run(market.returns, audit=True), None)
        rows.append(ledgers[code])
    return rows


def compare(reference, desktop, expected_reused=None):
    with Workbook(reference) as expected, Workbook(desktop) as actual:
        ref_rows, test_rows = positions(expected), positions(actual)
        if expected_reused is not None or actual['Market_Paths']['AH7'].formula:
            table = actual[trials.SHEET]
            no_cuts = sum(table[f'I{row}'].value == 0 for code, row in test_rows.items()
                          if 1001 <= code <= 2000 or 2101 <= code <= 2113)
            count = check_reuse_note(actual['Risk_Data']['Z9'].value, actual['Risk_Settings']['A44'].value,
                                     no_cuts, expected_reused)
            print(f'PASS run details report {count} verified flexible paths among {no_cuts} with no cut.')
        if not test_rows.keys() <= ref_rows.keys():
            raise AssertionError('The reference must contain all the desktop trials.')
        if actual['Risk_Data']['Z7'].value == expected['Risk_Data']['Z7'].value:
            raise AssertionError('The two saved run times are identical. Save the workbook after running Excel.')
        ref_inputs, test_inputs = assumptions(expected), assumptions(actual)
        if not ref_inputs or ref_inputs.keys() != test_inputs.keys():
            raise AssertionError('The workbooks use different sets of assumptions.')
        for name, left in ref_inputs.items():
            right = test_inputs[name]
            if type(left) in (int, float) and type(right) in (int, float):
                matches = isclose(left, right, rel_tol=1e-13, abs_tol=1e-13)
            else:
                matches = left == right
            if not matches:
                raise AssertionError(f'The assumption "{name}" differs.')

        differences, compared = Counter(), Counter()
        samples = []

        def check(label, code, column, observed, target):
            compared[label] += 1
            if not same(observed, target):
                differences[(label, code)] += 1
                if len(samples) < 10:
                    samples.append(f'{label} {code}, column {col_letter(column)}')

        ref_table, test_table = expected[trials.SHEET], actual[trials.SHEET]
        # The two assumption signatures include the run size. AB is the local refresh date.
        for code, row in test_rows.items():
            for column in range(3, trials.SIGNATURE_COLUMN):
                if column != 28:
                    check('trial', code, column, value(test_table, row, column),
                          value(ref_table, ref_rows[code], column))

        ref_paths, test_paths = expected[funding.DATA_SHEET], actual[funding.DATA_SHEET]
        ranked_rows = range(funding.FIRST_ROW, funding.LAST_ROW + 1)
        if len(ref_rows) == len(test_rows):
            expected_ledgers = [[value(ref_paths, row, col) for col in range(1, funding.SIGNATURE_COLUMN)]
                                for row in ranked_rows]
        else:
            expected_ledgers = reduced_funding_rows(reference, expected, ref_rows, actual['Risk_Settings']['B29'].value)
        total_offset = next(i for i, field in enumerate(funding.FIELDS) if field[0] == 'Total closing portfolio')
        for row, target in zip(ranked_rows, expected_ledgers, strict=True):
            code = value(test_paths, row, 1)
            if code not in test_rows or code >= 1001 or value(test_paths, row, 3) != 1:
                raise AssertionError('A ranked funding path is missing or invalid.')
            for year in range(trials.ANNUAL_CAPACITY):
                column = 4 + year * funding.FIELD_COUNT + total_offset
                check('funding balance', code, column, value(test_paths, row, column),
                      value(ref_table, ref_rows[code], trials.ANNUAL_FIRST + trials.ANNUAL_CAPACITY + year))
            check('ranked path ID', code, 1, code, target[0])
            for column in range(3, funding.SIGNATURE_COLUMN):
                check('funding ledger', code, column, value(test_paths, row, column), target[column - 1])
        print(f'{len(test_rows):,} outcomes. ' + ', '.join(f'{n:,} {key} fields' for key, n in compared.items()) + '.')
        if differences:
            print(f'{sum(differences.values()):,} differences in {len(differences)} comparisons by trial or path. '
                  + 'Examples: ' + ' / '.join(samples), file=sys.stderr)
            return False
        return True


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('reference', type=Path, help='Fictional reference or saved pre-exercise workbook')
    parser.add_argument('desktop', type=Path, help='Disposable copy saved after running Excel')
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument('--expect-reused', type=int, help='Require an independently verified reuse count for this example')
    modes.add_argument('--check-restoration', action='store_true',
                       help='Verify that a completed recovery exercise restored the original saved state')
    parser.add_argument('--expect-pass', metavar='SHEET!CELL',
                        help='Also require the native recovery exercise to have saved PASS in this cell')
    args = parser.parse_args()
    if args.expect_pass is not None and not args.check_restoration:
        parser.error('--expect-pass requires --check-restoration')
    passed = (compare_restoration(args.reference, args.desktop, args.expect_pass) if args.check_restoration
              else compare(args.reference, args.desktop, args.expect_reused))
    sys.exit(0 if passed else 1)
