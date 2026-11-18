"""Reconcile a saved desktop Excel run with an independent example using the same assumptions."""

import argparse
from collections import Counter
from math import isclose
from pathlib import Path
import re
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from native_capture_acceptance import same
from xlsx_read import Workbook, col_letter

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


def reduced_funding_rows(reference, expected, ref_rows, count):
    """Rank the smaller batch independently and rebuild its three full funding histories."""
    table = expected[trials.SHEET]
    outcomes = []
    for code in range(1, count + 1):
        row = ref_rows[code]
        outcomes.append(TrialOutcome(trial=code, first_failure_age=value(table, row, 5),
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
    parser.add_argument('reference', type=Path, help='Independently populated fictional example')
    parser.add_argument('desktop', type=Path, help='Disposable copy saved after running Excel')
    parser.add_argument('--expect-reused', type=int, help='Require an independently verified reuse count for this example')
    args = parser.parse_args()
    sys.exit(0 if compare(args.reference, args.desktop, args.expect_reused) else 1)
