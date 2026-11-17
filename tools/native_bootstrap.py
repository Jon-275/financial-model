"""Populate a freshly built template from the independent annual model.

Desktop Excel refreshes the same output areas with its VBA runner. This tool
only prepares and verifies the initial example distributed with a release.
"""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
from math import isfinite
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils.datetime import to_excel

import native_capture as trial
import native_funding_paths as funding_paths
from native_results import ENGINE, SCHEMA
from prepare_native_cache import prepare_cache_copy
from risk_engine.annual import SHORTFALL, reconcile
from risk_engine.native_annual import NativeAnnualModel, price_path
from risk_engine.native_paths import NativeReturnModel
from risk_engine.paths import Guardrail, RiskInputs
from risk_engine.report import _first_shortfall, _outcome, _representative_trials
from risk_engine.scenario import load_scenario
from workbook_recalc import clear_formula_results, recalculate
from xlsx_values import write_values


@dataclass(frozen=True)
class RunSettings:
    count: int
    seed: int
    volatility: float
    vary_inflation: bool
    inflation_sd: float
    persistence: float
    lower: float
    upper: float


def _read_plan(calculated):
    book = load_workbook(calculated, read_only=True, data_only=True, keep_links=False)
    try:
        if book['Risk_Data']['Z2'].value != SCHEMA or book['Risk_Data']['Z4'].value != ENGINE:
            raise ValueError('Build the native template before preparing its market results.')
        if book['Market_Paths']['AL3'].value != 'OK':
            raise ValueError('Review the market assumptions before building the template.')
        if book['Risk_Data']['Z1'].value != 'missing':
            raise ValueError('Build from a fresh template with empty market result areas.')
        context = book['Market_Paths']['BB3'].value
        if not isinstance(context, str) or not context.startswith('native-risk-v1|'):
            raise ValueError('The calculated market assumption record is missing.')
        values = [row[0] for row in book['Risk_Settings'].iter_rows(
            min_col=2, max_col=2, min_row=29, max_row=36, values_only=True)]
        count, seed, volatility, vary, deviation, persistence, lower, upper = values
        if not 1 <= count <= trial.CAPACITY or int(count) != count:
            raise ValueError('Enter a whole number of paired paths between 1 and 1,000.')
        if not 0 <= seed <= 2**32 - 1 or int(seed) != seed:
            raise ValueError('Enter a whole 32-bit random seed.')
        settings = RunSettings(int(count), int(seed), volatility, bool(vary), deviation, persistence, lower, upper)
        return (settings, context, int(book['Inputs']['B4'].value),
                int(book['Inputs']['B53'].value), book.epoch)
    finally:
        book.close()


def _stored_row(cells, row, values):
    for column, value in enumerate(values, 1):
        if value is None:
            continue
        if isinstance(value, (int, float)) and not isfinite(value):
            raise ValueError('A captured financial result is not finite.')
        cells[f'{get_column_letter(column)}{row}'] = value


def _trial_row(code, annual, path, context, timestamp):
    years, last = path.years, path.years[-1]
    gap = _first_shortfall(path, annual.periods)
    end_cash = last.audit['Cash_Flow']
    end_index = annual.periods[-1].cpi
    gap_fields = (None,) * 17
    if gap:
        period = annual.periods[gap['age'] - annual.opening_age - 1]
        gap_fields = (gap['year'], period.cpi, *(gap[key] for key in (
            'total_real', 'unpaid_spending_real', 'unpaid_tax_real', 'required_spending_real', 'tax_due_real',
            'income_available_real', 'income_used_real', 'hsa_used_real', 'hsa_eligible_expense_real',
            'brokerage_available_real', 'brokerage_draw_real', 'traditional_available_real', 'traditional_draw_real',
            'roth_available_real', 'roth_draw_real')))
    positions = [code, context, 1, int(gap is None), path.first_failure_age, path.max_cash_shortfall,
                 path.ending_balance_real, last.nominal_balance, sum(y.real_cut > 0 for y in years),
                 sum(y.real_cut for y in years), *gap_fields, int(timestamp), end_cash['O'] / end_index,
                 end_cash['Q'] / end_index, end_index, gap['total_real'] * gap_fields[1] if gap else None,
                 last.age]
    if len(positions) != trial.ANNUAL_FIRST - 1:
        raise ValueError('Risk trial fields do not match the workbook layout.')
    for column in ('real_balance', 'nominal_balance'):
        positions.extend([path.opening_balance, *(getattr(y, column) for y in years)])
        positions.extend([None] * (trial.ANNUAL_CAPACITY - 1 - len(years)))
    return (*positions, context)


def _funding_year(p, cash, health, cut):
    return (p['B'], p['C'], p['G'], cash['C'], min(cash['C'], cash['N'] + cash['P']),
            p['AW'], p['AX'], p['BA'], cash['D'], cash['F'] + cash['M'], health['Q'], cash['E'],
            cash['N'], cash['P'], cash['H'], cash['I'], cash['J'], p['V'], p['W'], p['Y'],
            p['AA'], p['M'], p['BB'], p['BC'], p['BE'], p['BF'], p['BG'], p['BH'],
            cash['O'], cash['Q'], cash['AU'], cash['AS'], cash['AV'], cut, p['N'], p['I'],
            cash['U'], cash['M'])


def _funding_row(code, annual, path, context):
    scenario = annual.scenario
    fields = [code, context, 1, *_funding_year(
        scenario.opening_projection, scenario.opening_cash, scenario.opening_health, 0)]
    for year, period in zip(path.years, annual.periods, strict=True):
        cut = year.real_cut * period.cpi
        projection = {**year.audit['Projection'], 'B': year.year, 'C': year.age, 'G': period.cpi,
                      'AW': period.core_spending - cut, 'AX': period.one_off,
                      'N': period.social_security, 'I': period.wages}
        cash = year.audit['Cash_Flow']
        residual = cash['Y'] - cash['N'] - cash['P'] - cash['R'] - cash['T']
        if abs(residual) > SHORTFALL or abs(cash['AW']) > SHORTFALL:
            raise ValueError(f'Funding path {code} has an unreconciled year.')
        fields.extend(_funding_year(projection, {**cash, 'U': residual}, year.audit['Healthcare'], cut))
    years = len(path.years) + 1
    if len(fields) != 3 + funding_paths.FIELD_COUNT * years:
        raise ValueError('Funding path fields do not match the workbook layout.')
    fields.extend([None] * (funding_paths.FIELD_COUNT * (trial.ANNUAL_CAPACITY - years)))
    return (*fields, context)


def build_captures(calendar, settings, end_age, context, timestamp, progress=print):
    """Capture each paired path, required stress and ranked funding ledger."""
    scenario = calendar.scenario
    base_rate = scenario.inputs[14]
    def priced(end, rate, *, floor=0):
        return price_path(calendar, (rate,) * (end - calendar.opening_age), end_age=end, medical_floor=floor)

    baseline = priced(end_age, base_rate)
    alloc = scenario.allocation
    returns = NativeReturnModel(RiskInputs(alloc.means, alloc.volatilities, alloc.correlation,
                                          alloc.working, alloc.retirement))
    planned = baseline.planned_openings(baseline.run(baseline.baseline_returns))
    guardrail = Guardrail(*(scenario.sequence[row] for row in (6, 7, 8)))
    flexible = {'guardrail': guardrail, 'planned_openings': planned}

    inflation_settings = ({
        'vary_inflation': True, 'inflation_sd': settings.inflation_sd, 'persistence': settings.persistence,
        'lower': settings.lower, 'upper': settings.upper,
    } if settings.vary_inflation else {})

    def random_path(code):
        paths = returns.market_path(baseline.ages, baseline.retirement_age, settings.seed, code - 1,
            settings.volatility, scenario.return_basis, baseline.baseline_returns, inflation=base_rate,
            **inflation_settings)
        annual = price_path(calendar, paths.inflation, end_age=end_age) if settings.vary_inflation else baseline
        return annual, paths.returns

    primary, secondary, outcomes = {}, {}, []
    def capture(index, code, annual, market, **kwargs):
        path = annual.run(market, audit=True, **kwargs)
        _stored_row(primary, trial.FIRST + index, _trial_row(code, annual, path, context, timestamp))
        return path

    n = settings.count
    for index in range(n):
        code = index + 1
        annual, market = random_path(code)
        path = capture(index, code, annual, market)
        outcomes.append(_outcome(path, code))
        capture(n + index, 1000 + code, annual, market, **flexible)
        if progress and (code % 100 == 0 or code == n):
            progress(f'Calculated {code:,} of {n:,} paired paths.')
    for index, stock_return in enumerate(scenario.stress_returns):
        market = returns.sequence_path(baseline.ages, max(baseline.opening_age + 1, baseline.retirement_age),
                                       int(scenario.sequence[5]), stock_return, baseline.baseline_returns)
        capture(2 * n + index, 2001 + index, baseline, market)
        capture(2 * n + 13 + index, 2101 + index, baseline, market, **flexible)
    for index, code in enumerate(range(2201, 2208)):
        end = max(end_age, 100 if code == 2203 else 105 if code in (2204, 2205) else end_age)
        rate = .04 if code == 2201 else .05 if code in (2202, 2205) else base_rate
        floor = {2206: 18_000, 2207: 25_000}.get(code, 0)
        annual = priced(end, rate, floor=floor)
        market = annual.baseline_returns
        if code == 2205:
            gross = tuple(1 + max(-1, mean - .02) for mean in alloc.means)
            market = tuple(returns._apply(age >= annual.retirement_age, original, gross)
                           for age, original in zip(annual.ages, market, strict=True))
        capture(2 * n + 26 + index, code, annual, market)
    for offset, (_, _, code) in enumerate(_representative_trials(outcomes)):
        annual, market = random_path(code)
        path = annual.run(market, audit=True)
        _stored_row(secondary, funding_paths.FIRST_ROW + offset, _funding_row(code, annual, path, context))
    return {trial.SHEET: primary, funding_paths.DATA_SHEET: secondary, 'Risk_Data': {
        'Z6': 'complete', 'Z7': timestamp, 'Z8': context,
        'Z9': f'Updated {n:,} paired paths and 33 stress cases.', 'Z10': n, 'Z11': settings.seed}}


def _calculate(source, work, name, executable, timeout):
    clean, calculated = work / f'{name}-clean', work / f'{name}-calculated'
    clean.mkdir()
    calculated.mkdir()
    temporary = clean / 'model.xlsx'
    clear_formula_results(source, temporary)
    return recalculate(temporary, calculated, work / f'{name}-profile', executable, timeout)


def verify_calculated(path, settings, context):
    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        results = book['Risk_Data']
        expected = {'Z1': 'current', 'Z5': 'current', 'Z6': 'complete', 'Z8': context,
                    'Z10': settings.count, 'Z11': settings.seed, 'AB3': 2 * settings.count + 33}
        for cell, value in expected.items():
            if results[cell].value != value:
                raise ValueError(f'The recalculated example failed its {cell} release check.')
        ids = [row[0] for row in results.iter_rows(min_col=50, max_col=50, min_row=5, max_row=7, values_only=True)]
        funding_ids = [row[0] for row in book[funding_paths.DATA_SHEET].iter_rows(
            min_col=1, max_col=1, min_row=8, max_row=10, values_only=True)]
        if ids != funding_ids:
            raise ValueError('The three full funding paths do not match the report ranks.')
    finally:
        book.close()


def bootstrap(source, destination, executable, timeout):
    source, destination = Path(source), Path(destination)
    if destination.exists() or source.resolve() == destination.resolve():
        raise ValueError('Use a new output path for the release workbook.')
    destination.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(dir=destination.parent) as directory:
        work = Path(directory)
        baseline = _calculate(source, work, 'baseline', executable, timeout)
        settings, context, selected, end_age, epoch = _read_plan(baseline)
        calendar = baseline
        if end_age < 105:
            extended = work / 'extended.xlsx'
            write_values(source, extended, {'Inputs': {f'{get_column_letter(3 + selected)}53': 105}})
            calendar = _calculate(extended, work, 'extended', executable, timeout)
        model = NativeAnnualModel(load_scenario(calendar))
        reconciled = reconcile(model, model.run(model.baseline_returns, audit=True))
        print(f'Checked {reconciled["values_compared"]:,} annual amounts against the workbook.', flush=True)
        changes = build_captures(model, settings, end_age, context, to_excel(datetime.now(), epoch))
        filled = work / 'captured.xlsx'
        write_values(source, filled, changes, shared_strings=(context,))
        cache_input = work / 'cache-input.xlsx'
        records = prepare_cache_copy(filled, cache_input, settings.count, context)
        print(f'Checked {records:,} complete assumption records before report recalculation.', flush=True)
        report = _calculate(cache_input, work, 'report', executable, timeout)
        verify_calculated(report, settings, context)
        final = work / 'release.xlsx'
        subprocess.run([sys.executable, str(Path(__file__).with_name('cache_model.py')), str(filled),
                        str(report), str(final)], check=True)
        verify_calculated(final, settings, context)
        os.link(final, destination)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('destination', type=Path)
    parser.add_argument('--soffice', default='soffice')
    parser.add_argument('--timeout', type=int, default=600)
    args = parser.parse_args()
    bootstrap(args.source, args.destination, args.soffice, args.timeout)
