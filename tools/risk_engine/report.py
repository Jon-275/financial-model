"""Compare identical market paths under fixed and flexible spending."""

from dataclasses import asdict
from math import sqrt
import platform

from . import __version__
from .annual import AnnualModel, reconcile
from .paths import Guardrail, ReturnModel, RiskInputs, TrialOutcome, percentile, summarize

ACCOUNT_NAMES = ('Taxable', 'Traditional', 'Roth', 'HSA')
MAX_TRIALS = 10_000


def _outcome(path, trial):
    return TrialOutcome(trial, path.first_failure_age, path.ending_balance_real, path.max_cash_shortfall)


def _result(path):
    return {
        'first_shortfall_age': path.first_failure_age,
        'ending_balance_real': path.ending_balance_real,
        'max_cash_shortfall_real': path.max_cash_shortfall,
        'years_with_cuts': sum(row.real_cut > 0 for row in path.years),
        'total_cuts_real': sum(row.real_cut for row in path.years),
    }


def _quantiles(values):
    ordered = sorted(values)
    return [percentile(ordered, fraction) for fraction in (.1, .5, .9)]


def _summarize(trials):
    stats = asdict(summarize(trials))
    n, p = stats['trials'], stats['success_rate']
    z = 1.959963984540054
    midpoint = (p + z * z / (2 * n)) / (1 + z * z / n)
    width = z * sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / (1 + z * z / n)
    stats['success_interval_95'] = [max(0, midpoint - width), min(1, midpoint + width)]
    return stats


def _allocations(scenario, returns, multiplier):
    a, i = scenario.allocation, scenario.inputs
    rows = []
    for retired, weights, selected, offset in ((False, a.working, i[15], 0), (True, a.retirement, i[16], 1)):
        covariance = returns.account_covariance(retired, multiplier, scenario.return_basis)
        for index, (account, mix, anchor) in enumerate(zip(
            ACCOUNT_NAMES, weights, (i[17 + offset], i[19 + offset], i[21 + offset], i[21 + offset]),
        )):
            own_anchor = sum(w * r for w, r in zip(mix, a.means)) / sum(mix)
            rows.append({'stage': 'Retirement' if retired else 'Working', 'profile': int(selected),
                         'account': account, 'base_return': anchor, 'weights': list(mix),
                         'return_volatility': sqrt(max(0, covariance[index][index]))
                         * (1 + anchor) / (1 + own_anchor)})
    return rows


def run_analysis(scenario, *, trials, seed, volatility, source_sha256, workbook_calculation,
                 progress=None):
    if not 1 <= trials <= MAX_TRIALS:
        raise ValueError(f'Trial count must be between 1 and {MAX_TRIALS:,}.')
    model = AnnualModel(scenario)
    if not model.periods:
        raise ValueError('Select an end age at least one year after the opening age to run portfolio risk.')
    a = scenario.allocation
    returns = ReturnModel(RiskInputs(a.means, a.volatilities, a.correlation, a.working, a.retirement))
    age, retirement = model.ages, model.retirement_age
    baseline = returns.random_path(age, retirement, seed, 0, 0, scenario.return_basis, model.baseline_returns)
    base = model.run(baseline, audit=True)
    parity = reconcile(model, base)
    planned_openings = model.planned_openings(base)
    if progress:
        progress(f'No-shock reconciliation passed for {parity["years"]} years '
                 f'and {parity["values_compared"]:,} values.')

    controls = scenario.sequence
    guardrail = Guardrail(controls[6], controls[7], controls[8])
    flexible_budget = (guardrail.trigger > 0 and guardrail.cut > 0 and guardrail.flexible_share > 0
                       and any(not p.working and p.flexible_base > 0 for p in model.periods))
    fixed, flexible = [], []
    fixed_balances = [[] for _ in age]
    flexible_balances = [[] for _ in age]
    cuts = []
    for trial in range(trials):
        path = returns.random_path(age, retirement, seed, trial, volatility, scenario.return_basis, baseline)
        a = model.run(path)
        b = (model.run(path, guardrail=guardrail, planned_openings=planned_openings)
             if flexible_budget else a)
        fixed.append(_outcome(a, trial))
        flexible.append(_outcome(b, trial))
        cuts.append(sum(row.real_cut > 0 for row in b.years))
        for slot, year in zip(fixed_balances, a.years):
            slot.append(year.real_balance)
        for slot, year in zip(flexible_balances, b.years):
            slot.append(year.real_balance)
        if progress and ((trial + 1) % 250 == 0 or trial + 1 == trials):
            progress(f'Completed {trial + 1:,} of {trials:,} paired paths.')

    stress = []
    applied_years = sum(retirement <= a < retirement + controls[5] for a in age)
    if applied_years:
        for equity in scenario.stress_returns:
            path = returns.sequence_path(age, retirement, controls[5], equity, baseline)
            a = model.run(path)
            b = (model.run(path, guardrail=guardrail, planned_openings=planned_openings)
                 if flexible_budget else a)
            stress.append({'equity_return': equity, 'fixed': _result(a), 'flexible': _result(b)})

    annual = []
    for row, fixed_year, flex_year in zip(base.years, fixed_balances, flexible_balances):
        annual.append({'year': row.year, 'age': row.age, 'base': row.real_balance,
                       'fixed_p10_p50_p90': _quantiles(fixed_year),
                       'flexible_p10_p50_p90': _quantiles(flex_year)})

    return {
        'provenance': {'engine_version': __version__, 'python': platform.python_version(),
                       'workbook_sha256': source_sha256, 'inputs_sha256': scenario.digest,
                       'workbook_calculation': workbook_calculation,
                       'scenario': scenario.selected, 'seed': str(seed),
                       'trials': trials, 'volatility_multiplier': volatility,
                       'return_basis': scenario.return_basis},
        'timeline': {'opening_year': model.opening_year, 'opening_age': model.opening_age,
                     'retirement_age': retirement, 'end_year': base.years[-1].year, 'end_age': age[-1]},
        'controls': {'trigger': controls[6], 'cut': controls[7], 'flexible_share': controls[8],
                     'stress_years_requested': controls[5], 'stress_years_applied': applied_years},
        'baseline': _result(base),
        'reconciliation': parity,
        'fixed': _summarize(fixed),
        'flexible': _summarize(flexible),
        'spending_cuts': {'trials_with_cuts': sum(count > 0 for count in cuts),
                          'median_years_with_cuts': _quantiles(cuts)[1]},
        'sequence': stress,
        'annual': annual,
        'asset_names': scenario.allocation.asset_names,
        'allocations': _allocations(scenario, returns, volatility),
    }
