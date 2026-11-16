"""Compare identical market paths under fixed and flexible spending."""

from dataclasses import asdict
from math import sqrt
import platform
from statistics import median

from . import __version__
from .annual import AnnualModel, SHORTFALL, reconcile
from .paths import Guardrail, ReturnModel, RiskInputs, TrialOutcome, percentile, summarize

ACCOUNT_NAMES = ('Taxable', 'Traditional', 'Roth', 'HSA')
MAX_TRIALS = 10_000


def _outcome(path, trial):
    return TrialOutcome(trial, path.first_failure_age, path.ending_balance_real, path.max_cash_shortfall)


def _first_shortfall(path, periods):
    """Use the same funding amounts that the no-shock audit reconciles to Excel."""
    for row, period in zip(path.years, periods, strict=True):
        if row.nominal_shortfall < SHORTFALL:
            continue
        cash = row.audit['Cash_Flow']
        projection = row.audit['Projection']
        qualified_healthcare = row.audit['Healthcare']['Q']
        return {
            'year': row.year,
            'age': row.age,
            'total_real': row.real_shortfall,
            'unpaid_spending_real': cash['Q'] / period.cpi,
            'unpaid_tax_real': cash['O'] / period.cpi,
            'required_spending_real': cash['D'] / period.cpi,
            'tax_due_real': (cash['F'] + cash['M']) / period.cpi,
            'income_available_real': cash['C'] / period.cpi,
            'income_used_real': min(cash['C'], cash['N'] + cash['P']) / period.cpi,
            'hsa_used_real': cash['E'] / period.cpi,
            'hsa_eligible_expense_real': qualified_healthcare / period.cpi,
            'brokerage_available_real': cash['H'] / period.cpi,
            'brokerage_draw_real': projection['V'] / period.cpi,
            'traditional_available_real': cash['I'] / period.cpi,
            'traditional_draw_real': projection['W'] / period.cpi,
            'roth_available_real': cash['J'] / period.cpi,
            'roth_draw_real': projection['Y'] / period.cpi,
        }
    return None


def _result(path, periods):
    first = _first_shortfall(path, periods)
    return {
        'first_shortfall_age': first['age'] if first is not None else None,
        'first_shortfall': first,
        'opening_balance_real': path.opening_balance,
        'ending_balance_real': path.ending_balance_real,
        'max_cash_shortfall_real': path.max_cash_shortfall,
        'years_with_cuts': sum(row.real_cut > 0 for row in path.years),
        'total_cuts_real': sum(row.real_cut for row in path.years),
    }


def _quantiles(values):
    ordered = sorted(values)
    return [percentile(ordered, fraction) for fraction in (.1, .5, .9)]


def _representative_trials(trials):
    """Choose observed ranks, ordering balance ties by first gap, worst gap, then trial ID."""
    def rank(trial):
        age = trial.first_failure_age
        return (trial.ending_balance_real, age if age is not None else float('inf'),
                -trial.max_cash_shortfall, trial.trial)

    ordered = sorted(trials, key=rank)
    last = len(ordered) - 1
    return [(label, p, ordered[(p * last + 50) // 100].trial)
            for label, p in (('Low', 10), ('Median', 50), ('High', 90))]


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
    for retired, weights, selected in ((False, a.working, i[15]), (True, a.retirement, i[16])):
        covariance = returns.account_covariance(retired, multiplier, scenario.return_basis)
        for index, (account, mix, anchor) in enumerate(zip(ACCOUNT_NAMES, weights, scenario.account_returns(retired))):
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

    def run_pair(path, *, audit=False):
        fixed_path = model.run(path, audit=audit)
        flexible_path = (model.run(path, guardrail=guardrail, planned_openings=planned_openings, audit=audit)
                         if flexible_budget else fixed_path)
        return fixed_path, flexible_path

    def detail(path):
        fixed_path, flexible_path = run_pair(path, audit=True)
        return {'fixed': _result(fixed_path, model.periods), 'flexible': _result(flexible_path, model.periods)}

    fixed, flexible = [], []
    fixed_balances = [[] for _ in age]
    flexible_balances = [[] for _ in age]
    cuts = []
    for trial in range(trials):
        path = returns.random_path(age, retirement, seed, trial, volatility, scenario.return_basis, baseline)
        a, b = run_pair(path)
        fixed.append(_outcome(a, trial))
        flexible.append(_outcome(b, trial))
        cuts.append(sum(row.real_cut > 0 for row in b.years))
        for slot, year in zip(fixed_balances, a.years):
            slot.append(year.real_balance)
        for slot, year in zip(flexible_balances, b.years):
            slot.append(year.real_balance)
        if progress and ((trial + 1) % 250 == 0 or trial + 1 == trials):
            progress(f'Completed {trial + 1:,} of {trials:,} paired paths.')

    # Re-create the selected seeds with funding details. Full ledgers for all
    # the trials would consume much more memory and are unnecessary here.
    representative = []
    replayed = {}
    for label, p, trial in _representative_trials(fixed):
        if trial not in replayed:
            path = returns.random_path(age, retirement, seed, trial, volatility, scenario.return_basis, baseline)
            replayed[trial] = detail(path)
        representative.append({'label': label, 'percentile': p, 'trial': trial, **replayed[trial]})

    stress = []
    applied_years = sum(retirement <= a < retirement + controls[5] for a in age)
    if applied_years:
        for equity in scenario.stress_returns:
            path = returns.sequence_path(age, retirement, controls[5], equity, baseline)
            stress.append({'equity_return': equity, **detail(path)})

    annual = []
    for row, fixed_year, flex_year in zip(base.years, fixed_balances, flexible_balances):
        annual.append({'year': row.year, 'age': row.age, 'base': row.real_balance,
                       'fixed_p10_p50_p90': _quantiles(fixed_year),
                       'flexible_p10_p50_p90': _quantiles(flex_year)})

    paths_with_cuts = [count for count in cuts if count > 0]
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
        'baseline': _result(base, model.periods),
        'reconciliation': parity,
        'fixed': _summarize(fixed),
        'flexible': _summarize(flexible),
        'spending_cuts': {'trials_with_cuts': len(paths_with_cuts),
                          'median_years_with_cuts': median(cuts),
                          'median_years_when_cut': median(paths_with_cuts) if paths_with_cuts else None},
        'representative_paths': representative,
        'sequence': stress,
        'annual': annual,
        'asset_names': scenario.allocation.asset_names,
        'allocations': _allocations(scenario, returns, volatility),
    }
