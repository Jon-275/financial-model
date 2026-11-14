"""Roll forward the workbook's cash and accounts under one market path."""

from bisect import bisect_right
from dataclasses import dataclass, field
from itertools import accumulate
from math import floor, isfinite

from .funding import FundingInputs, solve_funding
from .scenario import Scenario, YearSnapshot
from roth_contributions import annual_limits, reduced_limit

from .tax import TaxInputs, TaxTables, YearTax

SHORTFALL = 0.01


@dataclass(slots=True)
class State:
    taxable: float
    basis: float
    traditional: float
    hsa: float
    roth: tuple
    regular: tuple
    conversions: tuple
    seasoned: tuple
    first_ira_year: tuple
    federal_loss: float
    ca_loss: float
    irmaa_magi: tuple
    payable: float = 0
    sepp_balance: float = 0
    conversion_history: list = field(default_factory=list)

    @property
    def balance(self):
        return self.taxable + self.traditional + sum(self.roth) + self.hsa


@dataclass(frozen=True, slots=True)
class Period:
    sheet: YearSnapshot
    year_tax: YearTax
    age: int
    spouse_age: int
    owner_age: int
    working: bool
    single: bool
    cpi: float
    tax_index: float
    baseline: tuple
    wages: float
    deferral: float
    hsa_contribution: float
    match: float
    social_security: float
    other_income: float
    exempt: float
    payroll: float
    sdi: float
    cash_income: float
    core_spending: float
    flexible_base: float
    one_off: float
    medical: float
    qualified_medical: float
    medicare_count: int
    surcharge_qualifies: bool
    legacy_irmaa_count: int
    irmaa_single: bool
    aca: tuple
    roth_target: float
    savings_capacity: tuple
    roth_phaseout: tuple | None
    savings_share: float
    rmd_divisor: float
    sepp_window: bool


@dataclass(frozen=True, slots=True)
class YearResult:
    year: int
    age: int
    real_balance: float
    nominal_balance: float
    real_shortfall: float
    nominal_shortfall: float
    real_cut: float
    audit: dict | None = None


@dataclass(frozen=True, slots=True)
class PathResult:
    years: tuple
    opening_balance: float

    @property
    def ending_balance_real(self):
        return self.years[-1].real_balance if self.years else self.opening_balance

    @property
    def first_failure_age(self):
        return next((row.age for row in self.years if row.nominal_shortfall >= SHORTFALL), None)

    @property
    def max_cash_shortfall(self):
        return max((row.real_shortfall for row in self.years), default=0.0)


class AnnualModel:
    """Cache scenario inputs that are independent of returns, once per run."""

    def __init__(self, scenario: Scenario):
        self.scenario = scenario
        inputs, tax = scenario.inputs, scenario.tax
        self.inputs = inputs
        self.tables = TaxTables.from_cells(tax)
        self.owner = int(inputs[97]) - 1
        self.detailed_health = inputs[101] == 1
        self.retirement_age = int(inputs[10])
        self.opening_year = int(scenario.opening_projection['B'])
        self.opening_age = int(inputs[8])
        self.sepp_start = max(1, self.retirement_age - self.opening_age)
        self.sepp_owner_age = int(inputs[8 + self.owner]) + self.sepp_start
        self.surcharges = tuple(tax[f'E{row}'] for row in range(89, 95))
        self.irmaa_single = tuple(tax[f'B{row}'] for row in range(182, 188))
        self.irmaa_joint = tuple(tax[f'B{row}'] for row in range(89, 95))
        self.rmd_divisors = (27.4, *(tax[f'B{row}'] for row in range(99, 132)),
                             *(tax[f'H{row}'] for row in range(99, 114)))
        self.periods = tuple(self._period(row) for row in scenario.years)
        self.ages = tuple(period.age for period in self.periods)
        self.baseline_returns = tuple(period.baseline for period in self.periods)

    def _opening_state(self):
        p, r = self.scenario.opening_projection, self.scenario.opening_roth
        lookback = ((self.inputs[122], self.inputs[123]) if self.detailed_health
                    else (p['AK'] + p['AG'],) * 2)
        return State(p['BC'], p['BD'], p['BE'], p['BG'],
                     tuple(r[c] for c in ('AI', 'AJ', 'AK', 'AL')),
                     (r['AC'], r['AD']), (r['AE'], r['AF']), (r['AG'], r['AH']),
                     (int(r['N']), int(r['O'])), p['BW'], p['BY'], lookback)

    def _period(self, row):
        i, p, h = self.inputs, row.projection, row.health
        age, spouse_age = int(p['C']), int(p['D'])
        owner_age = (age, spouse_age)[self.owner]
        working, single = p['E'] == 'Working', p['F'] == 'Single'
        year_tax = self.tables.for_year(row.year, i[14], p['H'], single, age, spouse_age)
        rate_col = 0 if working else 1
        roth_return = i[21 + rate_col]
        baseline = (i[17 + rate_col], i[19 + rate_col], roth_return, roth_return)
        fixed = i[38] * (1 if row.year < i[39] else i[40] / 12 if row.year == i[39] else 0)
        fixed *= i[56] if single else 1
        wages, deferral, hsa_contribution, match = (p[c] for c in ('I', 'J', 'K', 'L'))
        ss, other, exempt = (p[c] for c in ('N', 'O', 'AG'))
        payroll, sdi = year_tax.payroll_taxes(p['BT'], p['BU'])
        cash_income = wages - deferral - hsa_contribution + ss + other * i[86] + exempt * i[89]
        if self.detailed_health:
            medicare_count = int(h['W'] + h['X'])
            irmaa_single = h['Z'] == 'Single'
            medical = sum(h[c] for c in 'FGHIJK')
            owner_alive = i[121] == 1 or not single
            surcharge_qualifies = not working and owner_alive and (age if i[121] == 1 else spouse_age) >= 65
            qualified = 0
            if not working and owner_alive:
                qualified = h['H'] + h['I'] + h['K'] * i[124]
                qualified += (h['F'] if h['D'] == 'COBRA' else 0)
                qualified += (h['G'] if h['E'] == 'COBRA' else 0)
                if surcharge_qualifies:
                    medical_index = (1 + i.get(132, i[14])) ** p['A']
                    qualified += (h['W'] * i[108] + h['X'] * i[109]) * medical_index
            legacy_irmaa_count = 0
        else:
            medicare_count = 0
            irmaa_single = single
            legacy_irmaa_count = (1 + int(not single and spouse_age >= 65)) if age >= 65 else 0
            if 130 in i:
                legacy_irmaa_count = int(not working) * (int(age >= 65) + int(not single and spouse_age >= 65))
            medical = p['AY'] if working else h['P'] - h['L']
            qualified, surcharge_qualifies = (medical, True) if not working else (0, False)
        divisor = (self.rmd_divisors[min(owner_age, 120) - 72]
                   if owner_age >= max(72, i[67]) else 0)
        sepp = (i[68] == 1 and not working and self.sepp_owner_age < 60
                and owner_age < max(60, self.sepp_owner_age + 5))
        limits = annual_limits(self.scenario.tax, row.year, p['H'], i[14], single)
        if limits is None:
            limit, catchup = (self.scenario.tax[c] * p['G'] for c in ('B61', 'B62'))
            roth_phaseout = None
        else:
            limit, catchup, low, high = limits
            roth_phaseout = (low, high)
        return Period(
            row, year_tax, age, spouse_age, owner_age, working, single, p['G'], p['H'], baseline,
            wages, deferral, hsa_contribution, match, ss, other, exempt, payroll, sdi, cash_income,
            p['AW'], max(0, p['AW'] - fixed), p['AX'], medical, qualified, medicare_count,
            surcharge_qualifies, legacy_irmaa_count, irmaa_single, (h['M'], h['N'], h['O']),
            i[35] * p['G'] if working else 0,
            (limit + (catchup if age >= 50 else 0),
             0 if single else limit + (catchup if spouse_age >= 50 else 0)), roth_phaseout,
            1 if single else i[91], divisor, sepp,
        )

    def _surcharge(self, state, period):
        magi, index = state.irmaa_magi[0], period.tax_index
        if not self.detailed_health:
            limits = self.irmaa_single if period.single else self.irmaa_joint
            tier = max(0, bisect_right(limits, magi / index) - 1)
            return self.surcharges[tier] * index * period.legacy_irmaa_count
        if not period.medicare_count or magi < 0:
            return 0
        joint = 1 if period.irmaa_single else 2
        top_joint = 1 if period.irmaa_single else 1.5
        top_index = (1 if period.sheet.year <= 2027
                     else max(1, index / (1 + self.inputs[14])))

        def threshold(value, factor):
            return floor(value * factor / 1000 + 0.5) * 1000

        tier = sum(magi > threshold(value, index) * joint for value in self.irmaa_single[1:-1])
        tier += magi >= threshold(self.irmaa_single[-1], top_index) * top_joint
        return self.surcharges[tier] * index * period.medicare_count

    def _roth_access(self, state, period, year):
        ages = period.age, period.spouse_age
        matured = state.conversion_history[-5] if len(state.conversion_history) >= 5 else (0, 0)
        seasoned = tuple(
            state.conversions[j] if age >= 60
            else min(state.conversions[j], state.seasoned[j] + matured[j])
            for j, age in enumerate(ages)
        )
        ira = tuple(
            state.roth[j] if age >= 60 and 0 < state.first_ira_year[j] <= year - 5
            else min(state.roth[j], state.regular[j] + seasoned[j])
            for j, age in enumerate(ages)
        )
        i = self.inputs
        workplace = tuple(
            state.roth[j + 2] if max(60, i[95 + j]) <= age and i[95 + j] < 126
            and 0 < i[93 + j] <= year - 5 else 0
            for j, age in enumerate(ages)
        )
        return seasoned, ira + workplace

    def planned_openings(self, baseline):
        return (baseline.opening_balance, *(row.nominal_balance for row in baseline.years))[:-1]

    def run(self, returns, *, guardrail=None, planned_openings=None, audit=False):
        if len(returns) != len(self.periods):
            raise ValueError('Market path does not cover the active projection years.')
        if guardrail is not None and (planned_openings is None or len(planned_openings) != len(returns)):
            raise ValueError('Guardrail requires a reconciled opening balance for each year.')
        state = self._opening_state()
        opening_balance = state.balance
        rows = []
        for position, (period, rates) in enumerate(zip(self.periods, returns), 1):
            if len(rates) != 4 or any(rate < -1 or not isfinite(rate) for rate in rates):
                raise ValueError(f'Invalid market return for {period.sheet.year}.')
            if position == self.sepp_start:
                state.sepp_balance = state.traditional
            cut = 0 if guardrail is None else guardrail.reduction(
                not period.working, state.balance, planned_openings[position - 1], period.flexible_base)
            rows.append(self._advance(state, period, rates, cut, audit))
        return PathResult(tuple(rows), opening_balance)

    def _advance(self, state, period, rates, cut, audit):
        i, p = self.inputs, period
        year = p.sheet.year
        surcharge = self._surcharge(state, p)
        healthcare = p.medical + (surcharge if self.detailed_health or not p.working else 0)
        qualified = min(healthcare, p.qualified_medical + (surcharge if p.surcharge_qualifies else 0))
        hsa_paid = min(state.hsa, qualified)
        required = p.core_spending - cut + p.one_off + healthcare

        offset = 0 if p.working else 1
        interest, dividends = state.taxable * i[65 + offset], state.taxable * i[63 + offset]
        basis_ratio = state.basis / state.taxable if state.taxable > 0 else 0
        sepp = min(state.traditional, state.sepp_balance * i[69]) if p.sepp_window else 0
        rmd = state.traditional / p.rmd_divisor if p.rmd_divisor else 0
        mandatory = min(state.traditional, max(sepp, rmd))
        headroom = max(0, i[49] * p.cpi + p.year_tax.federal_deduction
                       - (mandatory + p.other_income + interest + 0.85 * p.social_security))
        traditional_limit = mandatory if p.working or p.owner_age < 60 or sepp > 0 else state.traditional
        seasoned, access = self._roth_access(state, p, year)
        roth_limit = 0 if p.working else sum(access)
        net_wages = p.wages - p.deferral - p.hsa_contribution
        taxes = TaxInputs(
            ordinary_income=net_wages + p.other_income, qualified_dividends=dividends,
            investment_interest=interest, tax_exempt_interest=p.exempt, social_security=p.social_security,
            hsa_contribution=p.hsa_contribution, payroll_tax=p.payroll, sdi=p.sdi,
            opening_federal_loss=state.federal_loss, opening_ca_loss=state.ca_loss,
        )
        opening_payable = state.payable
        opening_taxable = state.taxable
        funding = solve_funding(p.year_tax, taxes, FundingInputs(
            cash_income=p.cash_income, required_spending=required, hsa_paid=hsa_paid,
            opening_tax_payable=opening_payable, taxable_available=state.taxable,
            traditional_available=traditional_limit, roth_available=roth_limit,
            mandatory=mandatory, headroom=headroom, opening_traditional=state.traditional,
            basis_ratio=basis_ratio, requested_conversion=i[47] * p.cpi, order=int(i[48]),
            conversion_allowed=not p.working and sepp == 0 and p.owner_age <= i[50],
            aca_actual=p.aca[0], aca_benchmark=p.aca[1], aca_fpl=p.aca[2],
        ))
        if funding.status == 'Check funding convergence':
            raise ValueError(f'Funding search did not converge for {year}.')
        tax = funding.tax
        draws = (funding.taxable_draw, funding.traditional_draw, funding.roth_draw)
        cash_income = p.cash_income + funding.premium_credit
        cash_available = cash_income + sum(draws)
        tax_due = opening_payable + tax.total_tax
        paid_tax = min(tax_due, cash_available)
        state.payable = max(0, tax_due - paid_tax)
        due_spending = required - hsa_paid
        paid_spending = min(due_spending, max(0, cash_available - paid_tax))
        unpaid_spending = max(0, due_spending - paid_spending)
        shortfall = state.payable + unpaid_spending
        residual = max(0, cash_available - paid_tax - paid_spending)
        roth_magi = tax.agi - funding.conversion
        capacity = tuple(reduced_limit(cap, roth_magi, p.roth_phaseout) if cap > 0 else 0
                         for cap in p.savings_capacity)
        savings_limit = 0
        if p.working and i[98] == 1:
            share = p.savings_share
            savings_limit = min(max(0, net_wages),
                                capacity[0] / share if share else p.roth_target,
                                capacity[1] / (1 - share) if share < 1 else p.roth_target)
        roth_savings = min(p.roth_target, savings_limit, residual)
        deposit = max(0, residual - roth_savings)

        removal = min(state.basis, draws[0] * basis_ratio)
        state.taxable = max(0, state.taxable - draws[0] + deposit) * (1 + rates[0])
        basis_addition = deposit + dividends + interest if state.taxable > 0 else 0
        state.basis = max(0, state.basis - removal) + basis_addition if state.taxable > 0 else 0
        state.traditional = max(0, state.traditional + p.deferral + p.match - draws[1]
                                - funding.conversion) * (1 + rates[1])
        state.hsa = max(0, state.hsa + p.hsa_contribution - hsa_paid) * (1 + rates[3])
        paid_roth = tuple(min(limit, max(0, draws[2] - used))
                          for limit, used in zip(access, accumulate((0,) + access)))
        primary = roth_savings * p.savings_share
        saved = primary, roth_savings - primary
        converted = (funding.conversion, 0) if self.owner == 0 else (0, funding.conversion)
        used_conversion = tuple(max(0, paid_roth[j] - state.regular[j]) for j in range(2))
        state.roth = tuple(
            max(0, opening + saving + conversion - draw) * (1 + rates[2])
            for opening, saving, conversion, draw in zip(state.roth, saved + (0, 0), converted + (0, 0), paid_roth)
        )
        first_ira_year = state.first_ira_year
        state.first_ira_year = tuple(first_ira_year[j] or (year if saved[j] + converted[j] > 0 else 0)
                                     for j in range(2))
        state.regular = tuple(max(0, state.regular[j] - paid_roth[j]) + saved[j] for j in range(2))
        state.conversions = tuple(max(0, state.conversions[j] - used_conversion[j]) + converted[j]
                                  for j in range(2))
        state.seasoned = tuple(max(0, seasoned[j] - used_conversion[j]) for j in range(2))
        state.conversion_history.append(converted)
        state.irmaa_magi = state.irmaa_magi[1], tax.agi + p.exempt
        state.federal_loss, state.ca_loss = tax.federal_loss, tax.ca_loss
        balance = state.balance
        if not all(isfinite(value) for value in (balance, state.basis, tax.total_tax, shortfall)):
            raise ValueError(f'Nonfinite account or cash calculation for {year}.')
        checks = None
        if audit:
            checks = {
                'Projection': {
                    'M': roth_savings, 'P': hsa_paid, 'R': sepp, 'S': rmd, 'T': mandatory, 'U': headroom,
                    'V': draws[0], 'W': draws[1], 'X': sum(access), 'Y': draws[2], 'Z': shortfall,
                    'AA': funding.conversion, 'AC': interest, 'AE': funding.realized_gain,
                    'AF': dividends, 'AI': tax.taxable_ss, 'AK': tax.agi, 'AL': p.year_tax.federal_deduction,
                    'AR': p.payroll, 'AU': p.sdi, 'AV': tax.total_tax, 'AZ': surcharge, 'BA': healthcare,
                    'BB': deposit, 'BC': state.taxable, 'BD': state.basis, 'BE': state.traditional,
                    'BF': sum(state.roth), 'BG': state.hsa, 'BH': balance,
                    'BJ': balance - state.traditional * i[61] - state.payable,
                    'BL': tax.aca_magi, 'BN': balance / p.cpi, 'BW': state.federal_loss,
                    'BY': state.ca_loss, 'CE': removal, 'CF': basis_addition,
                },
                'Cash_Flow': {
                    'C': cash_income, 'D': required, 'E': hsa_paid, 'F': opening_payable,
                    'G': p.roth_target, 'H': opening_taxable,
                    'I': traditional_limit, 'J': roth_limit, 'M': tax.total_tax,
                    'N': paid_tax, 'O': state.payable, 'P': paid_spending, 'Q': unpaid_spending,
                    'R': roth_savings, 'S': p.roth_target - roth_savings, 'T': deposit,
                    'V': shortfall, 'X': savings_limit, 'Y': cash_available, 'AB': p.cash_income,
                    'AC': funding.premium_credit,
                },
                'Healthcare': {
                    'L': surcharge, 'P': healthcare, 'Q': qualified, 'R': funding.premium_credit,
                    'AC': tax.aca_magi, 'AE': hsa_paid,
                },
                'Roth_Access': dict(zip(
                    ('N', 'O', 'P', 'Q', 'R', 'S', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB',
                     'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AN', 'AO'),
                    first_ira_year + access + paid_roth + saved + converted + state.regular
                    + state.conversions + state.seasoned + state.roth + capacity,
                )),
            }
        return YearResult(year, p.age, balance / p.cpi, balance, shortfall / p.cpi, shortfall, cut / p.cpi, checks)


def reconcile(model, base, *, tolerance=SHORTFALL):
    """Compare independently rolled results with the saved annual calculation."""
    discrepancies = []
    comparisons = 0
    max_delta = 0
    for source, result in zip(model.scenario.years, base.years, strict=True):
        for sheet, expected in (
            ('Projection', source.projection), ('Cash_Flow', source.cash),
            ('Healthcare', source.health), ('Roth_Access', source.roth),
        ):
            for column, calculated in result.audit[sheet].items():
                saved = expected[column]
                comparisons += 1
                delta = abs(calculated - saved)
                max_delta = max(max_delta, delta)
                if delta > tolerance:
                    discrepancies.append(f'{sheet}!{column}{source.row}')
    if discrepancies:
        samples = '\n  '.join(discrepancies[:8])
        raise ValueError(f'No-shock reconciliation failed for {len(discrepancies)} values.\n  {samples}')
    return {'years': len(base.years), 'values_compared': comparisons, 'max_difference': max_delta,
            'tolerance_dollars': tolerance}
