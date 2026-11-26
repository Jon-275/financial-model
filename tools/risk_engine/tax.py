"""Numeric equivalent of the workbook's annual candidate tax policy."""

from bisect import bisect_right
from collections.abc import Mapping
from dataclasses import dataclass
from math import floor, isfinite
from typing import cast

from state_policy import CALIFORNIA_POLICY, StatePolicy, resolve_state_policy


@dataclass(frozen=True, slots=True)
class TaxInputs:
    """Ordinary income includes net wages and other income, before interest or draws."""

    ordinary_income: float
    qualified_dividends: float = 0.0
    investment_interest: float = 0.0
    tax_exempt_interest: float = 0.0
    social_security: float = 0.0
    hsa_contribution: float = 0.0
    payroll_tax: float = 0.0
    sdi: float = 0.0
    opening_federal_loss: float = 0.0
    opening_ca_loss: float = 0.0


@dataclass(frozen=True, slots=True)
class TaxResult:
    total_tax: float
    agi: float
    taxable_ss: float
    federal_loss: float
    ca_loss: float
    aca_magi: float
    ordinary_before_ss: float
    provisional_income: float
    ordinary_income: float
    ordinary_taxable: float
    federal_ordinary: float
    federal_gains: float
    niit: float
    amt: float
    california_taxable: float
    california_tax: float
    federal_capital_income: float
    ca_capital_income: float
    amt_taxable: float
    amt_preferential: float
    amt_tentative: float
    california_agi: float
    california_exemption_credit: float


@dataclass(frozen=True, slots=True)
class _Brackets:
    floors: tuple[float, ...]
    rates: tuple[float, ...]
    bases: tuple[float, ...]

    @classmethod
    def from_slopes(cls, floors: tuple[float, ...], slopes: tuple[float, ...]) -> "_Brackets":
        """Accumulate the workbook's incremental slopes at the bracket floors."""
        if floors[0] != 0 or any(a > b for a, b in zip(floors, floors[1:])):
            raise ValueError("Tax bracket floors must start at zero and ascend.")
        rates = [slopes[0]]
        bases = [0.0]
        for lo, hi, slope in zip(floors, floors[1:], slopes[1:]):
            bases.append(bases[-1] + (hi - lo) * rates[-1])
            rates.append(rates[-1] + slope)
        return cls(floors, tuple(rates), tuple(bases))

    def tax(self, amount: float) -> float:
        band = bisect_right(self.floors, amount) - 1
        return self.bases[band] + (amount - self.floors[band]) * self.rates[band]


def _closing_loss(realized: float, opening: float, capital_income: float, agi: float, deduction: float) -> float:
    claim = max(0.0, -capital_income)
    used = min(claim, max(0.0, agi - deduction + claim))
    return max(0.0, opening - realized - used)


@dataclass(frozen=True, slots=True)
class YearTax:
    federal_deduction: float
    california_deduction: float
    loss_limit: float
    _ordinary: _Brackets
    _gains: _Brackets
    _california: _Brackets
    _ss_low: float
    _ss_high: float
    _ss_low_rate: float
    _ss_high_rate: float
    _niit_rate: float
    _niit_threshold: float
    _amt_exemption: float
    _amt_phaseout: float
    _amt_phaseout_rate: float
    _amt_threshold: float
    _amt_rate: float
    _amt_upper_rate: float
    _ca_surtax_threshold: float
    _ca_surtax_rate: float
    _ca_credit: float
    _ca_credit_count: int
    _ca_credit_threshold: float
    _ca_credit_reduction: float
    _ca_credit_step: float
    _payroll_wage_cap: float
    _payroll_ss_rate: float
    _payroll_medicare_rate: float
    _payroll_add_medicare_rate: float
    _payroll_add_medicare_threshold: float
    _sdi_rate: float
    _aca_bands: tuple[tuple[float, float, float, float], ...]

    def payroll_taxes(self, primary_wages: float, spouse_wages: float) -> tuple[float, float]:
        """Federal payroll tax and California SDI for this year's wages."""
        wages = primary_wages + spouse_wages
        ss_wages = min(primary_wages, self._payroll_wage_cap) + min(spouse_wages, self._payroll_wage_cap)
        payroll = (
            self._payroll_ss_rate * ss_wages
            + self._payroll_medicare_rate * wages
            + self._payroll_add_medicare_rate * max(0.0, wages - self._payroll_add_medicare_threshold)
        )
        return payroll, wages * self._sdi_rate

    def premium_credit(self, magi: float, actual: float, benchmark: float, fpl: float) -> float:
        """ACA contribution bands, with inclusive 100% and 400% boundaries."""
        if actual <= 0 or benchmark <= 0 or fpl <= 0 or magi < fpl or magi > 4 * fpl:
            return 0.0
        ratio = magi / fpl
        if ratio < self._aca_bands[0][1]:
            rate = self._aca_bands[0][2]
        else:
            rate = self._aca_bands[-1][2]
            for start, end, low, high in self._aca_bands[1:-1]:
                if ratio < end:
                    rate = low + (ratio - start) * (high - low) / (end - start)
                    break
        return min(actual, max(0.0, benchmark - magi * rate))

    def _amt_rates(self, amount: float) -> float:
        return self._amt_rate * min(amount, self._amt_threshold) + self._amt_upper_rate * max(
            0.0, amount - self._amt_threshold
        )

    def evaluate(
        self, inputs: TaxInputs, traditional_draw: float, conversion: float, realized_gain: float
    ) -> TaxResult:
        return cast(TaxResult, self._candidate(inputs, traditional_draw, conversion, realized_gain, True))

    def cash_tax(
        self, inputs: TaxInputs, traditional_draw: float, conversion: float, realized_gain: float
    ) -> tuple[float, float]:
        """Return the current tax and ACA MAGI used for each funding trial."""
        return cast(
            tuple[float, float], self._candidate(inputs, traditional_draw, conversion, realized_gain, False)
        )

    def _candidate(
        self, inputs: TaxInputs, traditional_draw: float, conversion: float, realized_gain: float, detail: bool
    ) -> TaxResult | tuple[float, float]:
        dividends = inputs.qualified_dividends
        social = inputs.social_security
        interest = inputs.investment_interest
        exempt = inputs.tax_exempt_interest
        ordinary_before_ss = inputs.ordinary_income + traditional_draw + conversion + interest
        fed_cap = max(-self.loss_limit, realized_gain - inputs.opening_federal_loss)
        ca_cap = max(-self.loss_limit, realized_gain - inputs.opening_ca_loss)
        provisional = ordinary_before_ss + fed_cap + dividends + exempt + 0.5 * social
        taxable_ss = min(
            self._ss_high_rate * social,
            min(
                self._ss_low_rate * min(max(0.0, provisional - self._ss_low), self._ss_high - self._ss_low),
                self._ss_low_rate * social,
            )
            + self._ss_high_rate * max(0.0, provisional - self._ss_high),
        ) if social else 0.0

        positive_gains = max(0.0, fed_cap)
        ordinary = ordinary_before_ss + taxable_ss + min(0.0, fed_cap)
        agi = ordinary + positive_gains + dividends
        deduction = self.federal_deduction
        ordinary_taxable = max(0.0, ordinary - deduction)
        federal_ordinary = self._ordinary.tax(ordinary_taxable) if ordinary_taxable else 0.0
        taxable = max(0.0, agi - deduction)
        gain_taxable = max(0.0, agi - deduction - ordinary_taxable)
        federal_gains = 0.0
        if gain_taxable:
            stacked = self._gains.tax(ordinary_taxable + gain_taxable) - self._gains.tax(ordinary_taxable)
            federal_gains = max(0.0, min(stacked, self._ordinary.tax(taxable) - federal_ordinary))

        niit = self._niit_rate * min(
            max(0.0, interest + dividends + fed_cap), max(0.0, agi - self._niit_threshold)
        ) if agi > self._niit_threshold else 0.0
        amt_taxable = max(
            0.0, agi - max(0.0, self._amt_exemption - self._amt_phaseout_rate * max(0.0, agi - self._amt_phaseout))
        )
        amt_preferential = min(amt_taxable, positive_gains + dividends)
        tentative = 0.0
        if amt_taxable:
            tentative = self._amt_rates(amt_taxable)
            if amt_preferential:
                stacked = self._gains.tax(ordinary_taxable + amt_preferential) - self._gains.tax(ordinary_taxable)
                tentative = min(tentative, self._amt_rates(amt_taxable - amt_preferential) + stacked)
        amt = max(0.0, tentative - federal_ordinary - federal_gains)

        ca_agi = ordinary_before_ss + inputs.hsa_contribution + ca_cap + dividends
        ca_taxable = max(0.0, ca_agi - self.california_deduction)
        ca_reduction = (
            self._ca_credit_reduction * max(0.0, agi - self._ca_credit_threshold) / self._ca_credit_step
        )
        ca_credit = max(0.0, self._ca_credit - self._ca_credit_count * ca_reduction)
        ca_tax = max(
            0.0,
            self._california.tax(ca_taxable)
            + max(0.0, ca_taxable - self._ca_surtax_threshold) * self._ca_surtax_rate
            - ca_credit,
        ) if ca_taxable else 0.0

        total = federal_ordinary + federal_gains + niit + amt + inputs.payroll_tax + ca_tax + inputs.sdi
        aca_magi = agi + social - taxable_ss + exempt
        if not detail:
            return total, aca_magi
        return TaxResult(
            total_tax=total,
            agi=agi,
            taxable_ss=taxable_ss,
            federal_loss=_closing_loss(realized_gain, inputs.opening_federal_loss, fed_cap, agi, deduction),
            ca_loss=_closing_loss(realized_gain, inputs.opening_ca_loss, ca_cap, ca_agi, self.california_deduction),
            aca_magi=aca_magi,
            ordinary_before_ss=ordinary_before_ss,
            provisional_income=provisional,
            ordinary_income=ordinary,
            ordinary_taxable=ordinary_taxable,
            federal_ordinary=federal_ordinary,
            federal_gains=federal_gains,
            niit=niit,
            amt=amt,
            california_taxable=ca_taxable,
            california_tax=ca_tax,
            federal_capital_income=fed_cap,
            ca_capital_income=ca_cap,
            amt_taxable=amt_taxable,
            amt_preferential=amt_preferential,
            amt_tentative=tentative,
            california_agi=ca_agi,
            california_exemption_credit=ca_credit,
        )


@dataclass(frozen=True, slots=True)
class TaxTables:
    """Read the same assumptions used by the saved Tax_Tables sheet."""

    _cells: Mapping[str, object]
    state_policy: StatePolicy = CALIFORNIA_POLICY

    def __post_init__(self) -> None:
        """Protect the policy version from a conflicting metadata record."""
        if not isinstance(self.state_policy, StatePolicy):
            raise ValueError('State tax policy metadata is invalid.')
        supported = resolve_state_policy(self.state_policy.jurisdiction, self.state_policy.version)
        if self.state_policy != supported:
            raise ValueError('State tax policy metadata does not match its version.')

    @classmethod
    def from_cells(
        cls, cells: Mapping[str, object], *, state_policy: StatePolicy = CALIFORNIA_POLICY
    ) -> "TaxTables":
        """Copy the worksheet values and associate the validated state policy."""
        return cls(dict(cells), state_policy)

    def for_year(
        self, year: int, inflation: float, tax_index: float, single: bool, primary_age: int, spouse_age: int
    ) -> YearTax:
        def num(cell: str) -> float:
            value = self._cells.get(cell)
            if isinstance(value, bool) or not isinstance(value, (int, float)) or not isfinite(value):
                raise ValueError(f"Tax_Tables!{cell} needs a finite numeric value.")
            return float(value)

        def b(row: int) -> float:
            return num(f"B{row}")

        published = b(219) == 1
        fed_published = published and year >= 2027
        ca_published = published and year >= self.state_policy.published_income_year
        sdi_published = published and year >= self.state_policy.published_sdi_year and 'B251' in self._cells
        fed_factor = tax_index / (1 + inflation) if fed_published else tax_index

        def fed(old: int, new: int) -> float:
            return b(new if fed_published else old) * fed_factor

        def ca(old: int, new: int) -> float:
            return b(new if ca_published else old) * tax_index

        def brackets(start: int, length: int, new: bool, factor: float) -> _Brackets:
            col = "E" if new else "C"
            rows = range(start, start + length)
            return _Brackets.from_slopes(
                tuple(num(f"{col}{row}") * factor for row in rows),
                tuple(num(f"D{row}") for row in rows),
            )

        seniors = int(primary_age >= 65) + int(not single and spouse_age >= 65)
        federal_deduction = fed(171, 232) + seniors * fed(172, 234) if single else (
            fed(23, 231) + seniors * fed(24, 233)
        )
        california_deduction = ca(178, 242) if single else ca(42, 241)
        ca_personal_credit = ca(179, 244) if single else ca(43, 243)
        ca_credit_threshold = b(193 if single else 192)
        if ca_published:
            ca_credit_threshold = floor(ca_credit_threshold * 1.034 + 0.5)
        ca_credit_threshold *= tax_index
        aca_first = 221 if year >= 2027 else 204
        aca_bands = tuple(
            (num(f"A{r}"), num(f"B{r}"), num(f"C{r}"), num(f"D{r}"))
            for r in range(aca_first, aca_first + 6)
        )

        return YearTax(
            federal_deduction=federal_deduction,
            california_deduction=california_deduction,
            loss_limit=b(190),
            _ordinary=brackets(144 if single else 6, 7, fed_published, fed_factor),
            _gains=brackets(154 if single else 17, 3, fed_published, fed_factor),
            _california=brackets(160 if single else 30, 9, ca_published, tax_index),
            _ss_low=b(176 if single else 135),
            _ss_high=b(177 if single else 136),
            _ss_low_rate=b(137),
            _ss_high_rate=b(138),
            _niit_rate=b(74),
            _niit_threshold=b(173 if single else 75),
            _amt_exemption=fed(174, 236) if single else fed(80, 235),
            _amt_phaseout=fed(175, 238) if single else fed(81, 237),
            _amt_phaseout_rate=b(82),
            _amt_threshold=fed(85, 239),
            _amt_rate=b(83),
            _amt_upper_rate=b(84),
            _ca_surtax_threshold=b(45),
            _ca_surtax_rate=b(44),
            _ca_credit=ca_personal_credit + ca(179, 244) * seniors,
            _ca_credit_count=(1 if single else 2) + seniors,
            _ca_credit_threshold=ca_credit_threshold,
            _ca_credit_reduction=b(195),
            _ca_credit_step=b(194),
            _payroll_wage_cap=fed(51, 240),
            _payroll_ss_rate=b(50),
            _payroll_medicare_rate=b(52),
            _payroll_add_medicare_rate=b(53),
            _payroll_add_medicare_threshold=b(191 if single else 54),
            _sdi_rate=b(251 if sdi_published else 46),
            _aca_bands=aca_bands,
        )
