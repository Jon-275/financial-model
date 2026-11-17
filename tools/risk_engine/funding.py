"""Annual draw search shared by deterministic and risk-path projections."""

from dataclasses import dataclass

from .tax import TaxInputs, TaxResult, YearTax


@dataclass(frozen=True, slots=True)
class FundingInputs:
    """Annual cash need, account access, and the selected withdrawal policy."""

    cash_income: float
    required_spending: float
    hsa_paid: float
    opening_tax_payable: float
    taxable_available: float
    traditional_available: float
    roth_available: float
    mandatory: float
    headroom: float
    opening_traditional: float
    basis_ratio: float
    requested_conversion: float
    order: int
    conversion_allowed: bool
    aca_actual: float = 0.0
    aca_benchmark: float = 0.0
    aca_fpl: float = 0.0


@dataclass(frozen=True, slots=True)
class FundingResult:
    requested_draw: float
    lower_draw: float
    taxable_draw: float
    traditional_draw: float
    roth_draw: float
    conversion: float
    realized_gain: float
    premium_credit: float
    margin: float
    tax: TaxResult
    status: str


class _FundingSearch:
    def __init__(self, tax: YearTax, income: TaxInputs, cash: FundingInputs):
        self.tax = tax
        self.income = income
        self.cash = cash
        self.capacity = cash.taxable_available + cash.traditional_available + cash.roth_available
        room = max(0.0, cash.traditional_available - cash.mandatory)
        preferred = min(room, cash.headroom) if cash.order == 2 else 0.0
        self.broker_low = (cash.traditional_available if preferred >= room
                           else cash.mandatory + preferred)
        self.broker_high = self.broker_low + cash.taxable_available
        self.traditional_high = cash.traditional_available + cash.taxable_available
        self.gain_share = 1.0 - cash.basis_ratio
        self.aca = cash.aca_actual > 0 and cash.aca_benchmark > 0 and cash.aca_fpl > 0
        self.conversion = cash.conversion_allowed and cash.requested_conversion > 0 and cash.opening_traditional > 0

    def _draws(self, draw: float, convert: bool) -> tuple[float, float, float, float]:
        cash = self.cash
        taxable = (cash.taxable_available if draw >= self.broker_high
                   else min(cash.taxable_available, max(0.0, draw - self.broker_low)))
        if draw >= self.traditional_high:
            traditional = cash.traditional_available
        elif draw <= self.broker_high:
            traditional = min(max(draw, cash.mandatory), self.broker_low)
        else:
            traditional = min(cash.traditional_available, max(self.broker_low, draw - cash.taxable_available))
        roth = (cash.roth_available if draw >= self.capacity
                else min(cash.roth_available, max(0.0, draw - self.traditional_high)))
        conversion = 0.0
        if convert:
            conversion = min(max(0.0, cash.opening_traditional - traditional), cash.requested_conversion)
            if cash.order == 2:
                conversion = min(conversion, max(0.0, cash.headroom - max(0.0, traditional - cash.mandatory)))
        return taxable, traditional, roth, conversion

    def _credit(self, magi: float) -> float:
        cash = self.cash
        return self.tax.premium_credit(magi, cash.aca_actual, cash.aca_benchmark, cash.aca_fpl) if self.aca else 0.0

    def _margin(self, taxable: float, traditional: float, roth: float, tax: float, credit: float) -> float:
        cash = self.cash
        return (
            cash.cash_income + (taxable + traditional + roth) - cash.required_spending + cash.hsa_paid
            - cash.opening_tax_payable - tax + credit
        )

    def _quote(self, draw: float, convert: bool) -> tuple[float, float]:
        taxable, traditional, roth, conversion = self._draws(draw, convert)
        tax, magi = self.tax.cash_tax(self.income, traditional, conversion, taxable * self.gain_share)
        return self._margin(taxable, traditional, roth, tax, self._credit(magi)), magi

    def _structural_nodes(self) -> list[float]:
        cash = self.cash

        def draw_at_traditional(target: float) -> float:
            withdrawal = max(cash.mandatory, min(cash.traditional_available, target))
            return withdrawal if withdrawal <= self.broker_low else withdrawal + cash.taxable_available

        # Also partition capped conversions and the signed capital-loss line.
        loss_boundary = self.broker_low
        if self.gain_share:
            loss_boundary += max(0.0, min(
                cash.taxable_available, (self.income.opening_federal_loss - self.tax.loss_limit) / self.gain_share
            ))
        return sorted({
            0.0, self.broker_low, self.broker_high, self.traditional_high,
            self.capacity, draw_at_traditional(cash.opening_traditional - cash.requested_conversion),
            draw_at_traditional(cash.mandatory + cash.headroom - cash.requested_conversion),
            draw_at_traditional(cash.mandatory + cash.headroom), loss_boundary,
        })

    def _aca_nodes(self, structural: list[float], magis: list[float]) -> set[float]:
        nodes = set()
        intervals = tuple(zip(structural, structural[1:], magis, magis[1:]))
        phases = ((0.0, self.broker_low), (self.broker_low, self.broker_high), (self.broker_high, self.capacity))
        fpl = self.cash.aca_fpl
        for low, high in phases:
            if high <= low:
                continue
            for cutoff, inclusive_upper in ((fpl, False), (1.33 * fpl, False), (4 * fpl, True)):
                for a, b, magi_a, magi_b in intervals:
                    if a < low or b > high:
                        continue
                    lower_magi, upper_magi = sorted((magi_a, magi_b))
                    # 400% belongs to the subsidized side of the upper cliff.
                    crosses = lower_magi <= cutoff < upper_magi if inclusive_upper else lower_magi < cutoff <= upper_magi
                    if crosses:
                        point = a + (cutoff - magi_a) * (b - a) / (magi_b - magi_a)
                        epsilon = max(0.00001, abs(point) * 2e-15)
                        nodes.update((max(low, point - epsilon), point, min(high, point + epsilon)))
                        # The workbook uses the first crossing within each phase.
                        break
        return nodes

    def _grid(self, convert: bool) -> dict[float, float]:
        margin_zero, magi_zero = self._quote(0.0, convert)
        grid = {0.0: margin_zero}
        if margin_zero >= 0 or self.capacity == 0:
            return grid
        if not self.aca:
            grid[self.capacity] = self._quote(self.capacity, convert)[0]
            return grid
        structural = self._structural_nodes()
        magis = [magi_zero]
        for draw in structural[1:]:
            grid[draw], magi = self._quote(draw, convert)
            magis.append(magi)
        for draw in self._aca_nodes(structural, magis) - grid.keys():
            grid[draw] = self._quote(draw, convert)[0]
        return grid

    def solve(self, exact: bool) -> FundingResult:
        convert = self.conversion
        grid = self._grid(convert)
        upper = min((draw for draw, margin in grid.items() if margin >= 0), default=None)
        if upper is None and convert and not exact:
            convert = False
            grid = self._grid(convert)
            upper = min((draw for draw, margin in grid.items() if margin >= 0), default=None)

        lower = 0.0
        searched = False
        if upper is None:
            # A credit cliff can make a partial draw preferable to full capacity.
            best_margin = max(grid.values())
            upper = min(draw for draw, margin in grid.items() if margin >= best_margin - 1e-7)
        elif upper > 0:
            lower = max(draw for draw in grid if draw < upper)
            searched = True
            for _ in range(49):
                middle = (lower + upper) / 2
                if middle == lower or middle == upper:
                    break
                if self._quote(middle, convert)[0] < 0:
                    lower = middle
                else:
                    upper = middle

        taxable, traditional, roth, conversion = self._draws(upper, convert)
        realized = taxable * self.gain_share
        tax = self.tax.evaluate(self.income, traditional, conversion, realized)
        credit = self._credit(tax.aca_magi)
        margin = self._margin(taxable, traditional, roth, tax.total_tax, credit)
        if searched and (upper - lower > 0.005 or margin < -0.005):
            status = "Check funding convergence"
        else:
            status = "Unfunded" if margin < -0.009999 else "Funded"
        return FundingResult(upper, lower, taxable, traditional, roth, conversion, realized, credit, margin, tax, status)


def solve_funding(
    year_tax: YearTax, tax_inputs: TaxInputs, funding_inputs: FundingInputs, *, exact: bool = False
) -> FundingResult:
    """Choose the first feasible interval, then refine the annual cash balance.

    A regular projection skips a conversion when its cash need is infeasible.
    Conversion comparisons pass ``exact=True`` to retain the requested policy.
    Both policies use the workbook's same fixed bisection limit.
    """
    return _FundingSearch(year_tax, tax_inputs, funding_inputs).solve(exact)
