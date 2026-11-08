"""Return paths, flexible spending cuts and path-level risk statistics."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from functools import lru_cache
from math import exp, expm1, floor, fsum, isclose, isfinite, log1p, sqrt
from random import Random
from statistics import median
from typing import Iterable, Sequence


AccountReturns = tuple[float, float, float, float]
Weights = tuple[tuple[float, ...], ...]
ReturnPath = tuple[AccountReturns, ...]
_ACCOUNT_COUNT = 4  # taxable, traditional, Roth, HSA
_MAX_AGE = 125
_MAX_SEED = (1 << 64) - 1
_MATRIX_TOLERANCE = 1e-12


@dataclass(frozen=True)
class RiskInputs:
    asset_returns: tuple[float, ...]
    asset_volatility: tuple[float, ...]
    correlation: tuple[tuple[float, ...], ...]
    work_weights: Weights
    retirement_weights: Weights
    equity_assets: tuple[int, ...] = (0, 1)


def _correlation_factor(
    matrix: Sequence[Sequence[float]], size: int
) -> tuple[tuple[float, ...], ...]:
    if len(matrix) != size or any(len(row) != size for row in matrix):
        raise ValueError("Asset correlation matrix must be square and match the asset count")
    for i, row in enumerate(matrix):
        for j, value in enumerate(row):
            if not isfinite(value) or not -1 <= value <= 1:
                raise ValueError("Asset correlations must be between -1 and 1")
            expected = 1 if i == j else matrix[j][i]
            if not isclose(value, expected, rel_tol=0, abs_tol=_MATRIX_TOLERANCE):
                raise ValueError("Asset correlations must be symmetric with 1 on the diagonal")

    # A correlation of 1 or -1 can make the matrix singular and is valid.
    lower = [[0.0] * (i + 1) for i in range(size)]
    for i, row in enumerate(lower):
        for j in range(i + 1):
            residual = matrix[i][j] - fsum(row[k] * lower[j][k] for k in range(j))
            if i == j and residual >= -_MATRIX_TOLERANCE:
                row[j] = sqrt(max(0.0, residual))
            elif i > j and lower[j][j] != 0:
                row[j] = residual / lower[j][j]
            elif i == j or abs(residual) > _MATRIX_TOLERANCE:
                raise ValueError("Asset correlation matrix must be positive semidefinite")
    return tuple(tuple(row) for row in lower)


def _allocation(weights: Weights, size: int, stage: str) -> Weights:
    if len(weights) != _ACCOUNT_COUNT:
        raise ValueError(f"{stage} allocation must have four account mixes")
    mixes = []
    for mix in weights:
        if len(mix) != size or any(not isfinite(weight) or weight < 0 for weight in mix):
            raise ValueError(f"{stage} allocation must contain a nonnegative weight for each asset")
        total = fsum(mix)
        if abs(total - 1) >= 0.001:
            raise ValueError(f"{stage} allocation weights must total 100% for each account")
        mixes.append(tuple(weight / total for weight in mix))
    return tuple(mixes)


def _mix_returns(weights: Weights, asset_values: Sequence[float]) -> AccountReturns:
    return tuple(fsum(weight * value for weight, value in zip(mix, asset_values)) for mix in weights)


class ReturnModel:
    """Sample common asset returns and apply each account's current allocation."""

    def __init__(self, inputs: RiskInputs):
        self.inputs = inputs
        size = len(inputs.asset_returns)
        if not size or any(not isfinite(rate) or rate <= -1 for rate in inputs.asset_returns):
            raise ValueError("Baseline asset returns must be finite and greater than -100%")
        if len(inputs.asset_volatility) != size or any(
            not isfinite(volatility) or volatility < 0 for volatility in inputs.asset_volatility
        ):
            raise ValueError("Asset volatility must be finite and nonnegative for each asset")
        if len(set(inputs.equity_assets)) != len(inputs.equity_assets) or any(
            type(index) is not int or not 0 <= index < size for index in inputs.equity_assets
        ):
            raise ValueError("Equity asset indexes must be unique and refer to defined assets")

        self._factor = _correlation_factor(inputs.correlation, size)
        self._weights = (
            _allocation(inputs.work_weights, size, "Working"),
            _allocation(inputs.retirement_weights, size, "Retirement"),
        )
        self._asset_gross = tuple(1 + rate for rate in inputs.asset_returns)
        self._profile_gross = tuple(_mix_returns(weights, self._asset_gross) for weights in self._weights)
        self._profile_returns = tuple(
            _mix_returns(weights, inputs.asset_returns) for weights in self._weights
        )

    def baseline_path(
        self,
        ages: Sequence[int],
        retirement_age: int,
        baseline_returns: Sequence[AccountReturns] | None = None,
    ) -> ReturnPath:
        if type(retirement_age) is not int or not 0 <= retirement_age <= _MAX_AGE:
            raise ValueError("Retirement age must be a whole number from 0 to 125")
        if any(type(age) is not int or not 0 <= age <= _MAX_AGE for age in ages) or any(
            age != previous + 1 for previous, age in zip(ages, ages[1:])
        ):
            raise ValueError("Path ages must be consecutive whole numbers from 0 to 125")
        if baseline_returns is None:
            return tuple(self._profile_returns[age >= retirement_age] for age in ages)
        if len(baseline_returns) != len(ages) or any(
            len(rates) != _ACCOUNT_COUNT or any(not isfinite(rate) or rate < -1 for rate in rates)
            for rates in baseline_returns
        ):
            raise ValueError("Baseline returns must include four rates of at least -100% per age")
        return tuple(tuple(rates) for rates in baseline_returns)

    def _apply(
        self, retired: bool, baseline: AccountReturns, asset_gross: Sequence[float]
    ) -> AccountReturns:
        # Scale each deterministic account gross return by the change in its
        # mix. This preserves workbook return overrides and the -100% floor.
        realized = _mix_returns(self._weights[retired], asset_gross)
        rates = tuple(
            rate if value == expected else (value / expected) * (1 + rate) - 1
            for value, expected, rate in zip(realized, self._profile_gross[retired], baseline)
        )
        if any(not isfinite(rate) for rate in rates):
            raise ValueError("Account return exceeds the supported numeric range")
        return rates

    @lru_cache(maxsize=32)
    def _log_parameters(
        self, volatility_multiplier: float, return_basis: int
    ) -> tuple[tuple[float, float], ...]:
        if not isfinite(volatility_multiplier) or volatility_multiplier < 0:
            raise ValueError("Volatility multiplier must be finite and nonnegative")
        if type(return_basis) is not int or return_basis not in (1, 2):
            raise ValueError("Return basis must be 1 (compound growth) or 2 (arithmetic mean)")

        parameters = []
        for gross, volatility in zip(self._asset_gross, self.inputs.asset_volatility):
            relative_sd = volatility / gross * volatility_multiplier
            relative_variance = relative_sd * relative_sd
            if not isfinite(relative_variance):
                raise ValueError("Scaled asset volatility is too large")
            # Calibrate the normal scale to the asset's simple-return SD.
            # For compound growth, gross is the median instead of the mean.
            if return_basis == 1:
                log_variance = log1p(relative_variance / (0.5 + sqrt(0.25 + relative_variance)))
                shift = 0.0
            else:
                log_variance = log1p(relative_variance)
                shift = -log_variance / 2
            parameters.append((sqrt(log_variance), shift))
        return tuple(parameters)

    def random_path(
        self,
        ages: Sequence[int],
        retirement_age: int,
        seed: int,
        trial: int,
        volatility_multiplier: float = 1.0,
        return_basis: int = 1,
        baseline_returns: Sequence[AccountReturns] | None = None,
    ) -> ReturnPath:
        for value, name in ((seed, "Seed"), (trial, "Trial number")):
            if type(value) is not int or not 0 <= value <= _MAX_SEED:
                raise ValueError(f"{name} must be an unsigned 64-bit integer")
        parameters = self._log_parameters(volatility_multiplier, return_basis)
        baseline = self.baseline_path(ages, retirement_age, baseline_returns)
        if not any(scale for scale, _ in parameters):
            return baseline

        # Distinct seed pairs give each trial its own stream. Requesting
        # another trial or appending years does not change earlier draws.
        rng = Random(seed << 64 | trial)
        result = []
        for age, rates in zip(ages, baseline):
            independent = [rng.gauss(0, 1) for _ in parameters]
            asset_gross = tuple(
                gross * exp(shift + scale * fsum(a * b for a, b in zip(row, independent)))
                for gross, (scale, shift), row in zip(self._asset_gross, parameters, self._factor)
            )
            result.append(self._apply(age >= retirement_age, rates, asset_gross))
        return tuple(result)

    def sequence_path(
        self,
        ages: Sequence[int],
        retirement_age: int,
        shock_years: int,
        equity_return: float,
        baseline_returns: Sequence[AccountReturns] | None = None,
        *,
        delay_years: int = 0,
    ) -> ReturnPath:
        for value, name in ((shock_years, "Shock years"), (delay_years, "Delay years")):
            if type(value) is not int or value < 0:
                raise ValueError(f"{name} must be a nonnegative integer")
        if not isfinite(equity_return) or equity_return < -1:
            raise ValueError("Equity shock return must be finite and at least -100%")
        baseline = self.baseline_path(ages, retirement_age, baseline_returns)
        if shock_years == 0:
            return baseline
        asset_gross = tuple(
            1 + equity_return if i in self.inputs.equity_assets else gross
            for i, gross in enumerate(self._asset_gross)
        )
        if asset_gross == self._asset_gross:
            return baseline
        first_shock_age = retirement_age + delay_years
        return tuple(
            self._apply(True, rates, asset_gross)
            if first_shock_age <= age < first_shock_age + shock_years
            else rates
            for age, rates in zip(ages, baseline)
        )

    def account_covariance(
        self, retired: bool, volatility_multiplier: float = 1.0, return_basis: int = 1
    ) -> tuple[tuple[float, ...], ...]:
        """Theoretical simple-return covariance under the selected profile."""
        parameters = self._log_parameters(volatility_multiplier, return_basis)
        gross_means = tuple(
            gross * exp(shift + scale * scale / 2)
            for gross, (scale, shift) in zip(self._asset_gross, parameters)
        )
        asset_covariance = tuple(
            tuple(
                gross_i * gross_j * expm1(scale_i * scale_j * rho)
                for gross_j, (scale_j, _), rho in zip(gross_means, parameters, row)
            )
            for gross_i, (scale_i, _), row in zip(gross_means, parameters, self.inputs.correlation)
        )
        weights = self._weights[retired]
        return tuple(
            tuple(
                fsum(
                    left * right * covariance
                    for left, row in zip(left_mix, asset_covariance)
                    for right, covariance in zip(right_mix, row)
                )
                for right_mix in weights
            )
            for left_mix in weights
        )


@dataclass(frozen=True)
class Guardrail:
    trigger: float
    cut: float
    flexible_share: float

    def __post_init__(self):
        if any(not isfinite(value) or not 0 <= value <= 1 for value in (
            self.trigger, self.cut, self.flexible_share
        )):
            raise ValueError("Guardrail trigger, cut and flexible share must be between 0 and 1")

    def reduction(
        self,
        retired: bool,
        opening_balance: float,
        planned_opening_balance: float,
        living_spending: float,
    ) -> float:
        if any(not isfinite(amount) or amount < 0 for amount in (
            opening_balance, planned_opening_balance, living_spending
        )):
            raise ValueError("Guardrail balances and living spending must be finite and nonnegative")
        if retired and opening_balance < self.trigger * planned_opening_balance:
            return living_spending * self.flexible_share * self.cut
        return 0.0


@dataclass(frozen=True)
class TrialOutcome:
    trial: int
    first_failure_age: int | None
    ending_balance_real: float
    max_cash_shortfall: float


@dataclass(frozen=True)
class RiskSummary:
    trials: int
    failed_trials: int
    success_rate: float
    median_failure_age: float | None
    failure_age_counts: dict[int, int]
    ending_balance_p10: float
    ending_balance_p25: float
    ending_balance_p50: float
    ending_balance_p75: float
    ending_balance_p90: float
    max_cash_shortfall: float


def percentile(sorted_values: Sequence[float], fraction: float) -> float:
    position = fraction * (len(sorted_values) - 1)
    index = floor(position)
    left = sorted_values[index]
    right = sorted_values[min(index + 1, len(sorted_values) - 1)]
    weight = position - index
    return left + (right - left) * weight


def summarize(trials: Iterable[TrialOutcome]) -> RiskSummary:
    """Use each trial's first funding failure, even if assets later recover."""
    outcomes = tuple(trials)
    if not outcomes:
        raise ValueError("At least one trial outcome is required")
    if len({outcome.trial for outcome in outcomes}) != len(outcomes):
        raise ValueError("Trial outcome identifiers must be unique")
    if any(
        not isfinite(outcome.ending_balance_real)
        or not isfinite(outcome.max_cash_shortfall)
        or outcome.max_cash_shortfall < 0
        for outcome in outcomes
    ):
        raise ValueError("Trial balances and nonnegative cash shortfalls must be finite")
    failures = [outcome.first_failure_age for outcome in outcomes if outcome.first_failure_age is not None]
    balances = sorted(outcome.ending_balance_real for outcome in outcomes)
    p10, p25, p50, p75, p90 = (percentile(balances, fraction) for fraction in (.1, .25, .5, .75, .9))
    return RiskSummary(
        trials=len(outcomes),
        failed_trials=len(failures),
        success_rate=(len(outcomes) - len(failures)) / len(outcomes),
        median_failure_age=median(failures) if failures else None,
        failure_age_counts=dict(sorted(Counter(failures).items())),
        ending_balance_p10=p10,
        ending_balance_p25=p25,
        ending_balance_p50=p50,
        ending_balance_p75=p75,
        ending_balance_p90=p90,
        max_cash_shortfall=max(outcome.max_cash_shortfall for outcome in outcomes),
    )
