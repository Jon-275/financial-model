"""Test-only counterexample for a claimed earliest funded draw."""

from collections.abc import Callable
from math import isfinite


def earlier_feasible_draw(claimed_draw: float, margin_at: Callable[[float], float]) -> tuple[float, float] | None:
    """Find one lower funded quote using fresh inputs and the same selected policy.

    Pass the gross requested draw and a repricing function bound to the
    current applicable conversion branch that resolves cash margin to half a
    cent. A missing witness is inconclusive. An unfunded selection needs a
    comparison of the applicable grid candidates.
    """
    if not isfinite(claimed_draw) or claimed_draw < .01:
        return None
    lower = claimed_draw - .01
    # Require a reduction outside the solver's half-cent convergence band.
    if claimed_draw - lower <= .005:
        return None
    margin = margin_at(lower)
    return (lower, margin) if isfinite(margin) and margin > .005 else None
