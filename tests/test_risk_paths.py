"""Check reproducibility, stage risk and spending policy independently of cash funding."""

from dataclasses import replace
from math import isclose, sqrt
import unittest

from risk_engine.paths import Guardrail, ReturnModel, RiskInputs, TrialOutcome, summarize


STOCK = (1.0, 0.0, 0.0)
INTERNATIONAL = (0.0, 1.0, 0.0)
BOND = (0.0, 0.0, 1.0)
INPUTS = RiskInputs(
    asset_returns=(0.075, 0.08, 0.045),
    asset_volatility=(0.2, 0.15, 0.05),
    correlation=((1.0, 0.8, 0.1), (0.8, 1.0, 0.0), (0.1, 0.0, 1.0)),
    work_weights=(STOCK, BOND, INTERNATIONAL, BOND),
    retirement_weights=(BOND, BOND, STOCK, INTERNATIONAL),
)


class ReturnPathTests(unittest.TestCase):
    def setUp(self):
        self.model = ReturnModel(INPUTS)

    def test_zero_shock_returns_workbook_baseline_exactly(self):
        ages = range(54, 57)
        baseline = ((0.06, 0.04, 0.075, 0.075), (0.03, 0.05, 0.07, 0.07), (-1.0, 0, 0, 0))
        for basis in (1, 2):
            self.assertEqual(
                self.model.random_path(ages, 55, 99, 10, 0.0, basis, baseline),
                baseline,
            )
        self.assertEqual(self.model.sequence_path(ages, 55, 0, -0.2, baseline), baseline)
        self.assertEqual(self.model.baseline_path(ages, 55)[0], (0.075, 0.045, 0.08, 0.045))
        self.assertEqual(self.model.baseline_path(ages, 55)[1], (0.045, 0.045, 0.075, 0.08))

    def test_seed_trial_order_and_horizon_extension(self):
        first = self.model.random_path(range(54, 126), 55, 2026, 19)
        different_trial = self.model.random_path(range(54, 126), 55, 2026, 1)
        different_seed = self.model.random_path(range(54, 126), 55, 2027, 19)
        self.assertEqual(len(first), 72)
        self.assertNotEqual(first, different_trial)
        self.assertNotEqual(first, different_seed)
        self.assertEqual(first, self.model.random_path(range(54, 126), 55, 2026, 19))
        self.assertEqual(first[:2], self.model.random_path(range(54, 56), 55, 2026, 19))
        for rates in first:
            self.assertTrue(all(rate >= -1 for rate in rates))
        self.assertEqual(first[0][1], first[0][3])  # same bond draw in the working mix
        self.assertEqual(first[-1][0], first[-1][1])  # same bond draw in the retirement mix

    def test_sequence_changes_equities_and_can_move_the_same_shock(self):
        ages = range(54, 60)
        baseline = self.model.baseline_path(ages, 55)
        early = self.model.sequence_path(ages, 55, 2, -0.25)
        late = self.model.sequence_path(ages, 55, 2, -0.25, delay_years=2)
        self.assertEqual(early[0], baseline[0])
        self.assertEqual(early[1][0:2], baseline[1][0:2])
        self.assertAlmostEqual(early[1][2], -0.25)
        self.assertAlmostEqual(early[1][3], -0.25)
        self.assertEqual(early[1:3], late[3:5])
        self.assertEqual(early[3:], baseline[3:])
        self.assertEqual(late[:3], baseline[:3])
        self.assertEqual(late[5:], baseline[5:])
        self.assertEqual(self.model.sequence_path((55,), 55, 5, -1)[0][2:], (-1, -1))

    def test_covariance_tracks_working_and_retirement_weights(self):
        for basis in (1, 2):
            work = self.model.account_covariance(False, return_basis=basis)
            retired = self.model.account_covariance(True, return_basis=basis)
            self.assertAlmostEqual(sqrt(work[0][0]), 0.2)
            self.assertAlmostEqual(sqrt(retired[0][0]), 0.05)
            self.assertAlmostEqual(work[1][3], 0.05 ** 2)  # identical bond exposure
            self.assertAlmostEqual(retired[0][3], 0)  # independent bonds and international stock
            self.assertAlmostEqual(retired[2][3], retired[3][2])
            self.assertTrue(isclose(retired[2][3], 0.8 * 0.2 * 0.15, rel_tol=0.03))
        self.assertEqual(self.model.account_covariance(True, 0), ((0,) * 4,) * 4)

    def test_perfect_correlation_valid_and_uses_common_shock(self):
        identical = replace(
            INPUTS,
            asset_returns=(0.075, 0.075, 0.045),
            asset_volatility=(0.2, 0.2, 0.05),
            correlation=((1, 1, 0), (1, 1, 0), (0, 0, 1)),
        )
        rates = ReturnModel(identical).random_path((55,), 55, 1, 1)[0]
        self.assertEqual(rates[2], rates[3])

    def test_invalid_assumptions_fail_before_simulation(self):
        for correlation in (
            ((1, 0.8, 0.1), (0.7, 1, 0), (0.1, 0, 1)),
            ((1, 0.9, 0.9), (0.9, 1, -0.9), (0.9, -0.9, 1)),
            ((1, 0, 0), (0, 0.5, 0), (0, 0, 1)),
        ):
            with self.assertRaisesRegex(ValueError, "correlation"):
                ReturnModel(replace(INPUTS, correlation=correlation))
        with self.assertRaisesRegex(ValueError, "100%"):
            ReturnModel(replace(INPUTS, work_weights=((0.2, 0.2, 0.2), BOND, BOND, STOCK)))
        with self.assertRaisesRegex(ValueError, "greater than -100%"):
            ReturnModel(replace(INPUTS, asset_returns=(-1.0, 0, 0)))
        for ages in ((124, 126), (54, 56), (55, 54), (54.0,), (55, 55)):
            with self.assertRaisesRegex(ValueError, "Path ages"):
                self.model.random_path(ages, 55, 10, 0)
        with self.assertRaisesRegex(ValueError, "Retirement age"):
            self.model.baseline_path((55,), 55.0)
        with self.assertRaisesRegex(ValueError, "Baseline returns"):
            self.model.sequence_path((55, 56), 55, 0, -0.2, [(0, 0, 0, 0)])
        with self.assertRaisesRegex(ValueError, "Seed"):
            self.model.random_path((55,), 55, 1 << 64, 0)
        with self.assertRaisesRegex(ValueError, "Trial number"):
            self.model.random_path((55,), 55, 10, -1)
        with self.assertRaisesRegex(ValueError, "Volatility multiplier"):
            self.model.random_path((55,), 55, 10, 0, -0.01)
        with self.assertRaisesRegex(ValueError, "Equity shock"):
            self.model.sequence_path((55,), 55, 1, -1.01)


class SpendingAndReportTests(unittest.TestCase):
    def test_only_flexible_living_spending_is_cut_after_threshold(self):
        policy = Guardrail(0.8, 0.15, 0.7)
        living_spending = 30000
        fixed_bills = 25000
        actual = fixed_bills + living_spending - policy.reduction(True, 79, 100, living_spending)
        self.assertEqual(actual, fixed_bills + 30000 * 0.3 + 30000 * 0.7 * 0.85)
        self.assertEqual(policy.reduction(False, 79, 100, living_spending), 0)
        self.assertEqual(policy.reduction(True, 80, 100, living_spending), 0)
        self.assertEqual(policy.reduction(True, 0, 0, living_spending), 0)
        self.assertEqual(policy.reduction(True, 79, 100, 0), 0)
        self.assertEqual(Guardrail(0.8, 1, 0).reduction(True, 79, 100, living_spending), 0)
        with self.assertRaisesRegex(ValueError, "between 0 and 1"):
            Guardrail(0.8, 0.15, 1.2)

    def test_failure_uses_first_cash_gap_even_if_terminal_assets_recover(self):
        report = summarize((
            TrialOutcome(2, 62, 100, 15),
            TrialOutcome(1, None, 0, 0),
            TrialOutcome(0, 58, 300, 30),
        ))
        self.assertEqual(report.failed_trials, 2)
        self.assertEqual(report.success_rate, 1 / 3)
        self.assertEqual(report.failure_age_counts, {58: 1, 62: 1})
        self.assertEqual(report.median_failure_age, 60)
        self.assertEqual(report.ending_balance_p50, 100)
        self.assertEqual(report.ending_balance_p10, 20)
        self.assertEqual(report.max_cash_shortfall, 30)
        no_failures = summarize((TrialOutcome(0, None, 0, 0),))
        self.assertEqual(no_failures.success_rate, 1)
        self.assertIsNone(no_failures.median_failure_age)


if __name__ == "__main__":
    unittest.main()
