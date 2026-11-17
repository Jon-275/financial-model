"""Verify the annual risk runner against the packaged release workbook."""

from dataclasses import replace
from math import sqrt
from pathlib import Path
import os
from types import SimpleNamespace
import unittest

from risk_engine.annual import AnnualModel, PathResult, YearResult, reconcile
from risk_engine.paths import Guardrail, ReturnModel, RiskInputs, TrialOutcome
from risk_engine.report import _allocations, _representative_trials, _result, run_analysis
from risk_engine.scenario import HSA_RETURN_ROWS, load_scenario


class AnnualRiskTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        default = Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsm'
        cls.scenario = load_scenario(os.environ.get('RISK_TEST_WORKBOOK', default))
        cls.model = AnnualModel(cls.scenario)

    def test_baseline_reconciles_and_guardrail_does_not_cut_it(self):
        model = self.model
        base = model.run(model.baseline_returns, audit=True)
        parity = reconcile(model, base)
        self.assertEqual(parity['years'], len(self.scenario.years))
        self.assertGreater(parity['values_compared'], 50 * parity['years'])
        guarded = model.run(model.baseline_returns, guardrail=Guardrail(1, 1, 1),
                            planned_openings=model.planned_openings(base))
        self.assertEqual(tuple(row.nominal_balance for row in base.years),
                         tuple(row.nominal_balance for row in guarded.years))
        self.assertTrue(all(row.real_cut == 0 for row in guarded.years))

    def test_report_repeats_and_preserves_the_timeline(self):
        def run():
            return run_analysis(self.scenario, trials=3, seed=2**64 - 1, volatility=1,
                                source_sha256=self.scenario.workbook_sha256,
                                workbook_calculation='Test')

        first, second = run(), run()
        self.assertEqual(first, second)
        self.assertEqual(first['provenance']['seed'], str(2**64 - 1))
        self.assertEqual(first['annual'][-1]['age'], self.scenario.inputs[53])
        self.assertAlmostEqual(first['baseline']['opening_balance_real'],
                               self.scenario.opening_projection['BH'], places=2)
        for policy in ('fixed', 'flexible'):
            result = first[policy]
            self.assertEqual(result['trials'], 3)
            self.assertEqual(sum(result['failure_age_counts'].values()), result['failed_trials'])

        representatives = first['representative_paths']
        self.assertEqual([(row['label'], row['percentile']) for row in representatives],
                         [('Low', 10), ('Median', 50), ('High', 90)])
        fixed_ends = [row['fixed']['ending_balance_real'] for row in representatives]
        self.assertEqual(fixed_ends, sorted(fixed_ends))
        allocation = self.scenario.allocation
        returns = ReturnModel(RiskInputs(allocation.means, allocation.volatilities, allocation.correlation,
                                        allocation.working, allocation.retirement))
        baseline = self.model.run(self.model.baseline_returns)
        guardrail = Guardrail(*(self.scenario.sequence[key] for key in (6, 7, 8)))
        for row in representatives:
            path = returns.random_path(self.model.ages, self.model.retirement_age, 2**64 - 1,
                                       row['trial'], 1, self.scenario.return_basis, self.model.baseline_returns)
            fixed = self.model.run(path)
            flexible = self.model.run(path, guardrail=guardrail,
                                      planned_openings=self.model.planned_openings(baseline))
            for policy, expected in [('fixed', fixed), ('flexible', flexible)]:
                self.assertEqual(row[policy]['ending_balance_real'], expected.ending_balance_real)
                self.assertEqual(row[policy]['first_shortfall_age'], expected.first_failure_age)

    def test_zero_volatility_preserves_base_results_for_both_policies(self):
        result = run_analysis(self.scenario, trials=3, seed=1234, volatility=0,
                              source_sha256=self.scenario.workbook_sha256,
                              workbook_calculation='Test')
        self.assertEqual(result['fixed'], result['flexible'])
        self.assertEqual(result['spending_cuts']['trials_with_cuts'], 0)
        self.assertIsNone(result['spending_cuts']['median_years_when_cut'])
        for row in result['annual']:
            self.assertEqual(row['fixed_p10_p50_p90'], [row['base']] * 3)
            self.assertEqual(row['flexible_p10_p50_p90'], [row['base']] * 3)
        for row in result['representative_paths']:
            self.assertEqual(row['fixed'], result['baseline'])
            self.assertEqual(row['flexible'], result['baseline'])

    def test_hsa_export_uses_the_simulated_rate_in_both_stages(self):
        working_rate, retirement_rate = HSA_RETURN_ROWS
        scenario = replace(self.scenario, inputs={**self.scenario.inputs, working_rate: .03, retirement_rate: .02})
        allocation = scenario.allocation
        returns = ReturnModel(RiskInputs(allocation.means, allocation.volatilities, allocation.correlation,
                                        allocation.working, allocation.retirement))
        hsa = [row for row in _allocations(scenario, returns, 1) if row['account'] == 'HSA']
        self.assertEqual([row['base_return'] for row in hsa], [.03, .02])
        for retired, row in enumerate(hsa):
            variance = returns.account_covariance(bool(retired), 1, scenario.return_basis)[3][3]
            gross_return = 1 + sum(w * r for w, r in zip(row['weights'], allocation.means)) / sum(row['weights'])
            self.assertAlmostEqual(row['return_volatility'], sqrt(variance) * (1 + row['base_return']) / gross_return)

        old_inputs = {row: value for row, value in scenario.inputs.items() if row not in HSA_RETURN_ROWS}
        old_report = _allocations(replace(scenario, inputs=old_inputs), returns, 1)
        self.assertEqual([row['base_return'] for row in old_report if row['account'] == 'HSA'],
                         [old_inputs[21], old_inputs[22]])


class FundingDetailTests(unittest.TestCase):
    def test_ranks_observed_paths_and_orders_balance_ties_by_funding(self):
        tied = [
            TrialOutcome(8, None, 100, 0),
            TrialOutcome(4, 70, 100, 8),
            TrialOutcome(6, 55, 100, 10),
            TrialOutcome(0, 55, 100, 12),
            TrialOutcome(5, 55, 100, 12),
        ]
        self.assertEqual(_representative_trials(tied), [('Low', 10, 0), ('Median', 50, 6), ('High', 90, 8)])
        self.assertEqual(_representative_trials(reversed(tied)), _representative_trials(tied))
        self.assertEqual(_representative_trials(tied[:1]), [('Low', 10, 8), ('Median', 50, 8), ('High', 90, 8)])

    def test_first_shortfall_reports_its_own_funding_and_uses_workbook_threshold(self):
        report = _result(PathResult((
            YearResult(2027, 54, 100, 100, 0.009, 0.009, 0),
            YearResult(2028, 55, 70, 140, 2.5, 5, 0, {
                'Cash_Flow': {'C': 50, 'D': 13, 'E': 10, 'F': 10, 'H': 20, 'I': 10, 'J': 8, 'M': 80,
                              'N': 88, 'O': 2, 'P': 0, 'Q': 3},
                'Projection': {'V': 20, 'W': 10, 'Y': 8},
                'Healthcare': {'Q': 10},
            }),
            YearResult(2029, 56, 60, 120, 15, 30, 20),
        ), opening_balance=100), [SimpleNamespace(cpi=2)] * 3)
        self.assertEqual(report['first_shortfall_age'], 55)
        self.assertEqual(report['max_cash_shortfall_real'], 15)
        self.assertEqual(report['first_shortfall'], {
            'year': 2028, 'age': 55, 'total_real': 2.5, 'unpaid_spending_real': 1.5, 'unpaid_tax_real': 1.0,
            'required_spending_real': 6.5, 'tax_due_real': 45.0, 'income_available_real': 25.0,
            'income_used_real': 25.0, 'hsa_used_real': 5.0, 'hsa_eligible_expense_real': 5.0,
            'brokerage_available_real': 10.0, 'brokerage_draw_real': 10.0, 'traditional_available_real': 5.0,
            'traditional_draw_real': 5.0, 'roth_available_real': 4.0, 'roth_draw_real': 4.0,
        })
        self.assertEqual(report['years_with_cuts'], 1)
        self.assertEqual(report['total_cuts_real'], 20)


if __name__ == '__main__':
    unittest.main()
