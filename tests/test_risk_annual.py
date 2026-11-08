"""Verify the annual risk runner against the packaged release workbook."""

from pathlib import Path
import os
import unittest

from risk_engine.annual import AnnualModel, reconcile
from risk_engine.paths import Guardrail
from risk_engine.report import run_analysis
from risk_engine.scenario import load_scenario


class AnnualRiskTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        default = Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsx'
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
        for policy in ('fixed', 'flexible'):
            result = first[policy]
            self.assertEqual(result['trials'], 3)
            self.assertEqual(sum(result['failure_age_counts'].values()), result['failed_trials'])

    def test_zero_volatility_preserves_base_results_for_both_policies(self):
        result = run_analysis(self.scenario, trials=3, seed=1234, volatility=0,
                              source_sha256=self.scenario.workbook_sha256,
                              workbook_calculation='Test')
        self.assertEqual(result['fixed'], result['flexible'])
        self.assertEqual(result['spending_cuts']['trials_with_cuts'], 0)
        for row in result['annual']:
            self.assertEqual(row['fixed_p10_p50_p90'], [row['base']] * 3)
            self.assertEqual(row['flexible_p10_p50_p90'], [row['base']] * 3)


if __name__ == '__main__':
    unittest.main()
