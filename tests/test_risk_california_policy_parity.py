"""Check the California v2 parity matrix without running a spreadsheet engine."""

from pathlib import Path
import unittest

from california_policy_parity import HISTORICAL_CASES, plan_cases, select_cases, verify_candidate


class CaliforniaParityTests(unittest.TestCase):
    def test_existing_cases_retain_independent_expectations(self) -> None:
        """Include the historical tax fixtures and an inspected no-shock trial."""
        shared, cases = plan_cases()
        self.assertEqual(shared['Tax_Tables']['B219'], 0)
        for name in HISTORICAL_CASES:
            with self.subTest(case=name):
                fixture = cases[name]
                self.assertEqual(fixture['overrides']['Risk_Trials']['B3'], 1)
                self.assertEqual(fixture['overrides']['Risk_Settings']['B31'], 0)
                self.assertTrue(any(check.get('sheet', 'Projection') == 'Projection' and 'age' in check
                                    for check in fixture['expects']))

    def test_new_threshold_cases_isolate_credit_and_surcharge(self) -> None:
        """Lock the hand-calculated deltas on both sides of the modeled boundaries."""
        _, cases = plan_cases()
        for index, offset in enumerate(('below', 'at', 'above')):
            with self.subTest(offset=offset):
                credit = cases[f'california_published_credit_{offset}_threshold']
                surtax = cases[f'california_published_surtax_{offset}_threshold']
                credit_check = next(check for check in credit['expects'] if check.get('column') == 'CD')
                surtax_check = next(check for check in surtax['expects'] if check.get('column') == 'AT')
                self.assertEqual(credit_check['equals'], (474, 474, 473.82)[index])
                self.assertEqual(surtax_check['equals'], (0, 0, 1)[index])
                self.assertEqual(surtax['overrides']['Risk_Trials']['B3'], 2201 if offset == 'at' else 1)

    def test_requested_cases_fail_closed_on_unknown_or_duplicate_name(self) -> None:
        """Never silently run a different case than the one specified."""
        _, cases = plan_cases()
        case = HISTORICAL_CASES[0]
        self.assertEqual(select_cases(cases, [case]), [cases[case]])
        self.assertEqual(select_cases(cases, []), list(cases.values()))
        with self.assertRaisesRegex(ValueError, 'Unknown California parity case'):
            select_cases(cases, [case, 'missing'])
        with self.assertRaisesRegex(ValueError, 'repeated'):
            select_cases(cases, [case, case])

    def test_released_legacy_workbook_cannot_count_as_v2_parity(self) -> None:
        """Require explicit authoring before starting expensive workbook fixtures."""
        published = Path(__file__).resolve().parents[1] / 'workbook/retirement-model.xlsm'
        with self.assertRaisesRegex(ValueError, 'explicit California v2'):
            verify_candidate(published)
