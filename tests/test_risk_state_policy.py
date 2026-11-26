"""Validate the supported state policy before a calculation uses it."""

from dataclasses import replace
import unittest

from openpyxl import Workbook

from risk_engine.scenario import ScenarioError, _workbook_state_policy, workbook_has_explicit_state_policy
from risk_engine.tax import TaxTables
from state_policy import CALIFORNIA_POLICY, resolve_state_policy


class StatePolicyTests(unittest.TestCase):
    def test_only_exact_california_policy_is_resolved(self) -> None:
        """Resolve a versioned jurisdiction without inferring missing input."""
        self.assertIs(resolve_state_policy('CA', 'ca-individual-v1'), CALIFORNIA_POLICY)
        for code in (None, '', ' CA', 'ca', 'California', 'FL', 'ZZ', 1, True):
            with self.subTest(code=code), self.assertRaisesRegex(
                ValueError, 'Unsupported state tax jurisdiction'
            ):
                resolve_state_policy(code, 'ca-individual-v1')
        for version in (None, '', 'CA-INDIVIDUAL-V1', 'ca-individual-v2', 1, True):
            with self.subTest(version=version), self.assertRaisesRegex(
                ValueError, 'Unsupported state tax policy version'
            ):
                resolve_state_policy('CA', version)

    def test_tax_tables_reject_changed_policy_metadata(self) -> None:
        """A version cannot be reused with different published data vintages."""
        changed = replace(CALIFORNIA_POLICY, published_income_year=2025)
        with self.assertRaisesRegex(ValueError, 'State tax policy metadata'):
            TaxTables.from_cells({}, state_policy=changed)
        self.assertIs(TaxTables.from_cells({}).state_policy, CALIFORNIA_POLICY)


class WorkbookStatePolicyTests(unittest.TestCase):
    def setUp(self) -> None:
        """Build raw and cached workbook views around the state-policy fields."""
        self.raw = Workbook()
        self.cached = Workbook()
        self.raw.active.title = 'Inputs'
        self.cached.active.title = 'Inputs'
        self.raw.create_sheet('Tax_Tables')
        self.raw.create_sheet('Projection')
        self.raw['Inputs']['A128'] = 'Opening California capital loss carryforward'
        self.raw['Projection']['BY4'] = 'California capital loss carryforward'

    def explicit(self) -> None:
        """Mark the explicit schema and select California in the first scenario."""
        inputs = self.raw['Inputs']
        inputs['A136'] = 'State tax jurisdiction'
        inputs['B136'] = '=INDEX(D136:F136,1,B4)'
        inputs['D136'] = 'CA'
        inputs['B129'] = '=IF(EXACT(B136,"CA"),"OK","Check state tax jurisdiction")'
        self.cached['Inputs']['B136'] = 'CA'
        self.raw['Tax_Tables']['B266'] = 'workbook-state-policy-v1'
        self.raw['Tax_Tables']['B267'] = 'ca-individual-v1'

    def test_legacy_requires_positive_california_identification(self) -> None:
        """Recognized legacy books have no explicit state identity in the digest."""
        self.assertIsNone(_workbook_state_policy(self.raw, self.cached, selected=1))
        self.raw['Projection']['BY4'] = None
        with self.assertRaisesRegex(ScenarioError, 'Projection!BY4'):
            _workbook_state_policy(self.raw, self.cached, selected=1)

    def test_explicit_policy_reads_only_the_active_scenario(self) -> None:
        """Editing an inactive input does not invalidate a valid selected scenario."""
        self.explicit()
        self.raw['Inputs']['E136'] = 'FL'
        self.assertIs(_workbook_state_policy(self.raw, self.cached, selected=1), CALIFORNIA_POLICY)
        with self.assertRaisesRegex(ScenarioError, 'Inputs!E136'):
            _workbook_state_policy(self.raw, self.cached, selected=2)

    def test_explicit_policy_does_not_fall_back_to_california(self) -> None:
        """Reject missing or unknown markers, invalid inputs and a stale active cache."""
        self.explicit()
        for book, sheet, cell, value in (
            (self.raw, 'Tax_Tables', 'B266', None),
            (self.raw, 'Tax_Tables', 'B266', 'workbook-state-policy-v2'),
            (self.raw, 'Tax_Tables', 'B267', None),
            (self.raw, 'Tax_Tables', 'B267', 'ca-individual-v2'),
            (self.raw, 'Inputs', 'A136', None),
            (self.raw, 'Inputs', 'D136', None),
            (self.raw, 'Inputs', 'D136', 'ca'),
            (self.raw, 'Inputs', 'D136', ' CA'),
            (self.raw, 'Inputs', 'D136', True),
            (self.raw, 'Inputs', 'D136', '="CA"'),
            (self.raw, 'Inputs', 'B136', 'CA'),
            (self.cached, 'Inputs', 'B136', None),
            (self.cached, 'Inputs', 'B136', 'FL'),
        ):
            previous = book[sheet][cell].value
            try:
                book[sheet][cell] = value
                with self.subTest(sheet=sheet, cell=cell, value=value), self.assertRaisesRegex(
                    ScenarioError, rf'{sheet}!{cell}'
                ):
                    _workbook_state_policy(self.raw, self.cached, selected=1)
            finally:
                book[sheet][cell] = previous

    def test_current_formula_prevents_fallback_when_new_cells_are_removed(self) -> None:
        """The state-aware tax gate also identifies a damaged explicit workbook."""
        self.raw['Inputs']['B129'] = '=IF(EXACT($B$136,"CA"),"OK","Check state tax jurisdiction")'
        with self.assertRaisesRegex(ScenarioError, 'Tax_Tables!B266'):
            _workbook_state_policy(self.raw, self.cached, selected=1)

    def test_partial_workbook_markers_survive_a_missing_sheet(self) -> None:
        """The snapshot guard still detects markers if one source sheet is missing."""
        self.explicit()
        del self.raw['Tax_Tables']
        self.assertTrue(workbook_has_explicit_state_policy(self.raw))
        self.raw.create_sheet('Tax_Tables')['B266'] = 'workbook-state-policy-v1'
        del self.raw['Inputs']
        self.assertTrue(workbook_has_explicit_state_policy(self.raw))
