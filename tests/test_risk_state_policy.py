"""Validate the supported state policy before a calculation uses it."""

from dataclasses import replace
import unittest

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
