"""Check the candidate workbook policy before building a release artifact."""

from collections import defaultdict
from types import SimpleNamespace
import unittest

from lxml import etree
from openpyxl import Workbook

import input_validation
import model_plan
import native_capture
import native_risk
from risk_engine.scenario import _workbook_state_policy, workbook_has_explicit_state_policy
from state_policy import CALIFORNIA_POLICY
import tax_policy
from workbook_protection import X, _entry_cells


class WorkbookPolicyAuthoringTests(unittest.TestCase):
    def setUp(self) -> None:
        """Collect edits from the workbook compiler in a small workbook."""
        self.model = SimpleNamespace(cells=defaultdict(dict), styles=defaultdict(dict), shared=defaultdict(dict))

        def put(sheet: str, address: str, value: str | float, style: str | None = None) -> None:
            self.model.cells[sheet][address] = value
            self.model.styles[sheet][address] = style

        self.model.put = put
        self.baseline = Workbook()
        self.baseline.active.title = 'Tax_Tables'

    def test_compiler_defaults_to_legacy(self) -> None:
        """Require a deliberate test invocation before adding explicit policy cells."""
        args = ['source.xlsx', 'plan.json']
        self.assertFalse(model_plan._arguments(args).test_california_policy)
        self.assertTrue(model_plan._arguments([*args, '--test-california-policy']).test_california_policy)

    def test_authoring_matches_the_reader_and_keeps_numeric_tax_validation(self) -> None:
        """Author each scenario's literal and require metadata and a working selector."""
        tax_policy.configure_workbook_state_policy(self.model.put)
        raw = Workbook()
        raw.active.title = 'Inputs'
        raw.create_sheet('Tax_Tables')
        for name, values in self.model.cells.items():
            for cell, value in values.items():
                raw[name][cell] = value
        cached = Workbook()
        cached.active.title = 'Inputs'
        cached['Inputs']['B136'] = 'CA'
        for selected in (1, 2, 3):
            with self.subTest(selected=selected):
                self.assertIs(_workbook_state_policy(raw, cached, selected=selected), CALIFORNIA_POLICY)

        check = self.model.cells['Inputs']['B129']
        self.assertIn('EXACT($B$136,"CA")', check)
        for column in 'DEF':
            self.assertIn(f'_xlfn.ISFORMULA(${column}$136)', check)
        self.assertIn('EXACT(Tax_Tables!$B$266,"workbook-state-policy-v1")', check)
        self.assertIn('EXACT(Tax_Tables!$B$267,"ca-individual-v1")', check)
        self.assertIn('COUNT(INDEX(D126:F128,0,B4))=3', check)
        self.assertIn('OR(B25>0,B51=0)', check)
        self.assertIn('Check state tax jurisdiction', check)

        # B129 by itself must still identify the new format after accidental stripping.
        for name, addresses in [('Inputs', ('A136', 'B136', 'D136', 'E136', 'F136')),
                                ('Tax_Tables', ('B266', 'B267'))]:
            for address in addresses:
                raw[name][address] = None
        self.assertTrue(workbook_has_explicit_state_policy(raw))

    def test_native_context_binds_state_and_trial_links(self) -> None:
        """Include only the active selection and its versions in each capture context."""
        native_capture._context(self.model, self.baseline)
        legacy = dict(self.model.cells['Market_Paths'])
        self.assertFalse(any('state tax policy' in value.lower() for value in legacy.values()))

        native_capture._context(self.model, self.baseline, with_state_policy=True)
        new = self.model.cells['Market_Paths']
        for label, reference in (
            ('State tax jurisdiction', 'Inputs!$B$136'),
            ('State tax workbook schema', 'Tax_Tables!$B$266'),
            ('State tax policy version', 'Tax_Tables!$B$267'),
        ):
            address = next(address for address, value in new.items() if value == label)
            self.assertEqual(new['BB' + address[2:]], f'=IF(ISBLANK({reference}),"",{reference})')
            self.assertNotIn(address, legacy)

        native_risk._clone_annual(self.model, self.baseline, with_state_policy=True)
        trial = self.model.cells['Trial_Inputs']
        self.assertEqual(trial['B136'], '=IF(ISBLANK(Inputs!B136),"",Inputs!B136)')
        self.assertIn('Inputs!B70>1', trial['B70'])
        self.assertIn('Inputs!B129', native_risk._validation(with_state_policy=True))
        self.assertNotIn('Inputs!B129', native_risk._validation())

    def test_candidate_validation_and_protection_are_scoped_to_present_input_row(self) -> None:
        """Keep the legacy layout and make only the three candidate literals editable."""
        root = etree.Element(X + 'worksheet')
        legacy = _entry_cells('Inputs', root)
        input_validation.apply('Inputs', root)
        self.assertFalse(root.xpath('//*[local-name()="dataValidation"][@sqref="D136:F136"]'))

        root = etree.Element(X + 'worksheet')
        body = etree.SubElement(root, X + 'sheetData')
        row = etree.SubElement(body, X + 'row', r='136')
        active = etree.SubElement(row, X + 'c', r='B136')
        etree.SubElement(active, X + 'f').text = 'INDEX($D$136:$F$136,1,$B$4)'
        self.assertEqual(_entry_cells('Inputs', root) - legacy, {'D136', 'E136', 'F136'})
        input_validation.apply('Inputs', root, with_state_policy=True)
        rule, = root.xpath('//*[local-name()="dataValidation"][@sqref="D136:F136"]')
        self.assertEqual(rule.get('allowBlank'), '0')
        self.assertEqual(rule.get('errorStyle'), 'stop')
        self.assertIn('EXACT(D136,"CA")', rule.findtext(X + 'formula1'))
