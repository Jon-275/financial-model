"""Exercise workbook provenance and fail-closed cache checks without editing the source."""

from __future__ import annotations

import hashlib
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree
from zipfile import ZipFile

from lxml import etree

from risk_engine.scenario import ScenarioError, load_scenario
from state_policy import CALIFORNIA_POLICY
from workbook_recalc import RecalculationError, clear_formula_results
from xlsx_read import q, sheet_parts
from xlsx_values import X, update_sheet, write_values


_WORKBOOK = Path(os.environ.get('RISK_TEST_WORKBOOK',
                                Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsm'))
_LEGACY_RELEASES = {
    '37e2c8d564b5512a1bad26e402c2cad18393fe3a439652ba1dd1c8bc79940ef3',  # v0.1.0
    'f5faea8318773c911800462d43488c5063dc20b9363f1121151454d5e8007af4',  # v1.0.0
}


def _replace_saved_number(source: Path, destination: Path, sheet: str, coordinate: str, value: float) -> None:
    with ZipFile(source) as original, ZipFile(destination, 'x') as modified:
        part = sheet_parts(original)[sheet]
        for entry in original.infolist():
            payload = original.read(entry)
            if entry.filename == part:
                tree = ElementTree.fromstring(payload)
                cell = next(cell for cell in tree.iter(q('c')) if cell.get('r') == coordinate)
                cell.find(q('v')).text = str(value)
                payload = ElementTree.tostring(tree, encoding='UTF-8', xml_declaration=True)
            modified.writestr(entry, payload)


def _write_explicit_policy_fixture(source: Path, destination: Path) -> None:
    """Write the future schema and active state formula cache into a fictional copy."""
    with ZipFile(source) as original, ZipFile(destination, 'x') as modified:
        parts = sheet_parts(original)
        for entry in original.infolist():
            payload = original.read(entry)
            if entry.filename == parts['Inputs']:
                sheet = etree.fromstring(payload)
                selected = sheet.find(f'.//{X}c[@r="B4"]/{X}v').text
                scenario_cell = f'{"DEF"[int(selected) - 1]}136'
                update_sheet(sheet, {'A136': 'State tax jurisdiction', 'B136': 'CA', scenario_cell: 'CA'})
                active = sheet.find(f'.//{X}c[@r="B136"]')
                active.remove(active.find(X + 'is'))
                active.set('t', 'str')
                etree.SubElement(active, X + 'f').text = 'INDEX(D136:F136,1,B4)'
                etree.SubElement(active, X + 'v').text = 'CA'
                payload = etree.tostring(sheet, encoding='UTF-8', xml_declaration=True)
            elif entry.filename == parts['Tax_Tables']:
                sheet = etree.fromstring(payload)
                update_sheet(sheet, {'B266': 'workbook-state-policy-v1', 'B267': 'ca-individual-v1'})
                payload = etree.tostring(sheet, encoding='UTF-8', xml_declaration=True)
            modified.writestr(entry, payload)


class ScenarioWorkbookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.scenario = load_scenario(_WORKBOOK)

    def test_opening_and_projection_horizon(self):
        s = self.scenario
        self.assertEqual(s.opening_projection['C'], s.inputs[8])
        self.assertEqual(len(s.years), s.inputs[53] - s.inputs[8])
        self.assertEqual(s.years[-1].projection['C'], s.inputs[53])
        self.assertEqual(s.workbook_sha256, hashlib.sha256(_WORKBOOK.read_bytes()).hexdigest())
        self.assertAlmostEqual(sum(s.allocation.working[0]), 1.0)

    def test_legacy_california_policy_preserves_the_release_digest(self) -> None:
        """Recognize published legacy California workbooks without changing their identity."""
        self.assertIs(self.scenario.state_policy, CALIFORNIA_POLICY)
        if self.scenario.workbook_sha256 in _LEGACY_RELEASES:
            self.assertEqual(
                self.scenario.digest, '5a5388fd511c3da3fa227945e9cfa531ad04e0a310d334c1cc35b13a0d9cee5a'
            )

    def test_workbook_without_california_layout_fails_closed(self) -> None:
        """Missing legacy policy evidence cannot silently select California."""
        with TemporaryDirectory() as folder:
            changed = Path(folder) / 'unidentified-tax-layout.xlsx'
            write_values(_WORKBOOK, changed, {'Projection': {'BY4': None}})
            with self.assertRaisesRegex(
                ScenarioError, 'Projection!BY4: cannot identify the legacy California tax layout'
            ):
                load_scenario(changed)

    def test_explicit_california_policy_versions_the_digest(self) -> None:
        """A workbook carrying the new state policy schema gets a fresh scenario identity."""
        with TemporaryDirectory() as folder:
            explicit = Path(folder) / 'explicit-california.xlsx'
            _write_explicit_policy_fixture(_WORKBOOK, explicit)
            candidate = load_scenario(explicit)
        self.assertIs(candidate.state_policy, CALIFORNIA_POLICY)
        self.assertNotEqual(candidate.digest, self.scenario.digest)
        self.assertEqual(candidate.inputs, self.scenario.inputs)
        self.assertEqual(candidate.years, self.scenario.years)

    def test_changed_active_scenario_literal_invalidates_its_cache(self):
        s = self.scenario
        active = 'DEF'[s.selected - 1]
        with TemporaryDirectory() as folder:
            changed = Path(folder) / 'active-input.xlsx'
            _replace_saved_number(_WORKBOOK, changed, 'Inputs', f'{active}12', float(s.inputs[12]) + 1)
            with self.assertRaisesRegex(ScenarioError, 'Inputs!B12: saved calculation differs'):
                load_scenario(changed)

    def test_changed_unselected_literal_does_not_change_scenario_digest(self):
        s = self.scenario
        inactive = 'DEF'[s.selected % 3]
        with TemporaryDirectory() as folder:
            changed = Path(folder) / 'inactive-input.xlsx'
            _replace_saved_number(_WORKBOOK, changed, 'Inputs', f'{inactive}12', float(s.inputs[12]) + 1)
            candidate = load_scenario(changed)
        self.assertEqual(candidate.digest, s.digest)
        self.assertNotEqual(candidate.workbook_sha256, s.workbook_sha256)

    def test_changed_projection_cache_is_rejected_at_source(self):
        s = self.scenario
        with TemporaryDirectory() as folder:
            changed = Path(folder) / 'projection-cache.xlsx'
            _replace_saved_number(_WORKBOOK, changed, 'Projection', 'I6', float(s.years[0].projection['I']) + 1)
            with self.assertRaisesRegex(ScenarioError, 'Projection!I6: saved calculation differs'):
                load_scenario(changed)

    def test_editing_return_basis_updates_report_inputs_without_a_recalculation(self):
        s = self.scenario
        with ZipFile(_WORKBOOK) as original:
            modern = 'Risk_Settings' in sheet_parts(original)
        sheet, address = ('Risk_Settings', 'B9') if modern else ('Monte_Carlo', 'B8')
        with TemporaryDirectory() as folder:
            changed = Path(folder) / 'report-settings.xlsx'
            _replace_saved_number(_WORKBOOK, changed, sheet, address, 3 - s.return_basis)
            candidate = load_scenario(changed)
            self.assertEqual(candidate.return_basis, 3 - s.return_basis)
            self.assertEqual(candidate.years, s.years)
            self.assertNotEqual(candidate.digest, s.digest)

    def test_disposable_copy_clears_dependent_formula_caches(self):
        with TemporaryDirectory() as folder:
            clean = Path(folder) / 'no-caches.xlsx'
            clear_formula_results(_WORKBOOK, clean)
            with ZipFile(_WORKBOOK) as source, ZipFile(clean) as prepared:
                part = sheet_parts(source)['Projection']
                for archive, expect_saved in ((source, True), (prepared, False)):
                    projection = ElementTree.fromstring(archive.read(part))
                    headroom = next(c for c in projection.iter(q('c')) if c.get('r') == 'U6')
                    self.assertIsNotNone(headroom.find(q('f')))
                    self.assertEqual(headroom.find(q('v')) is not None, expect_saved)
                self.assertEqual(set(source.namelist()), set(prepared.namelist()))
                self.assertEqual(source.read('xl/drawings/drawing1.xml'), prepared.read('xl/drawings/drawing1.xml'))
            self.assertEqual(hashlib.sha256(_WORKBOOK.read_bytes()).hexdigest(), self.scenario.workbook_sha256)
            with self.assertRaisesRegex(RecalculationError, 'Use a new destination'):
                clear_formula_results(_WORKBOOK, clean)


if __name__ == '__main__':
    unittest.main()
