"""Exercise the disposable VBA preflight without invoking desktop Excel."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

from desktop_policy_preflight import overlay, prepare
from native_bootstrap import _read_plan
from package_macro import SOURCE, package_macro, verify_project


LEGACY_SCHEMA = 'retirement-risk-native-v1'
POLICY_SCHEMA = 'retirement-risk-native-v2'
LEGACY_CONTEXT = 'native-risk-v1|n1|'
POLICY_CONTEXT = 'native-risk-v2|n1|tCA|tworkbook-state-policy-v1|tca-individual-v1|'
EXERCISE = '''
Public Function ExercisePolicy() As String
    Dim jurisdiction As String
    jurisdiction = ReadWorkbookStatePolicy(ThisWorkbook)
    RecalculateStateSelection ThisWorkbook, jurisdiction
    RequireStatePolicyContext ThisWorkbook, jurisdiction, AssumptionRecord(ThisWorkbook)
    ExercisePolicy = jurisdiction
End Function
'''


def fixture(*, explicit: bool) -> Workbook:
    """Build a small representative sheet set with one prior batch sentinel."""
    book = Workbook()
    inputs = book.active
    inputs.title = 'Inputs'
    inputs['B4'] = 1
    inputs['A128'] = 'Opening California capital loss carryforward'
    inputs['B129'] = '="OK"'
    book.create_sheet('Projection')['BY4'] = 'California capital loss carryforward'
    tables = book.create_sheet('Tax_Tables')
    risk = book.create_sheet('Risk_Data')
    risk['Z2'] = POLICY_SCHEMA if explicit else LEGACY_SCHEMA
    risk['Z4'], risk['Z6'], risk['Z7'], risk['Z8'] = 'Excel 1.0', 'complete', 46200, 'prior-record'
    book.create_sheet('Risk_Trials')['B3'] = 17
    book.create_sheet('Market_Paths')['BB3'] = f'="{POLICY_CONTEXT if explicit else LEGACY_CONTEXT}"'
    if explicit:
        inputs['A136'] = 'State tax jurisdiction'
        inputs['B136'] = '=INDEX($D$136:$F$136,1,$B$4)'
        inputs['B129'] = '=IF(EXACT($B$136,"CA"),"OK","Check state tax jurisdiction")'
        inputs['D136'] = inputs['E136'] = inputs['F136'] = 'CA'
        tables['A266'], tables['B266'] = 'State tax workbook schema', 'workbook-state-policy-v1'
        tables['A267'], tables['B267'] = 'California state tax policy version', 'ca-individual-v1'
    return book


class PolicyPreflightTests(unittest.TestCase):
    def test_default_runner_stays_v1_and_overlay_gates_running(self) -> None:
        """Raw validation and recalculated checks precede any running status."""
        source = SOURCE.read_text(encoding='ascii')
        code = overlay(source)
        run = code.split('Public Sub RunMarketRisk()', 1)[1].split('End Sub', 1)[0]
        markers = (
            'jurisdiction = ReadWorkbookStatePolicy(book)',
            'SaveExcelState settings',
            'RecalculateStateSelection book, jurisdiction',
            'selector.Value2 = 0',
            'Application.Calculate',
            'RequireStatePolicyContext book, jurisdiction, context',
            'runBegan = True',
            'results.Range("Z6").Value2 = "running"',
        )
        positions = [run.index(marker) for marker in markers]
        self.assertEqual(positions, sorted(positions))
        restored = run.index('    selectorRestored = True\n')
        rechecked = run.index('RequireText ReadWorkbookStatePolicy(book)', restored)
        self.assertLess(rechecked, run.index('    publishing = True\n'))
        self.assertIn('RequireText results.Range("Z2").Value2, EXPECTED_SCHEMA', source)
        self.assertEqual(run.split('RunFailed:', 1)[1], source.split('RunFailed:', 1)[1].split('End Sub', 1)[0])
        with self.assertRaisesRegex(ValueError, 'anchor'):
            overlay(source.replace('    runBegan = True\n', '', 1))

    def assert_fixture(self, explicit: bool, edits: dict[tuple[str, str], object], *, error: str = '') -> None:
        """Execute the reviewed fragment with its real literal and calculation helpers."""
        workbook = fixture(explicit=explicit)
        for (sheet, address), data in edits.items():
            workbook[sheet][address] = data
        with TemporaryDirectory() as folder:
            path = Path(folder) / 'policy.xlsx'
            workbook.save(path)
            app = ExcelApplication.open(path, with_vba=False)
            app.add_module(overlay(SOURCE.read_text(encoding='ascii')) + EXERCISE, name='MarketRiskRunner')
            if error:
                with self.assertRaisesRegex(VBARuntimeError, error):
                    app.run('ExercisePolicy')
            else:
                self.assertEqual(app.run('ExercisePolicy'), 'CA' if explicit else '')
            self.assertEqual(app.evaluate('Worksheets("Risk_Trials").Range("B3").Value2'), 17)
            self.assertEqual(app.evaluate('Worksheets("Risk_Data").Range("Z6").Value2'), 'complete')
            self.assertEqual(app.evaluate('Worksheets("Risk_Data").Range("Z7").Value2'), 46200)
            self.assertEqual(app.evaluate('Worksheets("Risk_Data").Range("Z8").Value2'), 'prior-record')

    def test_explicit_selection_and_recalculated_binding(self) -> None:
        """Accept one canonical scenario and reject forged values, shapes and context."""
        self.assert_fixture(True, {})
        self.assert_fixture(True, {('Inputs', 'B4'): 3, ('Inputs', 'D136'): 'FL'})
        for sheet, address, value, error in (
            ('Inputs', 'D136', None, 'state tax jurisdiction'),
            ('Inputs', 'D136', 'ca', 'state tax jurisdiction'),
            ('Inputs', 'D136', ' CA', 'state tax jurisdiction'),
            ('Inputs', 'D136', '="CA"', 'state tax jurisdiction'),
            ('Inputs', 'B4', '=1', 'scenario'),
            ('Inputs', 'B136', 'CA', 'state tax selector'),
            ('Inputs', 'B136', '="FL"', 'recalculated state tax jurisdiction'),
            ('Inputs', 'B129', '="OK"', 'state tax gate'),
            ('Tax_Tables', 'B266', None, 'state tax workbook schema'),
            ('Tax_Tables', 'B267', '="ca-individual-v1"', 'California state tax policy version'),
            ('Tax_Tables', 'B267', 'CA-INDIVIDUAL-V1', 'California state tax policy version'),
            ('Risk_Data', 'Z2', LEGACY_SCHEMA, 'explicit state policy'),
            ('Risk_Data', 'Z2', f'="{POLICY_SCHEMA}"', 'native risk schema'),
            ('Market_Paths', 'BB3', f'="{LEGACY_CONTEXT}"', 'assumption record version'),
            ('Market_Paths', 'BB3', f'="{POLICY_CONTEXT.replace("tCA|", "tFL|")}"', 'assumption record policy'),
            ('Market_Paths', 'BB3', POLICY_CONTEXT, 'policy assumption record formula'),
        ):
            with self.subTest(sheet=sheet, cell=address, value=value):
                self.assert_fixture(True, {(sheet, address): value}, error=error)

    def test_only_recognized_legacy_layouts_can_run_without_policy_markers(self) -> None:
        """Reject partial explicit inputs and formats whose legacy identity is missing."""
        self.assert_fixture(False, {('Inputs', 'B129'): '=IF(B1360>0,"OK","OK")'})
        self.assert_fixture(False, {('Tax_Tables', 'A266'): 'Unrelated caption'})
        for sheet, address, value, error in (
            ('Projection', 'BY4', None, 'legacy California tax layout'),
            ('Inputs', 'A128', '="Opening California capital loss carryforward"', 'legacy California tax layout'),
            ('Inputs', 'B129', '=IF(EXACT($B$136,"CA"),"OK","Check state tax jurisdiction")', 'explicit state policy'),
            ('Inputs', 'F136', '=""', 'explicit state policy'),
            ('Inputs', 'G136', 'Use CA for a full-year California household.', 'explicit state policy'),
            ('Tax_Tables', 'A266', 'State tax workbook schema', 'explicit state policy'),
            ('Tax_Tables', 'A267', 'California state tax policy version', 'explicit state policy'),
            ('Tax_Tables', 'B267', 'ca-individual-v1', 'explicit state policy'),
            ('Risk_Data', 'Z2', POLICY_SCHEMA, 'state tax workbook schema'),
        ):
            with self.subTest(sheet=sheet, cell=address, value=value):
                self.assert_fixture(False, {(sheet, address): value}, error=error)

    def test_full_runner_rejects_raw_values_without_touching_prior_results(self) -> None:
        """Execute the VBA startup error and inspect the retained batch metadata."""
        with TemporaryDirectory() as folder:
            path = Path(folder) / 'policy.xlsx'
            book = fixture(explicit=True)
            book['Inputs']['D136'] = '="CA"'
            book['Risk_Trials']['A8'], book['Risk_Trials']['B8'] = 1, 'prior-trial'
            book.save(path)
            app = ExcelApplication.open(path, with_vba=False)
            app.add_module(overlay(SOURCE.read_text(encoding='ascii')), name='MarketRiskRunner')
            app.run('RunMarketRisk')
            self.assertIn('Enter the literal state tax jurisdiction', app.interpreter.dialogs[-1][1])
            self.assertEqual(app.evaluate('Worksheets("Risk_Data").Range("Z6").Value2'), 'complete')
            self.assertEqual(app.evaluate('Worksheets("Risk_Data").Range("Z7").Value2'), 46200)
            self.assertEqual(app.evaluate('Worksheets("Risk_Data").Range("Z8").Value2'), 'prior-record')
            self.assertEqual(app.evaluate('Worksheets("Risk_Trials").Range("B3").Value2'), 17)
            self.assertEqual(app.evaluate('Worksheets("Risk_Trials").Range("B8").Value2'), 'prior-trial')

    def test_bootstrap_requires_the_workbook_and_record_versions_to_pair(self) -> None:
        """Require the matching policy suffix before Python creates a v2 saved batch."""
        with TemporaryDirectory() as folder:
            path = Path(folder) / 'calculated.xlsx'
            for explicit, context, error in (
                (False, LEGACY_CONTEXT, ''),
                (True, POLICY_CONTEXT, ''),
                (True, LEGACY_CONTEXT, 'market assumption record is missing'),
                (True, POLICY_CONTEXT.replace('tCA|', 'tFL|'), 'does not contain the California policy'),
                (False, POLICY_CONTEXT, 'market assumption record is missing'),
            ):
                with self.subTest(explicit=explicit, context=context):
                    book = fixture(explicit=explicit)
                    book['Risk_Data']['Z1'] = 'missing'
                    book['Market_Paths']['AL3'], book['Market_Paths']['BB3'] = 'OK', context
                    book['Inputs']['B53'] = 95
                    settings = book.create_sheet('Risk_Settings')
                    for row, value in enumerate((1, 2026, 1, 0, .015, .5, -.02, .10), 29):
                        settings[f'B{row}'] = value
                    book.save(path)
                    if error:
                        with self.assertRaisesRegex(ValueError, error):
                            _read_plan(path)
                    else:
                        self.assertEqual(_read_plan(path)[1], context)

    def test_package_is_confirmed_as_scratch_and_refuses_overwrite(self) -> None:
        """Use the normal package verifier with an explicitly generated overlay source."""
        with TemporaryDirectory() as folder:
            draft, candidate = (Path(folder) / name for name in ('draft.xlsx', 'preflight.xlsm'))
            fixture(explicit=True).save(draft)
            with self.assertRaisesRegex(ValueError, 'test overlay'):
                package_macro(draft, candidate)
            self.assertFalse(candidate.exists())
            self.assertEqual(prepare(draft, candidate), candidate)
            with self.assertRaisesRegex(ValueError, 'published source'):
                verify_project(candidate)
            with self.assertRaises(FileExistsError):
                prepare(draft, candidate)

    def test_package_refuses_mixed_formats_and_recognizes_legacy(self) -> None:
        """Reject incomplete or incompatible drafts before any package is written."""
        with TemporaryDirectory() as folder:
            draft, candidate = (Path(folder) / name for name in ('draft.xlsx', 'preflight.xlsm'))
            for explicit, edit, error in (
                (True, ('Risk_Data', 'Z2', LEGACY_SCHEMA), 'supported native schema'),
                (True, ('Tax_Tables', 'B267', None), 'complete state tax schema'),
                (True, ('Market_Paths', 'BB3', f'="{LEGACY_CONTEXT}"'), 'does not match the native risk schema'),
                (False, ('Projection', 'BY4', None), 'legacy California tax layout'),
            ):
                with self.subTest(edit=edit):
                    book = fixture(explicit=explicit)
                    sheet, address, value = edit
                    book[sheet][address] = value
                    book.save(draft)
                    with self.assertRaisesRegex(ValueError, error):
                        prepare(draft, candidate)
                    self.assertFalse(candidate.exists())
                    if explicit:
                        with self.assertRaisesRegex(ValueError, 'requires the test overlay'):
                            package_macro(draft, candidate)
                        self.assertFalse(candidate.exists())
            fixture(explicit=False).save(draft)
            self.assertEqual(prepare(draft, candidate), candidate)
