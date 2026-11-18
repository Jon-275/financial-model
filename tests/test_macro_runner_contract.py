"""Check the VBA/worksheet contract and its permitted automation surface.

These source checks complement native Excel compilation and run tests. They do
not establish that Excel executed the macro or that cancellation worked there.
"""
from collections import defaultdict
import os
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.utils.cell import range_boundaries
from pyopenvba.cfb import CFB
from pyopenvba.vba import compress, decompress

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))

import native_balance_checks
import native_capture
import native_funding_paths
import native_results
from package_macro import audit_project_metadata, package_macro
from vba_declarations import audit_vba_declarations

SOURCE = (ROOT / 'tools/vba/RunMarketRisk.bas').read_text()
CODE = '\n'.join(line for line in SOURCE.splitlines() if not line.lstrip().startswith("'"))


def constant(name):
    match = re.search(rf'^Private Const {name} As \w+ = (.+)$', CODE, re.M)
    if match is None:
        raise AssertionError(f'Missing runner constant: {name}')
    return match[1].strip('"')


def procedure(name):
    match = re.search(rf'^(?:Private|Public) (Sub|Function) {name}\b.*?^End \1$', CODE, re.M | re.S)
    if match is None:
        raise AssertionError(f'Missing runner procedure: {name}')
    return match[0]


class _TaxSheet:
    def __getitem__(self, _):
        return SimpleNamespace(value=None)


class MacroPublicationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        default = ROOT / 'workbook' / 'retirement-model.xlsm'
        with ZipFile(os.environ.get('RISK_TEST_WORKBOOK', default)) as archive:
            cls.payload = archive.read('xl/vbaProject.bin')
            book = ElementTree.fromstring(archive.read('xl/workbook.xml'))
        sheet_count = len(book.find('{*}sheets'))
        cls.module_names = {'ThisWorkbook', 'MarketRiskRunner',
                            *(f'Sheet{i}' for i in range(1, sheet_count + 1))}

    def test_rejects_extra_binary_records_at_each_level(self):
        for path, is_storage in (((), False), ((), True), (('VBA',), False), (('VBA',), True)):
            with self.subTest(path=path, is_storage=is_storage):
                project = CFB.from_bytes(self.payload)
                if is_storage:
                    project.add_substorage_at(path, 'Unexpected')
                else:
                    project.add_stream_at(path, 'Unexpected', b'')
                with self.assertRaisesRegex(ValueError, 'unexpected records'):
                    audit_project_metadata(project.to_bytes(), self.module_names)

    def test_rejects_profile_text_in_a_permitted_binary_record(self):
        for sample in ('/HOME/Example', r'C:\Users\Example', 'Macintosh HD:Users:Example'):
            for encoding in ('ascii', 'utf-16le', 'utf-16be'):
                with self.subTest(sample=sample, encoding=encoding):
                    project = CFB.from_bytes(self.payload)
                    original = project.get_stream_at((), 'PROJECTwm')
                    project.write_stream_at((), 'PROJECTwm', original + sample.encode(encoding))
                    with self.assertRaisesRegex(ValueError, 'local profile path'):
                        audit_project_metadata(project.to_bytes(), self.module_names)

    def test_rejects_profile_text_visible_only_after_directory_decompression(self):
        for encoding in ('ascii', 'utf-16le', 'utf-16be'):
            with self.subTest(encoding=encoding):
                project = CFB.from_bytes(self.payload)
                directory = decompress(project.get_stream_at(('VBA',), 'dir'))
                # Repeated pieces cause compression to split the complete path marker.
                sample = 'zzzz/Us**ers/*zz/Users/Example'.encode(encoding)
                project.write_stream_at(('VBA',), 'dir', compress(directory + sample))
                payload = project.to_bytes()
                self.assertNotIn('/users/'.encode(encoding), payload.lower())
                with self.assertRaisesRegex(ValueError, 'local profile path'):
                    audit_project_metadata(payload, self.module_names)


class MacroDeclarationTests(unittest.TestCase):
    def test_packaging_refuses_the_two_compile_defects_before_publishing(self):
        samples = (
            ('Private Function AuditBlank() As Boolean\nEnd Function\n'
             'Public Sub auditblank()\nEnd Sub', 'duplicate VBA procedure'),
            ('Private Sub LookupNative()\nDim dims(1 To 2, 1 To 3), _\n'
             '    tAb As Variant\nEnd Sub', 'reserved VBA variable name'),
        )
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            book = Workbook()
            book.save(root / 'example.xlsx')
            book.close()
            module = root / 'reviewed.bas'
            for source, error in samples:
                with self.subTest(error=error), patch('package_macro.SOURCE', module):
                    module.write_text('Attribute VB_Name = "MarketRiskRunner"\n' + source + '\n', encoding='ascii')
                    with self.assertRaisesRegex(ValueError, error):
                        package_macro(root / 'example.xlsx', root / 'example.xlsm')
                    self.assertFalse((root / 'example.xlsm').exists())

    def test_conditional_declarations_are_compared_on_each_target(self):
        source = """#If Mac Then
Private Sub DetectHost()
End Sub
#ElseIf Win64 Then
Private Sub DetectHost()
End Sub
#Else
Private Sub DetectHost()
End Sub
#End If
"""
        audit_vba_declarations(source)
        for inactive_marker in ('#If Mac Then', '#ElseIf Win64 Then', '#Else\n'):
            with self.subTest(branch=inactive_marker):
                duplicate = source.replace(inactive_marker, inactive_marker + '\nPrivate Function detectHost()\nEnd Function\n')
                with self.assertRaisesRegex(ValueError, 'duplicate VBA procedure'):
                    audit_vba_declarations(duplicate)

    def test_valid_names_attributes_and_distinct_property_accessors(self):
        source = """Attribute VB_Name = "Settings"
Private Type Holder
    Tab As Variant
End Type
Public Property Get Item() As Variant
Attribute Item.VB_Description = "Dim Tab: Public Sub Item()"
End Property
Public Property Let Item(ByVal value As Variant)
End Property
Public Property Set Item(ByVal value As Object)
End Property
Private Sub ReadTable()
    Dim mode As Long, year As Integer
    Dim [Tab] As Variant
    Dim tableRange(1 To 3, 1 To 2), _
        threshold As Double
    Rem Dim Tab: Public Sub Item()
    Print #1, Tab(2)
End Sub
"""
        audit_vba_declarations(source)
        with self.assertRaisesRegex(ValueError, 'duplicate VBA procedure'):
            audit_vba_declarations(source + 'Private Property Get ITEM() As Variant\nEnd Property')
        with self.assertRaisesRegex(ValueError, 'duplicate VBA procedure'):
            audit_vba_declarations(source + 'Private Function ITEM() As Variant\nEnd Function')

    def test_custom_directives_leave_unconditional_errors_visible(self):
        source = """#Const Custom = 1 + 2
#If Custom Then
Private Sub Alternate()
End Sub
#Else
Private Sub Alternate()
End Sub
#End If
"""
        with self.assertWarnsRegex(UserWarning, 'Conditional VBA declarations require Excel'):
            audit_vba_declarations(source)
        with self.assertWarns(UserWarning), self.assertRaisesRegex(ValueError, 'reserved VBA variable name'):
            audit_vba_declarations(source + 'Private Sub Work()\nDim Spc As Double\nEnd Sub')
        with self.assertRaisesRegex(ValueError, 'Unclosed #If'):
            audit_vba_declarations(source.removesuffix('#End If\n'))


class MacroRunnerContractTests(unittest.TestCase):
    def test_empty_and_text_inputs_cannot_be_coerced_to_numeric_seeds_or_counts(self):
        allowed = procedure('IsNumericInput')
        self.assertIn('Select Case VarType(value)', allowed)
        cases = re.search(r'^\s*Case (.+)$', allowed, re.M).group(1).split(', ')
        self.assertEqual(set(cases), {'vbByte','vbInteger','vbLong','vbSingle','vbDouble','vbCurrency','vbDecimal'})
        for function in ('SeedNumber','WholeNumber'):
            body = procedure(function)
            self.assertLess(body.index('If Not IsNumericInput(value) Then RaiseRunError'), body.index('CDbl(value)'))

    def test_compiler_exports_every_named_capture_area_used_by_runner(self):
        model = SimpleNamespace(cells=defaultdict(dict), styles=defaultdict(dict), MODULES=[])
        def put(sheet, address, value, style=None):
            model.cells[sheet][address] = value
        model.put = put
        native_capture.configure(model, {'Tax_Tables': _TaxSheet()})
        native_funding_paths.configure(model)
        used = set(re.findall(r'LocalRange\(trials, "([^"]+)"\)', procedure('RunMarketRisk')))
        exported = model.native_names['Risk_Trials']
        self.assertEqual(used, set(exported))
        primary = {name: range_boundaries(exported[name].replace('$', ''))
                   for name in ('RiskTrialCurrent', 'RiskTrialResults', 'RiskTrialIDs')}
        self.assertEqual(primary['RiskTrialCurrent'], (2, 7, native_capture.SIGNATURE_COLUMN, 7))
        self.assertEqual(primary['RiskTrialResults'], (2, 8, native_capture.SIGNATURE_COLUMN, native_capture.LAST))
        self.assertEqual(primary['RiskTrialIDs'], (1, 8, 1, native_capture.LAST))
        secondary = {}
        for name in ('RiskPathCurrent', 'RiskPathResults', 'RiskPathTrials'):
            sheet, area = exported[name].split('!')
            self.assertEqual(sheet.strip("'"), native_funding_paths.DATA_SHEET)
            secondary[name] = range_boundaries(area.replace('$', ''))
        self.assertEqual(secondary['RiskPathResults'][3] - secondary['RiskPathResults'][1] + 1, 3)
        self.assertEqual(secondary['RiskPathCurrent'][2] - secondary['RiskPathCurrent'][0] + 1,
                         len(native_funding_paths.FIELDS) * native_capture.ANNUAL_CAPACITY + 3)
        self.assertEqual(secondary['RiskPathTrials'], (1, native_funding_paths.FIRST_ROW, 1, native_funding_paths.LAST_ROW))
        checks = procedure('ValidateRanges')
        years = int(constant('ANNUAL_CAPACITY'))
        self.assertIn('current.Columns.Count <> FIRST_YEAR_FIELD + 2 * ANNUAL_CAPACITY', checks)
        self.assertEqual(int(constant('FIRST_YEAR_FIELD')) + 2 * years, native_capture.SIGNATURE_COLUMN - 1)
        self.assertIn('pathCurrent.Columns.Count <> 3 + ANNUAL_CAPACITY * FUNDING_YEAR_WIDTH', checks)
        self.assertEqual(3 + years * int(constant('FUNDING_YEAR_WIDTH')), native_funding_paths.SIGNATURE_COLUMN - 1)

    def test_ranking_fields_and_release_version_match_the_worksheet_compiler(self):
        expected = {
            'VALID_FIELD': 'Calculation valid',
            'FAILURE_AGE_FIELD': 'First shortfall age',
            'MAX_UNPAID_FIELD': 'Largest unpaid balance, today’s dollars',
            'ENDING_REAL_FIELD': 'Ending portfolio, today’s dollars',
        }
        for name, label in expected.items():
            self.assertEqual(int(constant(name)), native_capture.FIELDS.index(label) + 1)
        for name in ('ComparePaths','RepeatsPriorEnding','ValidateCaptureRecord','FailureAge'):
            self.assertIn('IsNumericInput(', procedure(name))
            self.assertNotIn('IsNumeric(', procedure(name))
        self.assertEqual(int(constant('MAX_PAIRED_PATHS')), native_capture.CAPACITY)
        self.assertEqual(constant('EXPECTED_SCHEMA'), native_results.SCHEMA)
        self.assertEqual(constant('EXPECTED_ENGINE'), native_results.ENGINE)

    def test_run_has_no_automatic_external_or_platform_specific_actions(self):
        public = re.findall(r'^Public Sub (\w+)\(', CODE, re.M)
        self.assertEqual(set(public), {'RunMarketRisk', 'CancelMarketRisk'})
        forbidden = r'\b(?:Auto_Open|Workbook_Open|OnTime|Shell|CreateObject|GetObject|VBProject|SaveAs|Kill|ChDir|Declare|MacScript|AppleScriptTask)\b'
        self.assertIsNone(re.search(forbidden, CODE, re.I))
        self.assertNotRegex(CODE, r'\bActive(?:Workbook|Sheet)\b|\bSelection\b')
        self.assertIn('Set book = ThisWorkbook', CODE)
        self.assertIn('Set owner = area.Parent.Parent', CODE)
        self.assertIn('If Not owner Is ThisWorkbook Then', CODE)
        self.assertNotRegex(CODE, r'\.Formula(?:2|R1C1)?\s*=')

    def test_primary_and_full_path_batches_finish_before_publication(self):
        run = procedure('RunMarketRisk')
        publish = run.index('publishing = True')
        self.assertLess(run.index('For index = 1 To total'), publish)
        self.assertLess(run.index('For caseIndex = 1 To 3'), publish)
        self.assertLess(run.index('RequireText AssumptionRecord(book), context'), publish)
        self.assertLess(run.index('oldBatch = trialResults.Value2'), publish)
        self.assertLess(run.index('oldPaths = pathResults.Value2'), publish)
        commit = run[publish:run.index('completed = True')]
        self.assertNotIn('DoEvents', commit)
        self.assertIn('trialResults.Value2 = batch', commit)
        self.assertIn('pathResults.Value2 = paths', commit)
        self.assertIn('RequireText results.Range("Z1").Value2, "current"', commit)
        self.assertIn('If publishing Then', run)
        restore = procedure('RestoreResults')
        for target in ('results.Value2 = oldResults', 'ids.Value2 = oldIDs',
                       'paths.Value2 = oldPaths', 'pathIDs.Value2 = oldPathIDs'):
            self.assertIn(target, restore)

    def test_reuse_requires_a_verified_current_fixed_partner(self):
        run = procedure('RunMarketRisk')
        proof = procedure('FixedPathAllowsReuse')
        partner = procedure('CanReuseFixedTrial')
        self.assertIn('cutProof.HasFormula', run)
        self.assertIn('lastCalculatedRow = index', run)
        self.assertIn('context, True, batch, pairedRow, lastCalculatedRow)', run)
        self.assertLess(run.index('captured = CaptureTrial('), run.index('fixedUnchanged(index) ='))
        self.assertIn('For period = 1 To CLng(years) - 1', proof)
        self.assertIn('FIRST_YEAR_FIELD + ANNUAL_CAPACITY + period - 1', proof)
        self.assertIn('proof(period, 1) <> 0#', proof)
        self.assertIn('opening < proof(period, 2)', proof)
        self.assertIn('expected = code - 1000', partner)
        self.assertIn('expected = code - 100', partner)
        for field in ('ids(pairedRow, 1), expected', 'batch(pairedRow, VALID_FIELD), 1#',
                      'batch(pairedRow, CUT_YEARS_FIELD), 0#', 'batch(pairedRow, 1), context',
                      'batch(pairedRow, width), context', 'batch(pairedRow, REFRESH_DATE_FIELD), currentDay'):
            self.assertIn('SameCaptureValue(' + field + ')', partner)
        self.assertIn('If book.Date1904 Then', partner)
        self.assertIn('RequireText AssumptionRecord(book), context', partner)

    def test_cancellation_cleanup_and_early_failure_keep_explicit_guards(self):
        run = procedure('RunMarketRisk')
        self.assertIn('If mRunning Then', run)
        self.assertIn('Application.EnableCancelKey = xlErrorHandler', run)
        self.assertIn('If failureNumber = 18 Then', run)
        self.assertIn('Resume CleanExit', run)
        cleanup = run.split('CleanExit:', 1)[1]
        self.assertRegex(cleanup, r'If savedSettings Then\s+Application.EnableCancelKey = xlDisabled')
        self.assertIn('If savedSelector And selectorTouched And Not selectorRestored Then', cleanup)
        self.assertIn('mLastFundingYears = ANNUAL_CAPACITY', run[:run.index('For index = 1 To total')])
        self.assertIn('mLastFundingYears = ANNUAL_CAPACITY\n        RestoreInspection book, selector, previousSelector', cleanup)
        self.assertIn('RestoreInspection book, selector, previousSelector\n    selectorRestored = True', run)
        self.assertIn('UpdateInactiveFunding book, years', procedure('CaptureBalancesAgree'))
        self.assertIn('UpdateInactiveFunding book, years', procedure('RestoreInspection'))
        inspection = procedure('RestoreInspection')
        self.assertLess(inspection.index('If years = 0 Then'), inspection.index('RecalculateTrialYears book, True'))
        self.assertLess(inspection.index('RecalculateTrialYears book, True'), inspection.index('values = current.Value2'))
        self.assertIn('CaptureBalancesAgree(book, values, True, noPartner, 0, detail)', inspection)
        self.assertIn('ValidateCaptureRecord values, code, AssumptionRecord(book)', inspection)
        self.assertIn('ValidateCaptureRecord values, code, context', procedure('CaptureTrial'))
        self.assertIn('If yearRow Mod 8 = 0 And Not restoringInspection Then', procedure('RecalculateTrialYears'))
        self.assertIn('If runBegan And Not results Is Nothing Then', cleanup)
        self.assertIn('mRunning = False', cleanup)
        self.assertIn('DoEvents', procedure('ShowProgress'))
        self.assertIn('If mCancelRequested Then Err.Raise 18', procedure('CheckCancelled'))
        saved = procedure('SaveExcelState')
        restored = procedure('RestoreExcelState')
        for prop in ('Calculation', 'EnableEvents', 'ScreenUpdating', 'StatusBar', 'DisplayStatusBar', 'Cursor'):
            self.assertIn(f'Application.{prop}', saved)
            self.assertIn(f'Application.{prop} =', restored)
        self.assertIn('Application.EnableCancelKey = settings.CancelKey', cleanup)

    def test_every_failed_or_suspect_full_calculation_gets_one_ordered_retry(self):
        capture = procedure('CaptureTrial')
        expected = (
            'For attempt = 1 To 2',
            'selector.Value2 = code',
            'If attempt = 1 Then',
            'FinishTrialCalculation',
            'Else',
            'CheckCancelled',
            'RecalculateTrialYears book',
            'End If',
            'source.Calculate',
        )
        primary = capture[capture.index(expected[0]):]
        for operation in expected:
            start = primary.index(operation)
            primary = primary[start + len(operation):]
        self.assertEqual(capture.count('RecalculateTrialYears book'), 1)
        self.assertNotIn('selector.Value2 = 0', capture)
        self.assertIn('If attempt > 2 Then RaiseRunError', capture)
        self.assertIn('Application.CalculateFull', procedure('FinishTrialCalculation'))
        self.assertNotIn('Application.CalculateFullRebuild', CODE)
        ordered = procedure('RecalculateTrialYears')
        self.assertIn('For yearRow = 5 To lastRow', ordered)
        self.assertIn('CheckCancelled', ordered)
        self.assertIn('LocalRange(book.Worksheets("Risk_Trials"), "RiskTrialCurrent").Calculate', ordered)
        self.assertNotIn('Application.Calculate', ordered)

    def test_capture_waits_for_dependencies_and_reconciles_before_storage(self):
        capture = procedure('CaptureTrial')
        run = procedure('RunMarketRisk')
        self.assertIn('trialCurrent, code, context, True, batch', run)
        self.assertIn('pathCurrent, code, context, False, batch', run)
        self.assertIn('ByVal primary As Boolean', capture)
        self.assertNotIn('source.Parent Is selector.Parent', capture)
        order = ('selector.Value2 = code', 'RecalculateTrialYears book', 'source.Calculate',
                 'values = source.Value2', 'CaptureBalancesAgree(', 'If attempt > 2 Then RaiseRunError',
                 'CaptureTrial = values')
        active_trial = capture[capture.index('selector.Value2 = code'):]
        self.assertEqual(sorted(map(active_trial.index, order)), list(map(active_trial.index, order)))
        self.assertLess(capture.index('CaptureBalancesAgree('), capture.index('If attempt = 2 Then Exit For'))
        self.assertIn('RepeatsPriorEnding(values, batch, previousRow)', capture)
        self.assertIn('2 * ANNUAL_CAPACITY - 1', procedure('RepeatsPriorEnding'))
        finish = procedure('FinishTrialCalculation')
        for step in ('Application.CalculateFull', 'Application.CalculationState = xlCalculating', 'CheckCancelled'):
            self.assertIn(step, finish)
        self.assertNotIn('xlPending Then Application.Calculate', finish)
        self.assertNotIn('CalculationState <> xlDone', finish)
        compare = procedure('CaptureBalancesAgree')
        self.assertIn('For period = 1 To years', compare)
        self.assertIn('batch(pairedRow, field)', compare)
        self.assertIn('If Not AnnualClosingsAgree(book, years, detail) Then Exit Function', compare)
        balances = procedure('AnnualClosingsAgree')
        self.assertEqual(int(constant('CLOSING_CHECKS')), len(native_balance_checks.CLOSINGS))
        first_check = native_balance_checks.FIRST_COLUMN
        self.assertEqual(range_boundaries('N5:N5')[0], first_check)
        self.assertIn('Range("N5").Resize(years, CLOSING_CHECKS)', balances)
        self.assertIn('If Not formulas Then', balances)
        self.assertIn('If Not checks.Cells(period, field).HasFormula Then', balances)
        self.assertIn('AtTrialAge(book, period', balances)
        self.assertLess(balances.index('checks.Calculate'), balances.index('differences = checks.Value2'))
        self.assertIn('SameCaptureValue(differences(period, field), 0#)', balances)
        for name, label in (('ENDING_NOMINAL_FIELD', 'Ending portfolio, future dollars'),
                            ('CUT_YEARS_FIELD', 'Years with spending cuts'),
                            ('REFRESH_DATE_FIELD', 'Refresh date')):
            self.assertEqual(int(constant(name)), native_capture.FIELDS.index(label) + 1)
        self.assertEqual(int(constant('FIRST_YEAR_FIELD')), native_capture.ANNUAL_FIRST - 1)
        self.assertEqual(int(constant('ANNUAL_CAPACITY')), native_capture.ANNUAL_CAPACITY)
        self.assertEqual(int(constant('FUNDING_YEAR_WIDTH')), native_funding_paths.FIELD_COUNT)
        closing = [item[0] for item in native_funding_paths.FIELDS].index('Total closing portfolio') + 1
        self.assertEqual(int(constant('FUNDING_BALANCE_FIELD')), closing)

    def test_run_timestamp_is_numeric_and_handles_either_excel_date_system(self):
        run = procedure('RunMarketRisk')
        self.assertIn('metadata(2, 1) = CDbl(Now)', run)
        self.assertIn('If book.Date1904 Then metadata(2, 1) = metadata(2, 1) - 1462', run)
        self.assertNotIn('metadata(2, 1) = Format$', run)


if __name__ == '__main__':
    unittest.main()
