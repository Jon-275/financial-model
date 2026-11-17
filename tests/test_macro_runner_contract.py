"""Check the VBA/worksheet contract and its permitted automation surface.

These source checks complement native Excel compilation and run tests. They do
not establish that Excel executed the macro or that cancellation worked there.
"""
from collections import defaultdict
from pathlib import Path
import re
from types import SimpleNamespace
import unittest

from openpyxl.utils.cell import range_boundaries
import native_capture
import native_funding_paths
import native_results

ROOT = Path(__file__).resolve().parents[1]
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


class MacroRunnerContractTests(unittest.TestCase):
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

    def test_ranking_fields_and_release_version_match_the_worksheet_compiler(self):
        expected = {
            'VALID_FIELD': 'Calculation valid',
            'FAILURE_AGE_FIELD': 'First shortfall age',
            'MAX_UNPAID_FIELD': 'Largest unpaid balance, today’s dollars',
            'ENDING_REAL_FIELD': 'Ending portfolio, today’s dollars',
        }
        for name, label in expected.items():
            self.assertEqual(int(constant(name)), native_capture.FIELDS.index(label) + 1)
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

    def test_cancellation_cleanup_and_early_failure_keep_explicit_guards(self):
        run = procedure('RunMarketRisk')
        self.assertIn('If mRunning Then', run)
        self.assertIn('Application.EnableCancelKey = xlErrorHandler', run)
        self.assertIn('If failureNumber = 18 Then', run)
        self.assertIn('Resume CleanExit', run)
        cleanup = run.split('CleanExit:', 1)[1]
        self.assertRegex(cleanup, r'If savedSettings Then\s+Application.EnableCancelKey = xlDisabled')
        self.assertIn('If savedSelector And selectorTouched And Not selectorRestored Then', cleanup)
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

    def test_run_timestamp_is_numeric_and_handles_either_excel_date_system(self):
        run = procedure('RunMarketRisk')
        self.assertIn('metadata(2, 1) = CDbl(Now)', run)
        self.assertIn('If book.Date1904 Then metadata(2, 1) = metadata(2, 1) - 1462', run)
        self.assertNotIn('metadata(2, 1) = Format$', run)


if __name__ == '__main__':
    unittest.main()
