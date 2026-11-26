"""Exercise the disposable funding guard in a local VBA interpreter."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

from desktop_guarded_calculation import overlay as guarded_overlay
from native_funding_proof_overlay import DECLARATIONS, PROCEDURES, instrument, parse_counts
from package_macro import SOURCE


HARNESS = '''
Private Enum RunnerProfileRetry
    ProfileBalanceMismatch = 1
    ProfileRepeatedEnding = 2
End Enum
Private Const ANNUAL_CAPACITY As Long = 126
Private Const ProfilePreparation As Long = 1
Private Const ProfilePrimaryCalculation As Long = 2
Private Const ProfileOrderedCalculation As Long = 3
Private Const ProfileVerifiedCapture As Long = 4
Private mRunning As Boolean
Private mGuardedFirst As Long, mGuardedFull As Long, mOldGuardCalls As Long, mAccepted As Long
Private mAttempts As String

Public Function Probe() As String
    Dim selector As Range, captured As Variant, batch As Variant
    ValidateFundingProofBank ThisWorkbook
    Set selector = ThisWorkbook.Worksheets("Risk_Trials").Range("B3")
    captured = CaptureTrial(ThisWorkbook, selector, selector.Offset(0, 1), 7, "context", False, batch, 0, 0)
    If captured <> 42 Then RaiseRunError "The guard did not return the expected capture."
    Probe = mAttempts & "|" & CStr(mOldGuardCalls) & "|" & CStr(mAccepted)
End Function

Public Sub RemoveProofFormula()
    ThisWorkbook.Worksheets("Trial_Checks").Range("BZ130").ClearContents
End Sub

Private Sub FinishTrialCalculation(Optional ByVal smart As Boolean = False)
    mAttempts = mAttempts & IIf(smart, "S", "F")
End Sub

Private Sub RecalculateTrialYears(ByVal book As Workbook)
    mAttempts = mAttempts & "O"
    book.Worksheets("Risk_Trials").Range("Z1").Value2 = -0.015
End Sub

Private Function CaptureBalancesAgree(ByVal book As Workbook, ByRef values As Variant, ByVal primary As Boolean, _
        ByRef batch As Variant, ByVal pairedRow As Long, ByRef detail As String) As Boolean
    mOldGuardCalls = mOldGuardCalls + 1
    CaptureBalancesAgree = (book.Worksheets("Trial_Checks").Range("N6").Value2 = 0#)
End Function

Private Function RepeatsPriorEnding(ByRef values As Variant, ByRef batch As Variant, ByVal previous As Long) As Boolean
End Function

Private Function AtTrialAge(ByVal book As Workbook, ByVal period As Long, ByVal label As String) As String
    AtTrialAge = label & " in period " & CStr(period)
End Function

Private Sub ValidateCaptureRecord(ByRef values As Variant, ByVal code As Long, ByVal context As String)
    mAccepted = mAccepted + 1
End Sub

Private Function AssumptionRecord(ByVal book As Workbook) As String
    AssumptionRecord = "context"
End Function

Private Sub RaiseRunError(ByVal detail As String)
    Err.Raise 1250, "ProofTest", detail
End Sub

Private Sub CheckCancelled()
End Sub
Private Sub MoveMarketRiskProfile(ByVal stage As Long)
End Sub
Private Sub CountProfileRetry(ByVal reason As RunnerProfileRetry)
End Sub
'''


def procedure(source: str, signature: str) -> str:
    start = source.index(signature)
    end = 'End Function' if 'Function' in signature else 'End Sub'
    return source[start:source.index(end, start) + len(end)]


class NativeFundingProofOverlayTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.published = SOURCE.read_text(encoding='ascii')
        cls.natural = guarded_overlay(cls.published, mode='natural')
        cls.candidate = instrument(cls.natural)
        helpers = (procedure(cls.published, f'Private {signature}') for signature in
                   ('Function WholeNumber(', 'Function IsNumericInput(', 'Sub RequireText('))
        cls.module = '\n'.join(('Option Explicit', DECLARATIONS, HARNESS, PROCEDURES,
                               procedure(cls.candidate, 'Private Function CaptureTrial('), *helpers))

    def fixture(self, disposition: str, margin: float, *, mode: int = 0) -> ExcelApplication:
        """Populate the actual input/output shapes with deterministic stand-ins."""
        workbook = Workbook()
        proof = workbook.active
        proof.title = 'Trial_Checks'
        for first, last in (('BA', 'BD'), ('BG', 'CB'), ('CD', 'CP')):
            for column in range(column_index_from_string(first), column_index_from_string(last) + 1):
                for row in range(6, 131):
                    proof.cell(row, column, '=0')
        proof['BA6'], proof['BB6'], proof['BD6'] = '="Funded"', '=1000', '=999.99'
        proof['BC6'], proof['BG6'] = f'={mode}', f'={mode}'
        proof['BZ6'] = '=Risk_Trials!Z1'
        proof['CP6'] = f'=IF(BZ6>0,"{disposition}","inconclusive")'
        proof['N6'] = '=382897338.323041-382897338.323041'
        workbook.create_sheet('Trial_Inputs')['B70'] = 2
        workbook.create_sheet('Market_Paths')['B3'] = 7
        trials = workbook.create_sheet('Risk_Trials')
        trials['B3'], trials['C3'], trials['Z1'] = 3, '=42', margin
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'proof.xlsx'
            workbook.save(path)
            app = ExcelApplication.open(path, with_vba=False)
        app.add_module(self.module, name='ProofTest')
        return app

    def test_proof_rejects_even_when_old_guard_passes(self) -> None:
        app = self.fixture('earlier-feasible', 200_440.854289053)
        self.assertEqual(app.run('Probe'), 'SFO|3|1')
        values = parse_counts(app.run('ReadFundingProofCounts'))
        self.assertEqual((values['attempts'], values['rejected_attempts'], values['findings'],
                          values['inconclusive']), (3, 2, 2, 1))

    def test_ambiguous_and_malformed_quotes_stay_with_the_old_guard(self) -> None:
        for disposition, margin, mode, category in (
            ('inconclusive', 1, 0, 'inconclusive'), ('grid-required', 1, 0, 'grid_required'),
            ('', 1, 0, 'blank'), ('earlier-feasible', .005, 0, 'malformed'),
            ('earlier-feasible', 1, 1, 'malformed'), ('Earlier-feasible', 1, 0, 'malformed'),
        ):
            with self.subTest(disposition=disposition, margin=margin, mode=mode):
                app = self.fixture(disposition, margin, mode=mode)
                self.assertEqual(app.run('Probe'), 'S|1|1')
                values = parse_counts(app.run('ReadFundingProofCounts'))
                self.assertEqual((values['rejected_attempts'], values[category]), (0, 1))

    def test_missing_formula_aborts_before_selection(self) -> None:
        app = self.fixture('earlier-feasible', 1)
        app.run('RemoveProofFormula')
        with self.assertRaisesRegex(VBARuntimeError, 'missing formulas'):
            app.run('Probe')
        self.assertEqual(app.evaluate('Worksheets("Risk_Trials").Range("B3").Value2'), 3)
        self.assertEqual(app.run('ReadFundingProofCounts'), '')
        self.assertLess(self.candidate.index('ValidateFundingProofBank book'),
                        self.candidate.index('SaveExcelState settings'))

    def test_receipt_and_natural_driver_contract_stay_distinct(self) -> None:
        for name in ('ReadMarketRiskProfile', 'ReadGuardedCalculationCounts'):
            original = procedure(self.natural, f'Public Function {name}(')
            self.assertIn(original, self.candidate)
        for packet in (
            '', 'FundingProofCounts/v1|1|2|0|1|1|0|0|0|0',
            'FundingProofCounts/v1|1|1|2|1|0|0|0|0|0',
            'FundingProofCounts/v1|1|1|0|0|0|0|0|0|0',
            'FundingProofCounts/v1|1|126|0|0|126|0|0|0|0',
            'FundingProofCounts/v1|١|1|0|0|1|0|0|0|0',
        ):
            with self.subTest(packet=packet), self.assertRaises(ValueError):
                parse_counts(packet)
        with self.assertRaisesRegex(ValueError, 'natural guarded overlay'):
            instrument(self.published)
