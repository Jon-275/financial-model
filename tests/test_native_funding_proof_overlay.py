"""Exercise the disposable funding guard in a local VBA interpreter."""

from dataclasses import replace
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from pyopenvba.apps.excel import ExcelApplication, Range
from pyopenvba.exceptions import VBARuntimeError

from desktop_guarded_calculation import overlay as guarded_overlay
from native_funding_proof_overlay import DECLARATIONS, PROCEDURES, instrument, instrument_canary, parse_counts
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

CANARY_SETUP = '''
Public Sub ProtectProofBank()
    With ThisWorkbook.Worksheets("Trial_Checks")
        .Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=False, AllowFiltering:=True
        .EnableSelection = 1
    End With
End Sub
'''


def procedure(source: str, signature: str) -> str:
    """Extract a single VBA routine from a reviewed driver."""
    start = source.index(signature)
    end = 'End Function' if 'Function' in signature else 'End Sub'
    return source[start:source.index(end, start) + len(end)]


class NativeFundingProofOverlayTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        """Compile the original capture loop with isolated stage and workbook spies."""
        cls.published = SOURCE.read_text(encoding='ascii')
        cls.natural = guarded_overlay(cls.published, mode='natural')
        cls.candidate = instrument(cls.natural)
        helpers = tuple(procedure(cls.published, f'Private {signature}') for signature in
                        ('Function WholeNumber(', 'Function IsNumericInput(', 'Sub RequireText('))
        common = ('Option Explicit', DECLARATIONS, HARNESS, *helpers)
        cls.module = '\n'.join((*common, PROCEDURES, procedure(cls.candidate, 'Private Function CaptureTrial(')))
        canary = instrument_canary(cls.natural)
        canary_start = canary.index('Private Sub ValidateFundingProofBank(')
        protection = (procedure(cls.published, f'Private {signature}') for signature in
                      ('Sub PrepareProtection(', 'Sub ApplyProtection(', 'Function RestoreProtection('))
        cls.canary_module = '\n'.join((*common, 'Private mProofCanaryInjected As Long', CANARY_SETUP,
                                      canary[canary_start:], procedure(canary, 'Private Function CaptureTrial('),
                                      *protection))

    def fixture(self, disposition: str, margin: float, *, mode: int = 0, canary: bool = False) -> ExcelApplication:
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
        workbook.create_sheet('Trial_Inputs')['B70'] = 122 if canary else 2
        workbook.create_sheet('Market_Paths')['B3'] = 7
        trials = workbook.create_sheet('Risk_Trials')
        trials['B3'], trials['C3'], trials['Z1'] = 3, '=42', margin
        if canary:
            trials['Z2'] = 3_761_432.87464338
            proof['CP6'] = '="inconclusive"'
            proof['BA126'], proof['BB126'], proof['BD126'] = '="Funded"', '=Risk_Trials!Z2', '=BB126-0.01'
            proof['BZ126'] = '=IF(BB126>3900000,200440.854289053,-0.00872003054246306)'
            proof['CP126'] = '=IF(BZ126>0.005,"earlier-feasible","inconclusive")'
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'proof.xlsx'
            workbook.save(path)
            app = ExcelApplication.open(path, with_vba=False)
        app.add_module(self.canary_module if canary else self.module, name='ProofTest')
        if canary:
            app.run('ProtectProofBank')
        return app

    def proof_state(self, app: ExcelApplication) -> tuple[object, ...]:
        """Read the original claim and the user-visible proof sheet protection settings."""
        sheet = 'Worksheets("Trial_Checks")'
        members = ('Range("BB126").Formula', 'ProtectDrawingObjects', 'ProtectContents',
                   'ProtectScenarios', 'ProtectionMode', 'EnableSelection', 'Protection.AllowFiltering')
        return tuple(app.evaluate(f'{sheet}.{member}') for member in members)

    def test_proof_rejects_even_when_old_guard_passes(self) -> None:
        app = self.fixture('earlier-feasible', 200_440.854289053)
        self.assertEqual(app.run('Probe'), 'SFO|3|1')
        values = parse_counts(app.run('ReadFundingProofCounts'))
        self.assertEqual((values['attempts'], values['rejected_attempts'], values['findings'],
                          values['inconclusive']), (3, 2, 2, 1))

    def test_canary_rejects_two_injections_and_restores_the_protected_claim(self) -> None:
        """Run the actual canary and accept its capture only after ordered recovery."""
        app = self.fixture('inconclusive', 0, canary=True)
        prior = self.proof_state(app)
        self.assertEqual(prior, ('=Risk_Trials!Z2', True, True, False, False, 1, True))
        self.assertEqual(app.run('Probe'), 'SFO|3|1')
        self.assertEqual(app.run('ReadFundingProofInjectionCount'), '2')
        counts = parse_counts(app.run('ReadFundingProofCounts'))
        self.assertEqual((counts['attempts'], counts['rejected_attempts'], counts['findings']), (3, 2, 2))
        self.assertEqual(self.proof_state(app), prior)

    def test_canary_restores_the_protected_claim_after_a_calculation_error(self) -> None:
        """Fail the second quote after the literal write, inside the real VBA error handler."""
        app = self.fixture('inconclusive', 0, canary=True)
        prior = self.proof_state(app)
        quote_calculations = 0

        def calculate_with_fault(area: Range) -> object:
            """Simulate an Excel host error on the second injected quote calculation."""
            nonlocal quote_calculations
            if area.sheet.name == 'Trial_Checks' and area.Address(False, False) == 'BD6:BG126':
                quote_calculations += 1
                if quote_calculations == 2:
                    self.assertAlmostEqual(app.evaluate('Worksheets("Trial_Checks").Range("BB126").Value2'),
                                           3_991_289.26141221)
                    self.assertTrue(app.evaluate('Worksheets("Trial_Checks").ProtectionMode'))
                    raise VBARuntimeError(1004, 'Injected quote calculation error')
            return Range.Calculate(area)

        calculate = Range._vba_members['calculate']
        # Inject the host failure without changing the VBA being exercised.
        with patch.dict(Range._vba_members, calculate=replace(calculate, getter=calculate_with_fault)):
            with self.assertRaisesRegex(VBARuntimeError, 'Injected quote calculation error') as raised:
                app.run('Probe')
        self.assertEqual((quote_calculations, raised.exception.number, raised.exception.source),
                         (2, 1004, 'StaleFundingClaimCanary'))
        self.assertEqual(app.run('ReadFundingProofInjectionCount'), '1')
        self.assertEqual(self.proof_state(app), prior)

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
