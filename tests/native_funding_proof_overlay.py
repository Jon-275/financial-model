"""Overlay a disposable lower-funding-quote check on the native runner."""

from argparse import ArgumentParser
from os import link
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from desktop_benchmark import digest
from desktop_evidence import vba_identity
from desktop_guarded_calculation import overlay as guarded_overlay
from desktop_stage_profile import RUNNER, VBA_PART
from native_funding_proof_package import prepare_package, verify_package
from pyopenvba import ExcelFile
from vba_declarations import audit_vba_declarations

COUNT_FIELDS = ('attempts', 'rows', 'rejected_attempts', 'findings', 'inconclusive',
                'grid_required', 'blank', 'malformed', 'selector_skips')
DECLARATIONS = '''
Private mFundingProofReady As Boolean
Private mProofAttempts As Long, mProofRows As Long, mProofRejected As Long, mProofFindings As Long
Private mProofInconclusive As Long, mProofGrid As Long, mProofBlank As Long, mProofMalformed As Long
Private mProofSelectorSkips As Long
'''

PROCEDURES = '''
Private Sub ValidateFundingProofBank(ByVal book As Workbook)
    Dim proof As Worksheet, band As Variant, formulas As Variant
    mFundingProofReady = False
    mProofAttempts = 0: mProofRows = 0: mProofRejected = 0: mProofFindings = 0
    mProofInconclusive = 0: mProofGrid = 0: mProofBlank = 0: mProofMalformed = 0: mProofSelectorSkips = 0
    Set proof = book.Worksheets("Trial_Checks")
    For Each band In Array("BA6:BD130", "BG6:CB130", "CD6:CP130")
        formulas = proof.Range(CStr(band)).HasFormula
        If VarType(formulas) <> vbBoolean Then RaiseRunError "The test funding proof bank has missing formulas."
        If Not CBool(formulas) Then RaiseRunError "The test funding proof bank has missing formulas."
    Next band
    mFundingProofReady = True
End Sub

Private Function FundingProofTrialMatches(ByVal book As Workbook, ByVal code As Long) As Boolean
    Dim selected As Variant, calculated As Variant
    selected = book.Worksheets("Risk_Trials").Range("B3").Value2
    calculated = book.Worksheets("Market_Paths").Range("B3").Value2
    If Not IsNumericInput(selected) Or Not IsNumericInput(calculated) Then Exit Function
    FundingProofTrialMatches = (CDbl(selected) = code And CDbl(calculated) = code)
End Function

Private Function NumericFundingCounterexample(ByVal status As Variant, ByVal gross As Variant, _
        ByVal mode As Variant, ByVal lower As Variant, ByVal quotedMode As Variant, ByVal margin As Variant) As Boolean
    If VarType(status) <> vbString Then Exit Function
    If StrComp(CStr(status), "Funded", vbBinaryCompare) <> 0 Then Exit Function
    If Not IsNumericInput(gross) Or Not IsNumericInput(lower) Or Not IsNumericInput(margin) Then Exit Function
    If Not IsNumericInput(mode) Or Not IsNumericInput(quotedMode) Then Exit Function
    If CDbl(mode) <> 0# Or CDbl(quotedMode) <> 0# Then Exit Function
    If CDbl(gross) < 0.01 Or CDbl(lower) < 0# Or CDbl(gross) - CDbl(lower) <= 0.005 Then Exit Function
    NumericFundingCounterexample = (CDbl(margin) > 0.005)
End Function

Private Function FundingProofFinding(ByVal book As Workbook, ByVal code As Long) As Long
    Dim proof As Worksheet, years As Long, lastRow As Long, year As Long
    Dim observed As Variant, disposition As Variant
    mProofAttempts = mProofAttempts + 1
    If Not FundingProofTrialMatches(book, code) Then
        mProofSelectorSkips = mProofSelectorSkips + 1
        Exit Function
    End If
    years = WholeNumber(book.Worksheets("Trial_Inputs").Range("B70").Value2, 2, ANNUAL_CAPACITY, "The trial horizon is invalid.")
    lastRow = 4 + years
    Set proof = book.Worksheets("Trial_Checks")
    With proof
        .Range("BA6:BC" & lastRow).Calculate
        .Range("BD6:BG" & lastRow).Calculate
        .Range("CM6:CO" & lastRow).Calculate
        .Range("BH6:CL" & lastRow).Calculate
        .Range("CP6:CP" & lastRow).Calculate
        observed = .Range("BA6:CP" & lastRow).Value2
    End With
    If Not FundingProofTrialMatches(book, code) Then
        mProofSelectorSkips = mProofSelectorSkips + 1
        Exit Function
    End If
    For year = 1 To UBound(observed, 1)
        mProofRows = mProofRows + 1
        disposition = observed(year, 42) ' CP
        If IsEmpty(disposition) Then
            mProofBlank = mProofBlank + 1
        ElseIf VarType(disposition) <> vbString Then
            mProofMalformed = mProofMalformed + 1
        Else
            Select Case CStr(disposition)
                Case "earlier-feasible"
                    ' BA status, BB selected, BC certified mode, BD lower, BG quoted mode, BZ margin.
                    If NumericFundingCounterexample(observed(year, 1), observed(year, 2), observed(year, 3), _
                            observed(year, 4), observed(year, 7), observed(year, 26)) Then
                        mProofFindings = mProofFindings + 1
                        If FundingProofFinding = 0 Then FundingProofFinding = year + 1
                    Else
                        mProofMalformed = mProofMalformed + 1
                    End If
                Case "inconclusive": mProofInconclusive = mProofInconclusive + 1
                Case "grid-required": mProofGrid = mProofGrid + 1
                Case "": mProofBlank = mProofBlank + 1
                Case Else: mProofMalformed = mProofMalformed + 1
            End Select
        End If
    Next year
    If FundingProofFinding > 0 Then mProofRejected = mProofRejected + 1
End Function

Public Function ReadFundingProofCounts() As String
    If mRunning Or Not mFundingProofReady Then Exit Function
    ReadFundingProofCounts = "FundingProofCounts/v1|" & CStr(mProofAttempts) & "|" & CStr(mProofRows) & "|" _
        & CStr(mProofRejected) & "|" & CStr(mProofFindings) & "|" & CStr(mProofInconclusive) & "|" _
        & CStr(mProofGrid) & "|" & CStr(mProofBlank) & "|" & CStr(mProofMalformed) & "|" & CStr(mProofSelectorSkips)
End Function
'''


def replace(source: str, before: str, after: str) -> str:
    """Require exactly one match for each reviewed runner anchor."""
    if source.count(before) != 1:
        raise ValueError(f'Expected one test overlay anchor for {before.strip().splitlines()[0]!r}.')
    return source.replace(before, after, 1)


def instrument(source: str) -> str:
    """Layer the proof check on the existing natural three-attempt driver."""
    vba = source.replace('\r\n', '\n').rstrip() + '\n'
    if 'Option Compare Text' in vba:
        raise ValueError('The disposable proof uses binary disposition comparisons.')
    if ('Public Function ReadGuardedCalculationCounts()' not in vba or
            'mGuardedCanary = mGuardedCanary + 1' in vba):
        raise ValueError('Compose the existing natural guarded overlay before the funding proof.')
    vba = replace(vba, 'Option Explicit\n', 'Option Explicit\n' + DECLARATIONS)
    anchor = '    Set trialCurrent = LocalRange(trials, "RiskTrialCurrent")\n'
    vba = replace(vba, anchor, '    ValidateFundingProofBank book\n' + anchor)
    anchor = '    CheckCancelled\n    For attempt = 1 To 3\n'
    vba = replace(vba, anchor, '    Dim proofPeriod As Long, existingChecksPassed As Boolean\n' + anchor)
    anchor = '        If CaptureBalancesAgree(book, values, primary, batch, pairedRow, detail) Then\n'
    vba = replace(vba, anchor, '''        proofPeriod = FundingProofFinding(book, code)
        CheckCancelled
        RequireText AssumptionRecord(book), context, "The assumptions changed during the funding proof. Run market risk again."
        existingChecksPassed = CaptureBalancesAgree(book, values, primary, batch, pairedRow, detail)
        If proofPeriod > 0 Then
            detail = AtTrialAge(book, proofPeriod, "The selected draw against a lower funded same-policy quote")
            existingChecksPassed = False
        End If
        If existingChecksPassed Then
''')
    vba += '\n' + PROCEDURES
    audit_vba_declarations(vba)
    return vba


def parse_counts(payload: str) -> dict[str, int]:
    """Require complete and internally consistent native proof counters."""
    fields = payload.split('|')
    if len(fields) != 1 + len(COUNT_FIELDS) or fields[0] != 'FundingProofCounts/v1':
        raise ValueError('Missing test-only funding proof counters.')
    if any(not count.isascii() or not count.isdecimal() or len(count) > 7 for count in fields[1:]):
        raise ValueError('Invalid test-only funding proof counter.')
    counts = dict(zip(COUNT_FIELDS, map(int, fields[1:]), strict=True))
    attempts, rows, rejected, findings, inconclusive, grid, blank, bad, skips = counts.values()
    if (skips > attempts or rejected > attempts - skips or rejected > findings or
            rows != findings + inconclusive + grid + blank + bad or
            not attempts - skips <= rows <= 125 * (attempts - skips) or bool(findings) != bool(rejected)):
        raise ValueError('Inconsistent funding proof counters.')
    return counts


def prepare_candidate(source: Path, destination: Path) -> None:
    """Build a new proof workbook with a verified bank and isolated runner."""
    if destination.exists() or destination.is_symlink():
        raise FileExistsError(destination)
    with TemporaryDirectory(dir=destination.parent) as folder:
        bank, project, candidate = (Path(folder) / name for name in ('bank.xlsm', 'project.xlsm', 'candidate.xlsm'))
        prepare_package(source, bank)
        verify_package(source, bank)
        with ExcelFile(bank) as original:
            expected = instrument(guarded_overlay(original.get_module(RUNNER), mode='natural'))
            original.set_module(RUNNER, expected.replace('\n', '\r\n'))
            original.save(project)
        with ZipFile(bank) as before, ZipFile(project) as compiled, ZipFile(candidate, 'x') as output:
            for part in before.infolist():
                archive = compiled if part.filename == VBA_PART else before
                output.writestr(part, archive.read(part.filename))
        with ZipFile(bank) as before, ZipFile(candidate) as after:
            if before.namelist() != after.namelist() or [
                    name for name in before.namelist() if before.read(name) != after.read(name)] != [VBA_PART]:
                raise ValueError('The runner layer changed more than the existing VBA project.')
        before_modules, before_libraries = vba_identity(bank)
        after_modules, after_libraries = vba_identity(candidate)
        before_modules.pop(RUNNER)
        after_modules.pop(RUNNER)
        if before_modules != after_modules or before_libraries != after_libraries:
            raise ValueError('The runner layer changed unrelated VBA modules or library references.')
        with ExcelFile(candidate) as output:
            if output.get_module(RUNNER).splitlines() != expected.splitlines():
                raise ValueError('The stored VBA differs from the reviewed proof runner.')
        link(candidate, destination)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('destination', type=Path)
    opts = parser.parse_args()
    prepare_candidate(opts.source, opts.destination)
    print(f'{opts.destination}\nInput SHA-256 {digest(opts.destination)}')
