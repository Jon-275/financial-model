"""Time the two reuse checks on a disposable desktop benchmark workbook."""

from argparse import ArgumentParser
from math import isclose, isfinite
from pathlib import Path
import re
from subprocess import CalledProcessError
from typing import Any
from zipfile import BadZipFile

from desktop_benchmark import read_json, write_json
import desktop_stage_profile as stage
from vba_declarations import audit_vba_declarations


PROFILE = 'RunnerReuseProfile/v1'
EVIDENCE = 'RunnerReuseEvidence/v1'
DECLARATIONS = '''
Private Enum ReuseVerificationCheck
    ReuseFixedPath = 1
    ReusePartnerAdmission
End Enum
Private mReuseSeconds(1 To 2) As Double
Private mReuseCalls(1 To 2) As Long
Private mReusePasses(1 To 2) As Long
'''
PROCEDURES = f'''
Private Sub RecordReuseCheck(ByVal check As ReuseVerificationCheck, ByVal passed As Boolean)
    Dim beforeSeconds As Double
    If mProfileActive <> ProfileReuseVerification Then Exit Sub
    beforeSeconds = mProfileSeconds(ProfileReuseVerification)
    MoveMarketRiskProfile ProfilePreparation
    If mProfileActive <> ProfilePreparation Then Exit Sub
    mReuseSeconds(check) = mReuseSeconds(check) + (mProfileSeconds(ProfileReuseVerification) - beforeSeconds)
    mReuseCalls(check) = mReuseCalls(check) + 1
    If passed Then mReusePasses(check) = mReusePasses(check) + 1
End Sub

Public Function ReadReuseVerificationProfile() As String
    Dim payload As String, check As Long
    payload = "{PROFILE}"
    If mRunning Or mProfileOutcome <> "complete" Then
        ReadReuseVerificationProfile = payload & "|unavailable"
        Exit Function
    End If
    payload = payload & "|" & PROFILE_CLOCK & "|" & Trim$(Str$(mProfileSerial))
    For check = ReuseFixedPath To ReusePartnerAdmission
        payload = payload & "|" & Trim$(Str$(mReuseSeconds(check))) _
            & "|" & Trim$(Str$(mReuseCalls(check))) & "|" & Trim$(Str$(mReusePasses(check)))
    Next check
    ReadReuseVerificationProfile = payload
End Function
'''


def instrument(source: str) -> str:
    """Attribute each existing reuse timer interval while preserving its aggregate."""
    result = stage.instrument(source)
    declarations = 'Private mProfileFull As Long\n'
    start = 'Private Sub StartMarketRiskProfile()\n    Dim failure As Long\n'
    replacements = [(declarations, declarations + DECLARATIONS),
                    (start, start + '    Erase mReuseSeconds\n    Erase mReuseCalls\n    Erase mReusePasses\n')]
    finish = '                MoveMarketRiskProfile ProfilePreparation\n'
    for check, assignment, function in (
        ('ReuseFixedPath', 'fixedUnchanged(index)', 'FixedPathAllowsReuse(cutProof, captured, book)'),
        ('ReusePartnerAdmission', 'reused', 'CanReuseFixedTrial(book, batch, ids, pairedRow, code, context, width)'),
    ):
        call = f'                {assignment} = {function}\n'
        replacements.append((call + finish, call + f'                RecordReuseCheck {check}, {assignment}\n'))
    for old, new in replacements:
        if result.count(old) != 1:
            raise ValueError(f'Expected one reuse profile anchor for {old.strip().splitlines()[0]}.')
        result = result.replace(old, new, 1)
    result += '\n' + PROCEDURES
    audit_vba_declarations(result)
    return result


def parse(payload: str, *, pairs: int, profile: dict[str, Any]) -> dict[str, dict[str, float | int]]:
    """Bind a complete receipt to the stage profile's clock, batch and reuse total."""
    if not isinstance(payload, str) or len(payload) > 512:
        raise ValueError('Expected one bounded reuse profile.')
    fields = payload.strip().split('|')
    if fields == [PROFILE, 'unavailable']:
        raise ValueError('The reuse profile is unavailable.')
    if len(fields) != 9 or fields[0] != PROFILE:
        raise ValueError('The reuse profile has unexpected fields.')
    if fields[1] != profile['clock']:
        raise ValueError('The reuse clock must match the stage profile.')
    if not all(stage.DECIMAL.fullmatch(fields[index]) for index in (2, 3, 6)):
        raise ValueError('The reuse timestamp and durations must be decimal numbers.')
    serial, fixed_seconds, partner_seconds = (float(fields[index]) for index in (2, 3, 6))
    if not all(isfinite(value) and value >= 0 for value in (serial, fixed_seconds, partner_seconds)):
        raise ValueError('The reuse timestamp and durations must be finite and nonnegative.')
    if serial != profile['batch_serial']:
        raise ValueError('The reuse profile came from a different batch.')
    counts = fields[4:6] + fields[7:9]
    if not all(stage.INTEGER.fullmatch(count) for count in counts):
        raise ValueError('The reuse counts must be nonnegative integers.')
    fixed_calls, fixed_passes, partner_calls, partner_passes = map(int, counts)
    # Each random pair and each of the 13 early-loss pairs gets one fixed check.
    if (fixed_calls != pairs + 13 or not 0 <= partner_passes <= partner_calls == fixed_passes <= fixed_calls
            or partner_passes != profile['counters']['reused'] or (not partner_calls and partner_seconds > 0)):
        raise ValueError('The reuse counts do not reconcile to the completed fixture.')
    # The two accumulators are serialized separately from the shared timer.
    if not isclose(fixed_seconds + partner_seconds, profile['stages']['reuse_verification'],
                   rel_tol=0, abs_tol=.000001):
        raise ValueError('The reuse timing partition does not match the stage profile.')
    return {'fixed_path': {'seconds': fixed_seconds, 'calls': fixed_calls, 'passes': fixed_passes},
            'partner_admission': {'seconds': partner_seconds, 'calls': partner_calls, 'passes': partner_passes}}


def record(directory: Path, name: str, timing: dict[str, Any], profile: str, payload: str) -> dict[str, Any]:
    """Record the extra receipt beside the independently reconciled stage evidence."""
    verified = stage.record(directory, name, timing, profile, transform=instrument)
    reuse = parse(payload, pairs=verified['pairs'], profile=verified['profile'])
    return {'schema': EVIDENCE, 'stage': verified, 'payload': payload.strip(), 'reuse': reuse}


def check(path: Path) -> dict[str, Any]:
    """Replay both receipts from the saved workbook and pinned reference."""
    saved = read_json(path)
    if saved.get('schema') != EVIDENCE or not isinstance(saved.get('stage'), dict):
        raise ValueError('Expected reuse profile evidence with the supported schema.')
    regular = saved['stage']
    replay = record(path.parent, regular['capture'], regular['run'], regular['payload'], saved['payload'])
    if replay != saved:
        raise ValueError('The recorded reuse profile differs from its replay.')
    return replay


def main() -> int:
    """Prepare, record or replay the optional reuse diagnostic."""
    parser = ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    commands.add_parser('prepare').add_argument('directory', type=Path)
    capture = commands.add_parser('record')
    capture.add_argument('directory', type=Path)
    capture.add_argument('workbook')
    capture.add_argument('timing', type=Path)
    capture.add_argument('profile', type=Path)
    capture.add_argument('reuse', type=Path)
    capture.add_argument('--output', required=True)
    commands.add_parser('check').add_argument('evidence', type=Path)
    args = parser.parse_args()
    try:
        if args.command == 'prepare':
            print(stage.prepare(args.directory, transform=instrument))
        elif args.command == 'record':
            if not re.fullmatch(r'reuse-evidence-[A-Za-z0-9_-]+\.json', args.output, re.ASCII):
                raise ValueError('Use a reuse-evidence-<id>.json filename.')
            with args.profile.open(encoding='ascii') as profile, args.reuse.open(encoding='ascii') as receipt:
                result = record(args.directory, args.workbook, read_json(args.timing), profile.read(1025), receipt.read(513))
            write_json(args.directory / args.output, result)
            print(result['reuse'])
        else:
            print(check(args.evidence)['reuse'])
    except (AssertionError, OSError, ValueError, KeyError, BadZipFile, CalledProcessError) as error:
        parser.exit(1, f'{error}\n')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
