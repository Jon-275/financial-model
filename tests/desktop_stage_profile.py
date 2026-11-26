"""Prepare and verify a local stage profile of a pinned fictional Excel fixture."""

from __future__ import annotations

import argparse
from collections.abc import Callable
from math import fsum, isfinite
from os import link
from pathlib import Path
import re
import subprocess
from tempfile import TemporaryDirectory
from typing import Any
from zipfile import BadZipFile, ZipFile

from pyopenvba import ExcelFile

from desktop_benchmark import INPUT, ROOT, STRESS_COUNT, digest, load_fixture, read_json, write_json
from desktop_evidence import capture_file, compare_capture, reference_file, validate_run, vba_identity
from vba_declarations import audit_vba_declarations


CANDIDATE = 'desktop-profile-input.xlsm'
RUNNER = 'MarketRiskRunner'
VBA_PART = 'xl/vbaProject.bin'
FRAGMENT = ROOT / 'tools/vba/tests/RunnerStageProfile.fragment'
PROFILE = 'RunnerStageProfile/v1'
EVIDENCE = 'RunnerStageEvidence/v1'
CLOCKS = {'mac': 'mac-raw', 'windows': 'windows-qpc'}
STAGES = ('preflight', 'preparation', 'primary_calculation', 'verified_capture', 'ordered_calculation',
          'ranking', 'reuse_verification', 'progress', 'publish', 'restore')
COUNTERS = ('computed', 'reused', 'stress', 'ranked', 'full_attempts', 'retry_balance', 'retry_repeated_ending')
DECIMAL = re.compile(r'-?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?', re.ASCII)
INTEGER = re.compile(r'[0-9]{1,4}', re.ASCII)
RECORD_NAME = re.compile(r'profile-evidence-[A-Za-z0-9_-]+\.json', re.ASCII)

# Each replacement is scoped to one exact source anchor. Any drift requires explicit review.
PROBES = (
    ('    On Error GoTo RunFailed\n', '    On Error GoTo RunFailed\n    StartMarketRiskProfile\n'),
    ('    For index = 1 To total\n',
     '    MoveMarketRiskProfile ProfilePreparation\n    For index = 1 To total\n'),
    ('            lastCalculatedRow = index\n',
     '            MoveMarketRiskProfile ProfilePreparation\n            lastCalculatedRow = index\n'),
    ('                reused = CanReuseFixedTrial(book, batch, ids, pairedRow, code, context, width)\n',
     '                MoveMarketRiskProfile ProfileReuseVerification\n'
     '                reused = CanReuseFixedTrial(book, batch, ids, pairedRow, code, context, width)\n'
     '                MoveMarketRiskProfile ProfilePreparation\n'),
    ('                fixedUnchanged(index) = FixedPathAllowsReuse(cutProof, captured, book)\n',
     '                MoveMarketRiskProfile ProfileReuseVerification\n'
     '                fixedUnchanged(index) = FixedPathAllowsReuse(cutProof, captured, book)\n'
     '                MoveMarketRiskProfile ProfilePreparation\n'),
    ('        ShowProgress index, total + 3, started\n',
     '        CountProfileTrial code, reused\n        ShowProgress index, total + 3, started\n'),
    ('    SortPaths order, batch, ids, 1, count\n',
     '    MoveMarketRiskProfile ProfileRanking\n    SortPaths order, batch, ids, 1, count\n'),
    ('        For column = 1 To pathWidth\n',
     '        MoveMarketRiskProfile ProfileRanking\n        For column = 1 To pathWidth\n'),
    ('        chosenIDs(caseIndex, 1) = code\n',
     '        chosenIDs(caseIndex, 1) = code\n        mProfileRanked = mProfileRanked + 1\n'),
    ('    RestoreInspection book, selector, previousSelector\n    selectorRestored = True\n',
     '    MoveMarketRiskProfile ProfileRestore\n'
     '    RestoreInspection book, selector, previousSelector\n    selectorRestored = True\n'),
    ('    oldBatch = trialResults.Value2\n',
     '    MoveMarketRiskProfile ProfilePublish\n    oldBatch = trialResults.Value2\n'),
    ('    If publishing Then\n',
     '    MoveMarketRiskProfile ProfileRestore\n    If publishing Then\n'),
    ('    On Error GoTo 0\n    If Not completed Or Len(cleanupText) > 0 Then\n',
     '    On Error GoTo 0\n'
     '    If completed And Len(cleanupText) = 0 Then\n'
     '        FinishMarketRiskProfile "complete", CDbl(metadata(2, 1))\n'
     '    ElseIf failureNumber = 18 And Len(cleanupText) = 0 Then\n'
     '        FinishMarketRiskProfile "cancelled"\n'
     '    Else\n'
     '        FinishMarketRiskProfile "failed"\n'
     '    End If\n'
     '    If Not completed Or Len(cleanupText) > 0 Then\n'),
    ('    Dim values As Variant, attempt As Long, detail As String\n',
     '    Dim values As Variant, attempt As Long, detail As String, retryReason As RunnerProfileRetry\n'),
    ('        selector.Value2 = code\n',
     '        MoveMarketRiskProfile ProfilePreparation\n        selector.Value2 = code\n'),
    ('            FinishTrialCalculation\n',
     '            MoveMarketRiskProfile ProfilePrimaryCalculation\n'
     '            mProfileFull = mProfileFull + 1\n'
     '            FinishTrialCalculation\n'),
    ('            RecalculateTrialYears book\n        End If\n        source.Calculate\n',
     '            MoveMarketRiskProfile ProfileOrderedCalculation\n'
     '            CountProfileRetry retryReason\n'
     '            RecalculateTrialYears book\n        End If\n'
     '        MoveMarketRiskProfile ProfileVerifiedCapture\n        source.Calculate\n'),
    ('            If attempt = 2 Then Exit For\n        End If\n    Next attempt\n',
     '            If attempt = 2 Then Exit For\n'
     '            retryReason = ProfileRepeatedEnding\n        Else\n'
     '            retryReason = ProfileBalanceMismatch\n        End If\n    Next attempt\n'),
    ('    Dim elapsed As Double\n    elapsed = Timer - started\n',
     '    Dim elapsed As Double, previousStage As Long\n'
     '    previousStage = mProfileActive\n'
     '    MoveMarketRiskProfile ProfileProgress\n'
     '    elapsed = Timer - started\n'),
    ('    DoEvents\n    CheckCancelled\nEnd Sub\n',
     '    DoEvents\n    CheckCancelled\n    MoveMarketRiskProfile previousStage\nEnd Sub\n'),
)


def instrument(source: str) -> str:
    """Apply the reviewed test fragment without patching the tracked runner source."""
    source = source.replace('\r\n', '\n').rstrip() + '\n'
    declarations, marker, procedures = FRAGMENT.read_text(encoding='ascii').partition(
        "' === Diagnostic procedures ===\n")
    if not marker:
        raise ValueError('The VBA diagnostic fragment has no declaration boundary.')
    for old, new in (('Option Explicit\n', 'Option Explicit\n\n' + declarations), *PROBES):
        if source.count(old) != 1:
            raise ValueError(f'Expected one VBA diagnostic anchor for {old.strip().splitlines()[0]!r}.')
        source = source.replace(old, new, 1)
    result = source + '\n' + procedures
    audit_vba_declarations(result)
    return result


def check_candidate(directory: Path, *, path: Path | None = None,
                    transform: Callable[[str], str] = instrument) -> Path:
    """Bind the test build to the deterministic overlay and all original workbook parts."""
    original = directory / INPUT
    candidate = path if path is not None else directory / CANDIDATE
    if candidate.is_symlink() or not candidate.is_file():
        raise ValueError('Prepare a local diagnostic workbook before recording a profile.')
    with ZipFile(original) as before, ZipFile(candidate) as after:
        if (before.namelist() != after.namelist()
                or any(before.read(name) != after.read(name) for name in before.namelist() if name != VBA_PART)):
            raise ValueError('The diagnostic input changed workbook parts outside its VBA project.')
    before_modules, before_libraries = vba_identity(original)
    after_modules, after_libraries = vba_identity(candidate)
    before_modules.pop(RUNNER)
    after_modules.pop(RUNNER)
    if before_modules != after_modules or before_libraries != after_libraries:
        raise ValueError('The diagnostic input changed other VBA modules or libraries.')
    with ExcelFile(original) as before, ExcelFile(candidate) as after:
        if transform(before.get_module(RUNNER)).splitlines() != after.get_module(RUNNER).splitlines():
            raise ValueError('The diagnostic input differs from the reviewed stage overlay.')
    return candidate


def prepare(directory: Path, *, transform: Callable[[str], str] = instrument) -> Path:
    """Publish a new disposable diagnostic workbook without overwriting fixture files."""
    load_fixture(directory)
    original = directory / INPUT
    destination = directory / CANDIDATE
    if destination.exists():
        raise FileExistsError(destination)
    with TemporaryDirectory(dir=directory) as temporary:
        package = Path(temporary) / 'package.xlsm'
        candidate = Path(temporary) / CANDIDATE
        with ExcelFile(original) as book:
            code = transform(book.get_module(RUNNER)).replace('\n', '\r\n')
            book.set_module(RUNNER, code)
            book.save(package)
        with ZipFile(original) as before, ZipFile(package) as project, ZipFile(candidate, 'x') as after:
            for part in before.infolist():
                contents = project.read(VBA_PART) if part.filename == VBA_PART else before.read(part)
                after.writestr(part, contents)
        check_candidate(directory, path=candidate, transform=transform)
        link(candidate, destination)
    return destination


def parse(payload: str, *, pairs: int, platform: str, seconds: float) -> dict[str, Any]:
    """Validate the complete, locale-independent profile and its exclusive stages."""
    if not isinstance(payload, str) or len(payload) > 1024:
        raise ValueError('Expected one bounded diagnostic payload.')
    fields = payload.strip().split('|')
    if fields == [PROFILE, 'unavailable']:
        raise ValueError('The diagnostic clock or profile is unavailable.')
    if len(fields) != 4 + len(STAGES) + len(COUNTERS) or fields[0] != PROFILE:
        raise ValueError('The diagnostic profile has unexpected fields.')
    if fields[1] != 'complete':
        raise ValueError('Only a complete runner profile can count as benchmark evidence.')
    if fields[2] != CLOCKS.get(platform):
        raise ValueError('The diagnostic clock does not match the desktop platform.')
    until = 4 + len(STAGES)
    if not all(DECIMAL.fullmatch(value) for value in fields[3:until]):
        raise ValueError('The diagnostic batch timestamp and durations must be decimal numbers.')
    serial, *durations = map(float, fields[3:until])
    if not isfinite(serial) or serial <= 0 or any(not isfinite(value) or value < 0 for value in durations):
        raise ValueError('The diagnostic batch timestamp and durations must be finite and nonnegative.')
    if any(value > seconds + 0.01 for value in durations):
        raise ValueError('The diagnostic stages must fit inside the external macro wall time.')
    total = fsum(durations)
    if not 0 < total <= seconds + 0.01:
        raise ValueError('The diagnostic stage total must fit inside the external macro wall time.')
    if not all(INTEGER.fullmatch(value) for value in fields[until:]):
        raise ValueError('The diagnostic counts must be nonnegative integers.')
    computed, reused, stress, ranked, full, balance, repeated = map(int, fields[until:])
    if (computed + reused != pairs * 2 + STRESS_COUNT or reused > pairs + 13 or stress != STRESS_COUNT
            or ranked != 3 or full != computed + ranked or balance + repeated > full or repeated > computed):
        raise ValueError('The diagnostic counts do not reconcile to the completed fixture.')
    stages = dict(zip(STAGES, durations, strict=True))
    ordered = stages['ordered_calculation']
    optional_stages = ('ordered_calculation', 'reuse_verification')
    if any(not duration for stage, duration in stages.items() if stage not in optional_stages) or (
            (ordered > 0) != (balance + repeated > 0)):
        raise ValueError('The diagnostic stage durations do not match the work performed.')
    return {
        'clock': fields[2], 'batch_serial': serial, 'seconds': total,
        'stages': stages,
        'counters': dict(zip(COUNTERS, (computed, reused, stress, ranked, full, balance, repeated), strict=True)),
    }


def record(directory: Path, name: str, timing: dict[str, Any], payload: str) -> dict[str, Any]:
    """Build evidence only for a matched capture, candidate, reference and batch."""
    validate_run(timing)
    if timing['excel_bitness'] != 64:
        raise ValueError('The diagnostic clock requires 64-bit desktop Excel.')
    fixture = load_fixture(directory)
    profile = parse(payload, pairs=fixture['pairs'], platform=timing['platform'], seconds=timing['seconds'])
    if name == CANDIDATE:
        raise ValueError('Save a distinct capture of the diagnostic workbook.')
    candidate = check_candidate(directory)
    reference, capture = reference_file(directory, fixture), capture_file(directory, name)
    comparison, serial = compare_capture(directory, fixture, reference, capture, candidate=candidate,
                                         expected_reused=profile['counters']['reused'])
    if abs(serial - profile['batch_serial']) > 1e-9:
        raise ValueError('The diagnostic and saved capture must describe the same batch timestamp.')
    return {
        'schema': EVIDENCE, 'fixture': fixture['fixture'], 'pairs': fixture['pairs'],
        'fixture_sha256': digest(directory / 'fixture.json'), 'source_sha256': fixture['source_sha256'],
        'candidate_sha256': digest(candidate), 'reference_sha256': digest(reference),
        'capture': name, 'capture_sha256': digest(capture), 'run': timing, 'profile': profile,
        'payload': payload.strip(), 'comparison': comparison,
    }


def check(path: Path) -> dict[str, Any]:
    """Rebuild the claimed evidence and require an exact replay."""
    saved = read_json(path)
    if saved.get('schema') != EVIDENCE:
        raise ValueError('Expected stage profile evidence with the supported schema.')
    replay = record(path.parent, saved['capture'], saved['run'], saved['payload'])
    if replay != saved:
        raise ValueError('The recorded stage profile differs from its replay.')
    return replay


def check_all(paths: list[Path]) -> None:
    """Reject repeat labels and saved batches reused in the same invocation."""
    repeats, batches, captures = set(), set(), set()
    for path in paths:
        evidence = check(path)
        run = evidence['run']
        context = (evidence['source_sha256'], evidence['fixture'], evidence['pairs'], run['platform'], run['machine'])
        repeat = (*context, run['excel_build'], run['excel_channel'], run['excel_bitness'],
                  run['os_build'], run['architecture'], run['distribution'], run['macro_policy'],
                  run['phase'], run['repeat'])
        batch = (*context, evidence['profile']['batch_serial'])
        capture = evidence['capture_sha256']
        if repeat in repeats or batch in batches or capture in captures:
            raise ValueError('Duplicate diagnostic repeat, saved batch timestamp or reused capture.')
        repeats.add(repeat)
        batches.add(batch)
        captures.add(capture)
        print(f'{path.name}: {evidence["comparison"]}')


def main() -> int:
    """Prepare or replay a diagnostic workbook without using the baseline lane."""
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    build = commands.add_parser('prepare')
    build.add_argument('directory', type=Path)
    capture = commands.add_parser('record')
    capture.add_argument('directory', type=Path)
    capture.add_argument('workbook')
    capture.add_argument('timing', type=Path)
    capture.add_argument('profile', type=Path)
    capture.add_argument('--output', required=True)
    verify = commands.add_parser('check')
    verify.add_argument('evidence', type=Path, nargs='+')
    args = parser.parse_args()
    try:
        if args.command == 'prepare':
            print(prepare(args.directory))
        elif args.command == 'record':
            if not RECORD_NAME.fullmatch(args.output):
                raise ValueError('Use a profile-evidence-<id>.json filename.')
            with args.profile.open(encoding='ascii') as stream:
                payload = stream.read(1025)
            result = record(args.directory, args.workbook, read_json(args.timing), payload)
            write_json(args.directory / args.output, result)
            print(result['comparison'])
        else:
            check_all(args.evidence)
    except (AssertionError, OSError, ValueError, KeyError, BadZipFile, subprocess.CalledProcessError) as error:
        parser.exit(1, f'{error}\n')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
