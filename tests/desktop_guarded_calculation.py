"""Prepare and reconcile a disposable smart-first Mac calculation experiment."""

from __future__ import annotations

import argparse
from functools import partial
from pathlib import Path
import re
from statistics import median
from typing import Any

from desktop_benchmark import PAIR_COUNTS, digest, load_fixture, read_json, write_json
from desktop_evidence import capture_file, compare_capture, reference_file, validate_run
from desktop_stage_profile import CANDIDATE, PROFILE, check_candidate, instrument, parse
from desktop_stage_profile import check as check_control
from desktop_stage_profile import prepare as prepare_profile
from vba_declarations import audit_vba_declarations


SCHEMA = 'GuardedCalculationEvidence/v1'
COMPARISON = 'GuardedMacExperiment/v1'
WARM_COMPARISON = 'GuardedMacWarmExperiment/v1'
GUARDED_PROFILE = 'RunnerGuardedProfile/v1'
MODES = ('natural', 'canary')
ENVIRONMENT = ('platform', 'excel_build', 'excel_channel', 'excel_bitness', 'os_build',
               'architecture', 'machine', 'macro_policy')
PROVENANCE = ('fixture_sha256', 'source_sha256', 'reference_sha256')
WARM_SEQUENCE = (('cold', 1), ('warm', 1), ('warm', 2), ('warm', 3))
MACRO_WALL_TIME = ('monotonic time from synchronous RunMarketRisk dispatch through return, '
                   'excluding save and profile read')
MACRO_WALL_TIME_LABELS = (
    MACRO_WALL_TIME,
    'monotonic wall time from synchronous RunMarketRisk dispatch through return',
)
FULL_BRANCH = (
    '        If attempt = 1 Then\n'
    '            MoveMarketRiskProfile ProfilePrimaryCalculation\n'
    '            mProfileFull = mProfileFull + 1\n'
    '            FinishTrialCalculation\n'
    '        Else\n'
)
GUARDED_BRANCH = (
    '        If attempt = 1 Then\n'
    '            MoveMarketRiskProfile ProfilePrimaryCalculation\n'
    '            mProfileFull = mProfileFull + 1\n'
    '            FinishTrialCalculation True\n'
    '        ElseIf attempt = 2 Then\n'
    '            MoveMarketRiskProfile ProfilePrimaryCalculation\n'
    '            mGuardedFull = mGuardedFull + 1\n'
    '            FinishTrialCalculation\n'
    '        Else\n'
)
COUNTS = '''
Public Function ReadGuardedCalculationCounts() As String
    If mRunning Or mProfileOutcome <> "complete" Then Exit Function
    ReadGuardedCalculationCounts = "GuardedCalculation/v1|" & Trim$(Str$(mProfileSerial)) _
        & "|" & Trim$(Str$(mGuardedFull)) & "|" & Trim$(Str$(mGuardedCanary))
End Function
'''


def replace(source: str, old: str, new: str) -> str:
    """Fail before writing a workbook if a runner anchor changes."""
    if source.count(old) != 1:
        raise ValueError(f'Expected one guarded calculation anchor for {old.strip().splitlines()[0]!r}.')
    return source.replace(old, new, 1)


def overlay(source: str, *, mode: str) -> str:
    """Keep all existing verification around a bounded three-attempt scratch runner."""
    if mode not in MODES:
        raise ValueError('Unknown guarded calculation mode.')
    result = instrument(source)
    result = replace(result, '    For attempt = 1 To 2\n', '    For attempt = 1 To 3\n')
    result = replace(result, FULL_BRANCH, GUARDED_BRANCH)
    result = replace(result, '            If attempt = 2 Then Exit For\n', '            If attempt = 3 Then Exit For\n')
    result = replace(result, '    If attempt > 2 Then RaiseRunError ', '    If attempt > 3 Then RaiseRunError ')
    result = replace(result, 'Private Sub FinishTrialCalculation()\n',
                     'Private Sub FinishTrialCalculation(Optional ByVal smart As Boolean = False)\n')
    result = replace(result, '    Application.CalculateFull\n',
                     '    If smart Then\n        Application.Calculate\n    Else\n'
                     '        Application.CalculateFull\n    End If\n')
    result = replace(result, 'Private mProfileFull As Long\n',
                     'Private mProfileFull As Long\nPrivate mGuardedFull As Long\n'
                     'Private mGuardedCanary As Long\n')
    result = replace(result, '    mProfileFull = 0\n',
                     '    mProfileFull = 0\n    mGuardedFull = 0\n    mGuardedCanary = 0\n')
    result = replace(result, f'"{PROFILE}"', f'"{GUARDED_PROFILE}"')
    if mode == 'canary':
        anchor = '        values = source.Value2\n'
        result = replace(result, anchor, anchor +
                         '        If primary And code = 1 And attempt <= 2 Then\n'
                         '            values(1, FIRST_YEAR_FIELD + ANNUAL_CAPACITY) = _\n'
                         '                values(1, FIRST_YEAR_FIELD + ANNUAL_CAPACITY) + 1#\n'
                         '            mGuardedCanary = mGuardedCanary + 1\n'
                         '        End If\n')
    result = result.replace('mProfileFull', 'mGuardedFirst')
    result += '\n' + COUNTS
    audit_vba_declarations(result)
    return result


def matched_fixture(directory: Path) -> dict[str, Any]:
    """Limit this experiment to the pinned full-horizon catalog fixtures."""
    fixture = load_fixture(directory)
    if fixture['fixture'] != 'full-0-125':
        raise ValueError('Use a full-0-125 diagnostic fixture.')
    return fixture


def read_profile(profile_text: str, count_text: str, *, seconds: float, mode: str,
                 pairs: int = 100) -> dict[str, Any]:
    """Check exclusive timings and reconcile smart, full, ordered and injected attempts."""
    if mode not in MODES:
        raise ValueError('Unknown guarded calculation mode.')
    tag, separator, content = profile_text.strip().partition('|')
    if not separator or tag != GUARDED_PROFILE:
        raise ValueError('Expected a guarded calculation stage profile.')
    # The first-attempt counter occupies the same position as the baseline full-attempt counter.
    profile = parse(PROFILE + '|' + content, pairs=pairs, platform='mac', seconds=seconds)
    counts = profile['counters']
    counts['smart_attempts'] = counts.pop('full_attempts')
    fields = count_text.strip().split('|')
    if (len(fields) != 4 or fields[0] != 'GuardedCalculation/v1'
            or not all(re.fullmatch(r'[0-9]{1,4}', value) for value in fields[2:])):
        raise ValueError('Expected guarded fallback counts.')
    if fields[1] != content.split('|', 3)[2]:
        raise ValueError('The fallback counts came from a different batch.')
    full, injected = map(int, fields[2:])
    ordered = counts['retry_balance'] + counts['retry_repeated_ending']
    if not 0 <= ordered <= full <= counts['smart_attempts']:
        raise ValueError('Fallback counts do not fit the bounded attempt ladder.')
    if (mode == 'natural' and injected != 0 or mode == 'canary'
            and (injected != 2 or full < 1 or counts['retry_balance'] < 1)):
        raise ValueError('The injected guard exercise did not follow the expected recovery path.')
    counts['full_fallbacks'] = full
    counts['injected_capture_errors'] = injected
    return profile


def prepare(directory: Path, mode: str) -> Path:
    """Build the exact guarded fixture without changing source artifacts."""
    matched_fixture(directory)
    return prepare_profile(directory, transform=partial(overlay, mode=mode))


def record(directory: Path, name: str, timing: dict[str, Any], profile_text: str,
           count_text: str, mode: str) -> dict[str, Any]:
    """Reconcile a fresh diagnostic capture and bind it to its measurements."""
    validate_run(timing)
    if timing['platform'] != 'mac' or timing['excel_bitness'] != 64:
        raise ValueError('This experiment requires 64-bit Excel for Mac.')
    fixture = matched_fixture(directory)
    if name == CANDIDATE:
        raise ValueError('Keep the diagnostic input and the saved capture separate.')
    profile = read_profile(profile_text, count_text, seconds=timing['seconds'], mode=mode, pairs=fixture['pairs'])
    candidate = check_candidate(directory, transform=partial(overlay, mode=mode))
    reference, capture = reference_file(directory, fixture), capture_file(directory, name)
    comparison, serial = compare_capture(directory, fixture, reference, capture, candidate=candidate,
                                         expected_reused=profile['counters']['reused'])
    if abs(serial - profile['batch_serial']) > 1e-9:
        raise ValueError('The diagnostic and saved capture describe different batches.')
    return {
        'schema': SCHEMA, 'mode': mode, 'fixture_sha256': digest(directory / 'fixture.json'),
        'candidate_sha256': digest(candidate), 'capture': name, 'capture_sha256': digest(capture),
        'source_sha256': fixture['source_sha256'], 'reference_sha256': digest(reference),
        'run': timing, 'profile': profile, 'profile_text': profile_text.strip(),
        'count_text': count_text.strip(), 'comparison': comparison,
    }


def check(path: Path) -> dict[str, Any]:
    """Require the same evidence when replayed against the source and saved workbook."""
    saved = read_json(path)
    replay = record(path.parent, saved['capture'], saved['run'], saved['profile_text'],
                    saved['count_text'], saved['mode'])
    if replay != saved:
        raise ValueError('The recorded guarded calculation differs from its replay.')
    return replay


def require_matched_runs(all_runs: list[dict], environment: tuple[str, ...] = ENVIRONMENT) -> None:
    """Require independent saved batches from a common fixture and environment."""
    for field in PROVENANCE:
        if len({item[field] for item in all_runs}) != 1:
            raise ValueError(f'The test arms used different {field}.')
    for field in environment:
        if len({item['run'][field] for item in all_runs}) != 1:
            raise ValueError(f'The test arms used different {field}.')
    if any(item['run']['timing_method'] not in MACRO_WALL_TIME_LABELS for item in all_runs):
        raise ValueError('The comparison requires monotonic macro wall time through return.')
    if (len({item['capture_sha256'] for item in all_runs}) != len(all_runs)
            or len({item['profile']['batch_serial'] for item in all_runs}) != len(all_runs)):
        raise ValueError('The test arms reused a saved capture or batch timestamp.')


def matched_results(controls: list[dict], candidates: list[dict], canary: dict) -> dict:
    """Compare the initial cold and warm pairs after replaying their evidence."""
    if len(controls) != 2 or len(candidates) != 2:
        raise ValueError('Provide one cold run and one warm run for each comparison arm.')
    fixture, pairs = controls[0]['fixture'], controls[0]['pairs']
    if (fixture != 'full-0-125' or pairs not in PAIR_COUNTS
            or (controls[1]['fixture'], controls[1]['pairs']) != (fixture, pairs)):
        raise ValueError('The controls must use the same full-horizon fixture and pair count.')
    if (any(item.get('mode') != 'natural' for item in candidates) or canary.get('mode') != 'canary'
            or (canary['run']['phase'], canary['run']['repeat']) != ('cold', 1)):
        raise ValueError('The comparison requires natural timings and a separate cold canary.')
    require_matched_runs([*controls, *candidates, canary])
    for arm in (controls, candidates):
        cold, warm = arm
        if [(item['run']['phase'], item['run']['repeat']) for item in arm] != [('cold', 1), ('warm', 1)]:
            raise ValueError('Pass the cold result before its warm repeat for each arm.')
        if cold['profile']['batch_serial'] >= warm['profile']['batch_serial']:
            raise ValueError('The warm run must have its own newer batch.')
        if cold['candidate_sha256'] != warm['candidate_sha256']:
            raise ValueError('Each arm must reuse its unchanged input template.')

    def compact(item: dict) -> dict:
        return {key: item[key] for key in ('candidate_sha256', 'capture_sha256', 'profile', 'comparison')}

    return {
        'schema': COMPARISON, 'experiment': f'{fixture}, {pairs} fictional pairs',
        'environment': {**{field: canary['run'][field] for field in ENVIRONMENT}, 'timing_method': MACRO_WALL_TIME},
        **{field: canary[field] for field in PROVENANCE},
        'matched_initial_observations': [
            {'phase': control['run']['phase'], 'control_seconds': control['run']['seconds'],
             'guarded_seconds': candidate['run']['seconds'], 'control': compact(control),
             'guarded': compact(candidate)}
            for control, candidate in zip(controls, candidates, strict=True)
        ],
        'ordered_recovery_canary': {**compact(canary), 'seconds': canary['run']['seconds']},
        'windows_status': 'unverified',
    }


def warm_session(records: list[tuple[dict, dict]], *, guarded: bool) -> list[dict]:
    """Bind each native timing to the unchanged open Excel session and its saved predecessor."""
    if len(records) != len(WARM_SEQUENCE):
        raise ValueError('Provide a cold run and three warm repeats for each arm.')
    first, first_receipt = records[0]
    expected_pid = first_receipt.get('excel_pid')
    if not isinstance(expected_pid, str) or not re.fullmatch(r'[1-9][0-9]*', expected_pid):
        raise ValueError('The timing receipt needs an Excel process ID.')
    candidate_hash = preceding_hash = first['candidate_sha256']
    preceding_batch = 0.0
    result = []
    for (evidence, receipt), (phase, repeat) in zip(records, WARM_SEQUENCE, strict=True):
        run = evidence['run']
        validate_run(run)
        if (run['phase'], run['repeat']) != (phase, repeat):
            raise ValueError('File order must be cold, then warm repeats 1, 2 and 3.')
        if run['platform'] != 'mac' or run['excel_bitness'] != 64:
            raise ValueError('The matched warm experiment requires 64-bit Excel for Mac.')
        if evidence.get('mode') != ('natural' if guarded else None):
            raise ValueError('The warm arms require a full-first control and natural smart-first candidate.')
        if receipt.get('excel_pid') != expected_pid:
            raise ValueError('A warm arm must stay in the same Excel process as its cold run.')
        if receipt.get('starting_sha256') != preceding_hash:
            raise ValueError('A warm run must follow its preceding saved capture.')
        profile = evidence['profile_text' if guarded else 'payload']
        counts = evidence['count_text'] if guarded else ''
        if (receipt.get('timing') != run or receipt.get('profile') != profile
                or receipt.get('counts') != counts):
            raise ValueError('The timing receipt and independently replayed evidence disagree.')
        batch = evidence['profile']['batch_serial']
        if evidence['candidate_sha256'] != candidate_hash or batch <= preceding_batch:
            raise ValueError('The unchanged arm must publish a distinct newer batch for each run.')
        result.append(evidence)
        preceding_hash, preceding_batch = evidence['capture_sha256'], batch
    return result


def matched_warm_results(controls: list[tuple[dict, dict]], candidates: list[tuple[dict, dict]]) -> dict:
    """Report warm medians only after all matched capture and timing replays."""
    control = warm_session(controls, guarded=False)
    guarded = warm_session(candidates, guarded=True)
    fixture, pairs = control[0]['fixture'], control[0]['pairs']
    if (fixture != 'full-0-125' or pairs not in PAIR_COUNTS
            or any((item['fixture'], item['pairs']) != (fixture, pairs) for item in control)):
        raise ValueError('Use matched full-0-125 controls at a single allowed pair count.')
    if controls[0][1]['excel_pid'] == candidates[0][1]['excel_pid']:
        raise ValueError('Each cold arm requires its own fresh Excel process.')
    require_matched_runs([*control, *guarded], (*ENVIRONMENT, 'distribution'))
    data = {}
    medians = {}
    for label, arm in (('control', control), ('guarded', guarded)):
        data[label] = [
            {key: item[key] for key in ('run', 'candidate_sha256', 'capture_sha256', 'profile', 'comparison')}
            for item in arm
        ]
        medians[label] = median(item['run']['seconds'] for item in arm[1:])
    return {
        'schema': WARM_COMPARISON, 'experiment': f'{fixture}, {pairs} fictional pairs',
        **{field: control[0][field] for field in PROVENANCE},
        'arms': data, 'warm_median_seconds': medians,
        'warm_improvement_percent': (medians['control'] - medians['guarded']) / medians['control'] * 100,
        'windows_status': 'unverified',
    }


def replay_warm_arm(paths: list[Path], *, guarded: bool) -> list[tuple[dict, dict]]:
    """Replay each workbook capture and load its corresponding native timing receipt."""
    replay = check if guarded else check_control
    return [
        (replay(path), read_json(path.with_name(f'raw-{phase}-{repeat:02}.json')))
        for path, (phase, repeat) in zip(paths, WARM_SEQUENCE, strict=True)
    ]


def main() -> int:
    """Prepare a local fixture, or record or replay a saved result."""
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    build = commands.add_parser('prepare')
    write = commands.add_parser('record')
    for command in (build, write):
        command.add_argument('directory', type=Path)
        command.add_argument('--mode', choices=MODES, default='natural')
    write.add_argument('capture')
    for name in ('timing', 'profile', 'counts', 'output'):
        write.add_argument(name, type=Path)
    verify = commands.add_parser('check')
    verify.add_argument('evidence', type=Path)
    compare = commands.add_parser('compare')
    for option in ('control', 'candidate'):
        compare.add_argument(f'--{option}', type=Path, nargs=2, required=True,
                             metavar=('COLD', 'WARM'))
    compare.add_argument('--canary', type=Path, required=True)
    compare.add_argument('--output', type=Path, required=True)
    warm = commands.add_parser('compare-warm')
    for option in ('control', 'candidate'):
        warm.add_argument(f'--{option}', type=Path, nargs=4, required=True,
                          metavar=('COLD', 'WARM1', 'WARM2', 'WARM3'))
    warm.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if args.command == 'prepare':
        print(prepare(args.directory, args.mode))
    elif args.command == 'record':
        result = record(args.directory, args.capture, read_json(args.timing), args.profile.read_text(),
                        args.counts.read_text(), args.mode)
        write_json(args.output, result)
        print(result['comparison'])
    elif args.command == 'check':
        print(check(args.evidence)['comparison'])
    elif args.command == 'compare':
        write_json(args.output, matched_results([check_control(path) for path in args.control],
                                                [check(path) for path in args.candidate], check(args.canary)))
    else:
        write_json(args.output, matched_warm_results(replay_warm_arm(args.control, guarded=False),
                                                     replay_warm_arm(args.candidate, guarded=True)))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
