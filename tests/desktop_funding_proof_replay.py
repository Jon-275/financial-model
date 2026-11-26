"""Replay matched full-scale Mac funding-proof runs from two local evidence bundles."""

from argparse import ArgumentParser
import json
from pathlib import Path
from statistics import median
from subprocess import CalledProcessError
from typing import Any
from zipfile import BadZipFile, ZipFile

from pyopenvba import ExcelFile

from desktop_benchmark import INPUT, SEED, digest, load_fixture, read_json, write_json
from desktop_evidence import capture_file, compare_capture, reference_file, validate_run, vba_identity
from desktop_guarded_calculation import (
    ENVIRONMENT, MACRO_WALL_TIME, PROVENANCE, WARM_SEQUENCE, overlay as guarded_overlay,
    read_profile, replay_warm_arm, require_matched_runs, warm_session,
)
from desktop_stage_profile import RUNNER, VBA_PART
from native_funding_proof_overlay import instrument, parse_counts
from native_funding_proof_package import verify_package


SCHEMA = 'NativeFundingProofMacWarmComparison/v1'
FIXTURE = 'full-0-125'
PAIRS = 1000
BANK = 'proof-bank.xlsm'
CANDIDATE = 'proof-runner-input.xlsm'
SAVE_METHODS = ('standard Excel save', 'scoped alert setting, restored')
LABELS = tuple(f'{phase}-{repeat:02}' for phase, repeat in WARM_SEQUENCE)


def full_scale_fixture(directory: Path) -> dict[str, Any]:
    """Load a pinned catalogue input and restrict the comparison to the full-scale case."""
    fixture = load_fixture(directory)
    if (fixture['fixture'], fixture['pairs'], fixture['seed']) != (FIXTURE, PAIRS, SEED):
        raise ValueError('Expected the fictional full-0-125, 1,000-pair fixture and seed.')
    return fixture


def verify_candidate(directory: Path) -> Path:
    """Require the generated worksheet bank and exactly the integrated natural runner."""
    bank = verify_package(directory / INPUT, capture_file(directory, BANK))
    candidate = capture_file(directory, CANDIDATE)
    with ZipFile(bank) as before, ZipFile(candidate) as after:
        parts = before.namelist()
        if (VBA_PART not in parts or parts != after.namelist() or before.comment != after.comment
                or [name for name in parts if before.read(name) != after.read(name)] != [VBA_PART]):
            raise ValueError('The candidate must change only the verified bank’s VBA project.')
    before_modules, before_libraries = vba_identity(bank)
    after_modules, after_libraries = vba_identity(candidate)
    if (RUNNER not in before_modules or RUNNER not in after_modules or before_libraries != after_libraries
            or {k: v for k, v in before_modules.items() if k != RUNNER}
            != {k: v for k, v in after_modules.items() if k != RUNNER}):
        raise ValueError('The proof candidate changed an unrelated VBA module or library.')
    with ExcelFile(bank) as before, ExcelFile(candidate) as after:
        expected = instrument(guarded_overlay(before.get_module(RUNNER), mode='natural'))
        if after.get_module(RUNNER).splitlines() != expected.splitlines():
            raise ValueError('The proof candidate differs from the reviewed natural VBA overlay.')
    return candidate


def replay_capture(directory: Path, capture: Path, receipt: dict[str, Any], *, fixture: dict[str, Any],
                   candidate: Path, reference: Path) -> dict[str, Any]:
    """Bind the raw native receipts and the independently reconciled saved batch."""
    timing, profile_text, count_text, proof_text = (receipt[key] for key in ('timing', 'profile', 'counts', 'proof'))
    validate_run(timing)
    if (timing['platform'], timing['excel_bitness'], timing['timing_method']) != ('mac', 64, MACRO_WALL_TIME):
        raise ValueError('Expected the matched 64-bit Mac macro wall time.')
    profile = read_profile(profile_text, count_text, pairs=PAIRS, seconds=timing['seconds'], mode='natural')
    proof = parse_counts(proof_text)
    counts = profile['counters']
    retries = counts['full_fallbacks'] + counts['retry_balance'] + counts['retry_repeated_ending']
    if proof['attempts'] != counts['smart_attempts'] + retries or proof['rejected_attempts'] > retries:
        raise ValueError('The funding proof counts do not reconcile to the verified runner attempts.')
    comparison, serial = compare_capture(directory, fixture, reference, capture, candidate=candidate,
                                         expected_reused=counts['reused'])
    if abs(serial - profile['batch_serial']) > 1e-9:
        raise ValueError('The proof receipt and saved batch identity disagree.')
    return {
        'arm': 'smart-first with independent funded draw witness', 'run': timing,
        'proof': proof, 'proof_payload': proof_text, 'profile': profile, 'comparison': comparison,
        'capture': capture.name, 'capture_sha256': digest(capture), 'candidate_sha256': digest(candidate),
        'fixture_sha256': digest(directory / 'fixture.json'), 'reference_sha256': digest(reference),
        'source_sha256': fixture['source_sha256'], 'windows_status': 'unverified',
    }


def check_api_receipts(directory: Path, records: list[dict[str, Any]]) -> dict[str, Any]:
    """Bind API open/save actions to the verified bytes independently of timing metadata."""
    opening = read_json(directory / 'api-open.json')
    saves = [read_json(directory / f'save-{label}.json') for label in LABELS]
    if (opening.get('input_sha256') != records[0]['candidate_sha256']
            or opening.get('automation_policy_restored') is not True
            or type(opening.get('temporary_automation_policy')) is not bool):
        raise ValueError('The API open receipt does not match the verified candidate or its restored policy.')
    for saved, record in zip(saves, records, strict=True):
        if saved.get('capture_sha256') != record['capture_sha256'] or saved.get('method') not in SAVE_METHODS:
            raise ValueError('An API save receipt does not match its verified Excel capture.')
    return {'open': opening, 'saves': saves}


def replay(control_dir: Path, candidate_dir: Path) -> dict[str, Any]:
    """Replay both complete warm arms without modifying the supplied evidence bundles."""
    full_scale_fixture(control_dir)
    fixture = full_scale_fixture(candidate_dir)
    candidate = verify_candidate(candidate_dir)
    reference = reference_file(candidate_dir, fixture)
    control_records = replay_warm_arm([control_dir / f'profile-evidence-{label}.json' for label in LABELS],
                                     guarded=False)
    controls = warm_session(control_records, guarded=False)
    if any((item['fixture'], item['pairs']) != (FIXTURE, PAIRS) for item in controls):
        raise ValueError('Expected the verified full-scale control.')

    candidates, session_records = [], []
    for label in LABELS:
        claim = read_json(candidate_dir / f'evidence-{label}.json')
        receipt = read_json(candidate_dir / f'raw-{label}.json')
        capture = capture_file(candidate_dir, f'capture-{label}.xlsm')
        record = replay_capture(candidate_dir, capture, receipt, fixture=fixture,
                                candidate=candidate, reference=reference)
        if record != claim:
            raise ValueError('The native proof evidence differs from its independent replay.')
        candidates.append(record)
        session_records.append(({**record, 'mode': 'natural', 'profile_text': receipt['profile'].strip(),
                                 'count_text': receipt['counts'].strip()}, receipt))
    warm_session(session_records, guarded=True)
    if control_records[0][1]['excel_pid'] == session_records[0][1]['excel_pid']:
        raise ValueError('Each arm must begin in its own Excel process.')
    require_matched_runs([*controls, *candidates], (*ENVIRONMENT, 'distribution'))
    api = check_api_receipts(candidate_dir, candidates)

    medians = {label: median(record['run']['seconds'] for record in arm[1:])
               for label, arm in (('full_first', controls), ('funding_proof', candidates))}
    return {
        'schema': SCHEMA, 'fixture': FIXTURE, 'pairs': PAIRS, 'seed': SEED,
        **{field: controls[0][field] for field in PROVENANCE},
        'run_environment': {field: controls[0]['run'][field] for field in ENVIRONMENT},
        'warm_median_seconds': medians,
        'warm_improvement_percent': (medians['full_first'] - medians['funding_proof']) / medians['full_first'] * 100,
        'arms': {'full_first': controls, 'funding_proof': candidates},
        'candidate_excel_api': api, 'windows_status': 'unverified',
    }


def main() -> int:
    """Read two evidence directory paths and print or write the independently replayed report."""
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('control', type=Path, help='full-first control evidence directory')
    parser.add_argument('candidate', type=Path, help='natural funding-proof evidence directory')
    parser.add_argument('--output', type=Path, help='write a new JSON report instead of using stdout')
    options = parser.parse_args()
    try:
        report = replay(options.control, options.candidate)
        if options.output is None:
            print(json.dumps(report, indent=2, sort_keys=True, allow_nan=False))
        else:
            write_json(options.output, report)
    except (OSError, ValueError, BadZipFile, CalledProcessError, AttributeError, KeyError, TypeError) as error:
        parser.error(str(error))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
