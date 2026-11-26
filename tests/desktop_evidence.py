"""Record or replay a saved desktop Excel benchmark against its reference."""

from __future__ import annotations

import argparse
from contextlib import redirect_stdout
from io import StringIO
from math import isfinite
from pathlib import Path
import re
import subprocess
from typing import Any
from zipfile import BadZipFile

from pyopenvba import ExcelFile

from desktop_benchmark import INPUT, check_workbook, digest, load_fixture, read_json, write_json
from native_excel_acceptance import compare
from xlsx_read import Workbook


SCHEMA = 'DesktopEvidence/v1'
CHECKS = {'fresh_capture': 'pass', 'runner_source': 'pass', 'independent_reconciliation': 'pass'}
METADATA = ('excel_build', 'excel_channel', 'os_build', 'architecture', 'machine',
            'distribution', 'macro_policy', 'timing_method')
VBA_LITERAL = re.compile(r'("(?:[^"]|"")*")')
EXCEL_SHEET_BASE = re.compile(r'Attribute VB_Base = "0\{000208(?:19|20)-0000-0000-C000-000000000046\}"')
LIBRARY_ID = re.compile(r'\{[0-9a-f-]+\}#[^#]+', re.IGNORECASE)


def reference_file(directory: Path, fixture: dict[str, Any]) -> Path:
    """Select and verify the independent artifact for the declared case."""
    if fixture['fixture'] == 'example-95':
        # The 100-pair capture uses trial IDs shared with the 1,000-pair reference.
        # The existing comparator rebuilds the ranked ledgers for the smaller batch.
        reference = directory / 'reference.xlsm'
        if digest(reference) != fixture['source_sha256']:
            raise ValueError('The example reference must match the published snapshot.')
        check_workbook(reference, fixture, pairs=1000)
    else:
        reference = directory / 'reference.xlsx'
        contract = read_json(directory / 'reference.json')
        if contract != {'fixture_sha256': digest(directory / 'fixture.json'),
                        'reference_sha256': digest(reference)}:
            raise ValueError('The independent reference does not match this fixture.')
        check_workbook(reference, fixture, pairs=fixture['pairs'])
    return reference


def capture_file(directory: Path, name: str) -> Path:
    """Accept only a separate, local saved workbook in this bundle."""
    if (not isinstance(name, str) or not re.fullmatch(r'[A-Za-z0-9_-]+\.xlsm', name)
            or name in (INPUT, 'reference.xlsm')):
        raise ValueError('Use a distinct capture .xlsm filename inside the fixture directory.')
    capture = directory / name
    if capture.is_symlink() or not capture.is_file():
        raise ValueError('The capture must be an existing regular workbook in the fixture directory.')
    return capture


def validate_run(run: dict[str, Any]) -> None:
    """Check the operator-supplied platform and single-run timing fields."""
    if (not isinstance(run, dict) or run.get('platform') not in ('mac', 'windows')
            or run.get('phase') not in ('cold', 'warm')
            or type(run.get('repeat')) is not int or run['repeat'] < 1
            or (run['phase'] == 'cold' and run['repeat'] != 1)
            or type(run.get('excel_bitness')) is not int or run['excel_bitness'] not in (32, 64)
            or not all(isinstance(run.get(field), str) and run[field].strip() for field in METADATA)):
        raise ValueError('Run metadata must identify the Excel environment, policy and repeat.')
    seconds = run.get('seconds')
    if type(seconds) not in (float, int) or not isfinite(seconds) or seconds <= 0:
        raise ValueError('Record a positive, finite elapsed time in seconds.')


def vba_identity(path: Path) -> tuple[dict[str, tuple[str, ...]], tuple[tuple[str, str], ...]]:
    """Read executable source and stable library IDs across native Excel saves."""
    with ExcelFile(path) as workbook:
        if workbook.forms():
            raise ValueError('The benchmark runner must not contain VBA forms.')
        libraries = []
        for reference in workbook.references():
            library_id = LIBRARY_ID.search(reference.libid)
            if library_id is None:
                raise ValueError('The benchmark runner has an unsupported VBA library reference.')
            libraries.append((reference.name.casefold(), library_id[0].casefold()))
        modules = {}
        for name, source in workbook.vba_modules().items():
            lines = source.splitlines()
            if (re.fullmatch(r'Sheet[0-9]+', name)
                    and all(not line.strip() or line.startswith('Attribute VB_') for line in lines)):
                # Desktop Excel can substitute the default host class for an empty sheet module.
                lines = [line for line in lines if not EXCEL_SHEET_BASE.fullmatch(line)]
            # Excel can recase identifiers on save. String literals retain their exact text.
            modules[name] = tuple(''.join(part if index % 2 else part.casefold()
                                          for index, part in enumerate(VBA_LITERAL.split(line))) for line in lines)
    return modules, tuple(sorted(libraries))


def compare_capture(directory: Path, fixture: dict[str, Any], reference: Path, capture: Path, *,
                    candidate: Path | None = None, expected_reused: int | None = None) -> tuple[str, float]:
    """Check freshness and run the existing field-by-field reconciliation."""
    check_workbook(capture, fixture, pairs=fixture['pairs'])
    with Workbook(directory / INPUT) as before, Workbook(capture) as after:
        old_stamp, new_stamp = (book['Risk_Data']['Z7'].value for book in (before, after))
        if (not all(type(stamp) in (int, float) and isfinite(stamp) for stamp in (old_stamp, new_stamp))
                or new_stamp <= old_stamp):
            raise ValueError('The capture must contain a newly completed Excel batch.')
    if vba_identity(capture) != vba_identity(candidate if candidate is not None else directory / INPUT):
        raise ValueError('The capture VBA source or libraries differ from the pinned runner.')
    output = StringIO()
    with redirect_stdout(output):
        passed = compare(reference, capture, expected_reused)
    if not passed:
        raise ValueError('Independent reconciliation failed.')
    return output.getvalue().strip(), new_stamp


def evidence(directory: Path, name: str, run: dict[str, Any]) -> dict[str, Any]:
    """Assemble a record only after real workbook evidence passes."""
    validate_run(run)
    fixture = load_fixture(directory)
    reference, capture = reference_file(directory, fixture), capture_file(directory, name)
    comparison, batch_serial = compare_capture(directory, fixture, reference, capture)
    return {
        'schema': SCHEMA, 'fixture_sha256': digest(directory / 'fixture.json'),
        'input_sha256': fixture['input_sha256'], 'source_sha256': fixture['source_sha256'],
        'fixture': fixture['fixture'], 'pairs': fixture['pairs'], 'seed': fixture['seed'], 'outcomes': fixture['outcomes'],
        'reference_sha256': digest(reference), 'capture': name, 'capture_sha256': digest(capture),
        'run': run, 'checks': CHECKS, 'comparison': comparison, 'batch_serial': batch_serial,
    }


def check(path: Path) -> tuple[tuple[Any, ...], tuple[Any, ...], str]:
    """Replay a complete, correctly bound record and return its repeat identity."""
    record = read_json(path)
    if record.get('schema') != SCHEMA or record.get('checks') != CHECKS:
        raise ValueError('A complete DesktopEvidence/v1 record with all passing checks is required.')
    run = record.get('run', {})
    validate_run(run)
    directory = path.parent
    fixture = load_fixture(directory)
    reference = reference_file(directory, fixture)
    capture = capture_file(directory, record.get('capture', ''))
    expected = {field: fixture[field] for field in ('fixture', 'pairs', 'seed', 'outcomes',
                                                    'input_sha256', 'source_sha256')}
    expected.update(fixture_sha256=digest(directory / 'fixture.json'),
                    reference_sha256=digest(reference), capture_sha256=digest(capture))
    if any(record.get(field) != value for field, value in expected.items()):
        raise ValueError('The evidence does not match the fixture or artifact hashes.')
    observed, batch_serial = compare_capture(directory, fixture, reference, capture)
    if record.get('comparison') != observed:
        raise ValueError('The saved reconciliation text differs from the replay.')
    if record.get('batch_serial') != batch_serial:
        raise ValueError('The saved batch timestamp differs from the workbook.')
    print(f'{path.name}: {observed}')
    candidate = (fixture['source_sha256'], fixture['fixture'], fixture['pairs'], run['platform'], run['machine'])
    repeat = (
        *candidate, run['excel_build'], run['excel_channel'], run['excel_bitness'], run['os_build'],
        run['architecture'], run['distribution'], run['macro_policy'], run['phase'], run['repeat'],
    )
    return repeat, (*candidate, batch_serial), expected['capture_sha256']


def main() -> int:
    """Record a run or check one or more existing records."""
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    record = commands.add_parser('record')
    record.add_argument('directory', type=Path)
    record.add_argument('capture', help='saved .xlsm filename in the fixture directory')
    record.add_argument('timing', type=Path, help='JSON file with operator timing and environment fields')
    record.add_argument('--output', required=True, help='new evidence JSON filename in the fixture directory')
    verify = commands.add_parser('check')
    verify.add_argument('evidence', nargs='+', type=Path)
    args = parser.parse_args()
    try:
        if args.command == 'record':
            if not re.fullmatch(r'evidence-[A-Za-z0-9_-]+\.json', args.output):
                raise ValueError('Use an evidence-<id>.json filename.')
            path = args.directory / args.output
            result = evidence(args.directory, args.capture, read_json(args.timing))
            write_json(path, result)
            print(f'{path}: {result["comparison"]}')
        else:
            seen, batches, captures = set(), set(), set()
            for path in args.evidence:
                key, batch, capture_hash = check(path)
                if key in seen or batch in batches or capture_hash in captures:
                    raise ValueError('Duplicate benchmark repeat, saved batch timestamp or reused capture.')
                seen.add(key)
                batches.add(batch)
                captures.add(capture_hash)
    except (AssertionError, OSError, ValueError, KeyError, BadZipFile, subprocess.CalledProcessError) as error:
        parser.exit(1, f'{error}\n')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
