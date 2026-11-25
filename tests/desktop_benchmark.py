"""Prepare pinned fictional desktop fixtures and independently built references."""

from __future__ import annotations

import argparse
from hashlib import file_digest, sha256
from io import BytesIO
import json
from pathlib import Path
import re
from shutil import copyfile
import subprocess
import sys
from tempfile import TemporaryDirectory
from typing import Any
from zipfile import BadZipFile, ZipFile

from xlsx_read import Workbook


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))
from xlsx_values import write_values


FIXTURES = {
    'example-95': (35, 95, {}),
    'extended-125': (35, 125, {'D53': 125}),
    'full-0-125': (0, 125, {'D8': 0, 'D9': 0, 'D53': 125}),
}
PAIR_COUNTS = (100, 1000)
SEED = 2026
STRESS_COUNT = 33
FIXTURE_SCHEMA = 'DesktopFixture/v1'
SOURCE = 'workbook/retirement-model.xlsm'
INPUT = 'desktop-input.xlsm'


def digest(path: Path) -> str:
    """Hash exact artifact bytes."""
    with path.open('rb') as stream:
        return file_digest(stream, 'sha256').hexdigest()


def git(*args: str) -> bytes:
    """Read repository objects without changing the checkout."""
    return subprocess.check_output(['git', *args], cwd=ROOT)


def read_json(path: Path) -> dict[str, Any]:
    """Read a manifest object."""
    value = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(value, dict):
        raise ValueError(f'{path.name} must contain a JSON object.')
    return value


def write_json(path: Path, value: dict[str, Any]) -> None:
    """Write a new manifest without overwriting an existing record."""
    with path.open('x', encoding='utf-8') as stream:
        stream.write(json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + '\n')


def overrides(case: str, pairs: int) -> dict[str, dict[str, int]]:
    """Derive the only allowed edits for this fixture."""
    if case not in FIXTURES or type(pairs) is not int or pairs not in PAIR_COUNTS:
        raise ValueError('Choose a catalog fixture and 100 or 1,000 pairs.')
    return {'Inputs': dict(FIXTURES[case][2]), 'Risk_Settings': {'B29': pairs, 'B30': SEED}}


def check_source(commit: str, expected: str) -> None:
    """Bind a release checksum to the fictional workbook tracked at the commit."""
    if not re.fullmatch(r'[0-9a-f]{40}', commit) or not re.fullmatch(r'[0-9a-f]{64}', expected):
        raise ValueError('Expected a full commit SHA and lowercase SHA-256.')
    if sha256(git('show', f'{commit}:{SOURCE}')).hexdigest() != expected:
        raise ValueError('The workbook checksum differs from the pinned source commit.')


def check_workbook(path: Path, fixture: dict[str, Any], *, pairs: int) -> None:
    """Require the exact benchmark settings and the workbook risk version."""
    opening, horizon, values = FIXTURES[fixture['fixture']]
    required = {
        'Inputs': {'B4': 1, 'D8': opening, 'D53': horizon, **values},
        'Risk_Settings': {'B29': pairs, 'B30': SEED},
        'Risk_Data': {'Z2': fixture['risk_schema'], 'Z4': fixture['risk_engine']},
    }
    with Workbook(path) as book:
        for sheet, fields in required.items():
            for address, expected in fields.items():
                if book[sheet][address].value != expected:
                    raise ValueError(f'{path.name} does not match the fixture at {sheet}!{address}.')


def load_fixture(directory: Path) -> dict[str, Any]:
    """Verify the manifest, source pin and untouched desktop input."""
    fixture = read_json(directory / 'fixture.json')
    case, pairs = fixture.get('fixture'), fixture.get('pairs')
    if (fixture.get('schema') != FIXTURE_SCHEMA
            or not isinstance(case, str) or case not in FIXTURES
            or type(pairs) is not int or pairs not in PAIR_COUNTS
            or fixture.get('overrides') != overrides(case, pairs)
            or fixture.get('seed') != SEED or fixture.get('outcomes') != 2 * pairs + STRESS_COUNT
            or not all(isinstance(fixture.get(field), str) and fixture[field]
                       for field in ('revision', 'commit', 'source_sha256', 'input_sha256',
                                     'risk_schema', 'risk_engine'))):
        raise ValueError('The fixture manifest does not match the desktop catalog.')
    check_source(fixture['commit'], fixture['source_sha256'])
    revision = git('rev-parse', '--verify', '--end-of-options', f"{fixture['revision']}^{{commit}}").decode().strip()
    if revision != fixture['commit']:
        raise ValueError('The source revision does not resolve to the fixture commit.')
    input_path = directory / INPUT
    if digest(input_path) != fixture['input_sha256']:
        raise ValueError('The desktop input checksum changed.')
    check_workbook(input_path, fixture, pairs=pairs)
    return fixture


def prepare(source: Path, revision: str, expected: str, directory: Path, case: str, pairs: int) -> None:
    """Build a new macro-enabled fixture from the checked fictional release."""
    edits = overrides(case, pairs)
    commit = git('rev-parse', '--verify', '--end-of-options', f'{revision}^{{commit}}').decode().strip()
    check_source(commit, expected)
    if source.suffix.lower() != '.xlsm' or digest(source) != expected:
        raise ValueError('The supplied .xlsm does not match the approved checksum.')
    with Workbook(source) as book:
        schema, engine = (book['Risk_Data'][cell].value for cell in ('Z2', 'Z4'))
        changes = {sheet: {address: value for address, value in cells.items()
                           if book[sheet][address].value != value} for sheet, cells in edits.items()}
    changes = {sheet: cells for sheet, cells in changes.items() if cells}
    directory.mkdir(parents=True)
    destination = directory / INPUT
    if changes:
        write_values(source, destination, changes)
    else:
        copyfile(source, destination)
    fixture = {
        'schema': FIXTURE_SCHEMA, 'revision': revision, 'commit': commit, 'source_sha256': expected,
        'fixture': case, 'pairs': pairs, 'seed': SEED, 'outcomes': 2 * pairs + STRESS_COUNT,
        'overrides': edits, 'input_sha256': digest(destination), 'risk_schema': schema, 'risk_engine': engine,
    }
    check_workbook(destination, fixture, pairs=pairs)
    write_json(directory / 'fixture.json', fixture)


def make_reference(directory: Path, soffice: str) -> None:
    """Reuse the pinned standard snapshot or run the pinned independent bootstrap."""
    fixture = load_fixture(directory)
    case = fixture['fixture']
    if case == 'example-95':
        with (directory / 'reference.xlsm').open('xb') as stream:
            stream.write(git('show', f"{fixture['commit']}:{SOURCE}"))
        return
    reference = directory.resolve() / 'reference.xlsx'
    if reference.exists() or (directory / 'reference.json').exists():
        raise ValueError('The fixture already has a reference. Use a new fixture directory to rebuild it.')
    with TemporaryDirectory(dir=directory) as temporary:
        build = Path(temporary).resolve()
        archive = git('archive', '--format=zip', fixture['commit'], 'tools', 'workbook/build-template.xlsx')
        with ZipFile(BytesIO(archive)) as files:
            files.extractall(build)
        tools = build / 'tools'
        template = build / 'workbook/build-template.xlsx'
        plan, authored, draft, edited = (build / name for name in ('plan.json', 'author.xlsx', 'draft.xlsx', 'edited.xlsx'))
        for script, arguments in (
            ('model_plan.py', (template, plan)), ('author_model.py', (plan, authored)),
            ('package_model.py', (template, plan, authored, draft)),
        ):
            subprocess.run([sys.executable, str(tools / script), *map(str, arguments)], check=True)
        write_values(draft, edited, fixture['overrides'])
        subprocess.run([sys.executable, str(tools / 'native_bootstrap.py'), str(edited),
                        str(reference), '--soffice', soffice], check=True)
    check_workbook(reference, fixture, pairs=fixture['pairs'])
    write_json(directory / 'reference.json', {'fixture_sha256': digest(directory / 'fixture.json'),
                                              'reference_sha256': digest(reference)})


def main() -> int:
    """Dispatch fixture preparation or reference generation."""
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    build = commands.add_parser('prepare', help='create a new desktop fixture directory')
    build.add_argument('source', type=Path)
    build.add_argument('directory', type=Path)
    build.add_argument('--revision', required=True)
    build.add_argument('--sha256', required=True)
    build.add_argument('--fixture', required=True, choices=FIXTURES)
    build.add_argument('--pairs', required=True, type=int, choices=PAIR_COUNTS)
    reference = commands.add_parser('reference', help='make the independent comparison artifact')
    reference.add_argument('directory', type=Path)
    reference.add_argument('--soffice', default='soffice')
    args = parser.parse_args()
    try:
        if args.command == 'prepare':
            prepare(args.source, args.revision, args.sha256, args.directory, args.fixture, args.pairs)
        else:
            make_reference(args.directory, args.soffice)
    except (OSError, ValueError, KeyError, BadZipFile, subprocess.CalledProcessError) as error:
        parser.exit(1, f'{error}\n')
    print(f'{args.command}: {args.directory}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
