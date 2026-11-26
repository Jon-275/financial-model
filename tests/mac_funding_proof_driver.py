"""Run a pinned full-scale control or funding-proof Mac session via Excel's API."""

from argparse import ArgumentParser
from dataclasses import dataclass
from pathlib import Path
import platform
import plistlib
import re
from shutil import copyfileobj
from subprocess import CalledProcessError, check_output, run
import sys
from time import perf_counter
from typing import Any

from desktop_benchmark import INPUT, digest, write_json
from desktop_evidence import reference_file, validate_run
from desktop_funding_proof_replay import BANK, CANDIDATE as PROOF, full_scale_fixture, replay_capture, verify_candidate
from desktop_guarded_calculation import GUARDED_PROFILE, MACRO_WALL_TIME, WARM_SEQUENCE, read_profile
from desktop_stage_profile import CANDIDATE as CONTROL, PROFILE, check_candidate, parse, record as record_control
from native_funding_proof_overlay import parse_counts


ARMS = ('full-first', 'funding-proof')
SCRIPT = Path(__file__).with_suffix('.applescript')
SESSION = 'native-mac-test-session.xlsm'
CONFIRMED_SAVE = 'Risk results are current.'


@dataclass(frozen=True)
class Session:
    """Keep validated inputs apart from the one workbook Excel may save."""

    directory: Path
    working: Path
    candidate: Path
    reference: Path
    fixture: dict[str, Any]
    hashes: dict[str, str]


class MacExcel:
    """Send fixed operations to the exact workbook via the native object model."""

    def pid(self) -> str | None:
        """Return the sole Excel process ID or fail when it is ambiguous."""
        found = run(['/usr/bin/pgrep', '-x', 'Microsoft Excel'], capture_output=True, text=True)
        process = found.stdout.strip()
        if found.returncode == 1 and not process:
            return None
        if found.returncode != 0 or not re.fullmatch(r'[1-9][0-9]*', process):
            raise ValueError('Expected at most one Microsoft Excel process.')
        return process

    def call(self, action: str, workbook: Path, expected_hash: str = '') -> str:
        """Send a scoped Apple event and return its text receipt."""
        return check_output(['/usr/bin/osascript', str(SCRIPT), action, workbook.name, str(workbook),
                             expected_hash], text=True).strip()

    def environment(self, bundle: Path, machine: str, channel: str) -> dict[str, Any]:
        """Record the installed build and configured session metadata for both arms."""
        with (bundle / 'Contents' / 'Info.plist').open('rb') as stream:
            app = plistlib.load(stream)
        if app.get('CFBundleIdentifier') != 'com.microsoft.Excel':
            raise ValueError('The open API did not identify the Microsoft Excel application bundle.')
        version, build = (check_output(['/usr/bin/sw_vers', option], text=True).strip()
                          for option in ('-productVersion', '-buildVersion'))
        return {'platform': 'mac', 'excel_bitness': 64, 'excel_build': app['CFBundleVersion'],
                'excel_channel': channel, 'os_build': f'macOS {version} ({build})',
                'architecture': platform.machine(), 'machine': machine,
                'distribution': 'local guarded diagnostic fixture',
                'macro_policy': 'existing Excel automation security', 'timing_method': MACRO_WALL_TIME}


def copy_new(source: Path, destination: Path) -> None:
    """Copy without overwriting an existing input, capture, or session."""
    with source.open('rb') as original, destination.open('xb') as copy:
        copyfileobj(original, copy)


def prepare(source: Path, destination: Path, arm: str) -> Session:
    """Authenticate a catalogue bundle and copy its exact bytes into a new session directory."""
    if arm not in ARMS:
        raise ValueError('Choose the full-first or natural funding-proof arm.')
    if destination.exists() or destination.is_symlink():
        raise FileExistsError('Use a new evidence directory for each arm.')
    proof = arm == 'funding-proof'
    template = PROOF if proof else CONTROL
    names = (INPUT, 'fixture.json', 'reference.json', 'reference.xlsx', template, *((BANK,) if proof else ()))
    artifacts = {name: source / name for name in names}
    if any(path.is_symlink() or not path.is_file() for path in artifacts.values()):
        raise ValueError('Use a complete local bundle of regular fixture files.')
    hashes = {name: digest(path) for name, path in artifacts.items()}
    fixture = full_scale_fixture(source)
    reference_file(source, fixture)
    if proof:
        verify_candidate(source)
    else:
        check_candidate(source)

    destination.mkdir()
    directory = destination.resolve(strict=True)
    for name, artifact in artifacts.items():
        copy_new(artifact, directory / name)
    if any(digest(directory / name) != expected for name, expected in hashes.items()):
        raise ValueError('An input artifact changed while the benchmark was being prepared.')
    working = directory / SESSION
    copy_new(directory / template, working)
    if digest(working) != hashes[template]:
        raise ValueError('The working copy differs from the verified synthetic candidate.')
    return Session(directory, working, directory / template, directory / 'reference.xlsx', fixture, hashes)


def same_process(excel: MacExcel, pid: str) -> None:
    """Require every native operation to stay in the original cold Excel process."""
    if excel.pid() != pid:
        raise ValueError('The benchmark changed Excel processes.')


def invoke(excel: MacExcel, action: str, working: Path, pid: str) -> str:
    """Bind the host PID before and after the path-checked AppleScript call."""
    same_process(excel, pid)
    result = excel.call(action, working)
    same_process(excel, pid)
    return result


def run_session(source: Path, destination: Path, arm: str, *, machine: str, channel: str,
                excel: MacExcel) -> list[dict[str, Any]]:
    """Record once cold and three times warm, stopping on any failed session or evidence check."""
    if not machine.strip() or not channel.strip():
        raise ValueError('Record a nonempty machine description and Excel channel.')
    if excel.pid() is not None:
        raise ValueError('Start the cold arm with Microsoft Excel closed.')
    session = prepare(source, destination, arm)
    proof = arm == 'funding-proof'
    if excel.pid() is not None:
        raise ValueError('Excel started before the verified cold workbook was opened.')
    opening_hash = session.hashes[session.candidate.name]
    opened = excel.call('open', session.working, opening_hash)
    marker, separator, bundle = opened.partition('|')
    pid = excel.pid()
    if marker != 'policy-unchanged' or not separator or not Path(bundle).is_absolute() or pid is None:
        raise ValueError('Excel did not confirm opening the exact fixture with its existing security setting.')
    write_json(session.directory / 'api-open.json', {'input_sha256': opening_hash,
               'automation_policy_restored': True, 'temporary_automation_policy': False,
               'display_state': 'not observed'})
    environment = excel.environment(Path(bundle), machine, channel)
    if invoke(excel, 'profile', session.working, pid) != f'{GUARDED_PROFILE if proof else PROFILE}|unavailable':
        raise ValueError('The opened workbook does not have a fresh, accessible test runner.')

    preceding_hash, preceding_batch = opening_hash, 0.0
    records = []
    for phase, repeat in WARM_SEQUENCE:
        label = f'{phase}-{repeat:02}'
        if session.working.is_symlink() or digest(session.working) != preceding_hash:
            raise ValueError('The working workbook differs from its verified predecessor.')
        same_process(excel, pid)
        start = perf_counter()
        returned = excel.call('run', session.working)
        elapsed = perf_counter() - start
        same_process(excel, pid)
        if returned != 'returned':
            raise ValueError('Excel did not finish the synchronous workbook runner.')

        timing = {**environment, 'phase': phase, 'repeat': repeat, 'seconds': elapsed}
        validate_run(timing)
        profile = invoke(excel, 'profile', session.working, pid)
        counts = invoke(excel, 'guarded', session.working, pid) if proof else ''
        raw = {'excel_pid': pid, 'starting_sha256': preceding_hash, 'profile': profile,
               'counts': counts, 'timing': timing}
        if proof:
            raw['proof'] = invoke(excel, 'proof', session.working, pid)
        write_json(session.directory / f'raw-{label}.json', raw)
        if proof:
            live = read_profile(profile, counts, seconds=elapsed, mode='natural', pairs=1000)
            parse_counts(raw['proof'])
        else:
            live = parse(profile, pairs=1000, platform='mac', seconds=elapsed)
        if live['batch_serial'] <= preceding_batch:
            raise ValueError('The live runner did not complete a distinct newer batch.')

        if invoke(excel, 'save', session.working, pid) != CONFIRMED_SAVE:
            raise ValueError('Excel did not verify the completed workbook save.')
        capture = session.directory / f'capture-{label}.xlsm'
        if session.working.is_symlink():
            raise ValueError('The working workbook no longer refers to its local session copy.')
        copy_new(session.working, capture)
        saved_hash = digest(capture)
        write_json(session.directory / f'save-{label}.json', {'capture_sha256': saved_hash,
                   'method': 'scoped alert setting, restored', 'display_state': 'not observed'})
        if (saved_hash == preceding_hash or digest(session.working) != saved_hash
                or any(digest(session.directory / name) != expected for name, expected in session.hashes.items())):
            raise ValueError('The saved capture or its immutable benchmark inputs changed.')
        if proof:
            record = replay_capture(session.directory, capture, raw, fixture=session.fixture,
                                    candidate=session.candidate, reference=session.reference)
        else:
            record = record_control(session.directory, capture.name, timing, profile)
        batch = record['profile']['batch_serial']
        if record['capture_sha256'] != saved_hash or batch <= preceding_batch:
            raise ValueError('The verified capture did not publish a distinct newer batch.')
        prefix = 'evidence' if proof else 'profile-evidence'
        write_json(session.directory / f'{prefix}-{label}.json', record)
        records.append(record)
        preceding_hash, preceding_batch = saved_hash, batch
        print(f'{arm} {label}: verified {elapsed:.3f} seconds', flush=True)
    return records


def main() -> int:
    """Run one complete disposable arm on macOS with its current trust configuration."""
    parser = ArgumentParser(description=__doc__, epilog='Close Excel before the cold run. A trusted macro must already be available under the installed Excel setting.')
    parser.add_argument('arm', choices=ARMS)
    parser.add_argument('source', type=Path, help='prepared evidence bundle containing the pinned input and reference')
    parser.add_argument('destination', type=Path, help='new directory for the isolated session and its receipts')
    parser.add_argument('--machine', required=True, help='matching machine description for both benchmark arms')
    parser.add_argument('--excel-channel', default='not observed', help='configured Excel channel if known')
    args = parser.parse_args()
    if sys.platform != 'darwin':
        parser.error('The Excel API benchmark requires macOS.')
    try:
        run_session(args.source, args.destination, args.arm, machine=args.machine, channel=args.excel_channel,
                    excel=MacExcel())
    except (OSError, ValueError, CalledProcessError, KeyError) as error:
        parser.error(str(error))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
