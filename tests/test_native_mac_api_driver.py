"""Check the Mac driver lifecycle with a file-only Excel substitute."""

from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, call, patch

import mac_funding_proof_driver as driver
from desktop_benchmark import digest, read_json


class MemoryExcel(driver.MacExcel):
    """Record API dispatch and simulate a saved workbook for each completed run."""

    def __init__(self, directory: Path, arm: str) -> None:
        """Set the one path the simulated Excel process will acknowledge."""
        self.directory, self.arm = directory, arm
        self.process: str | None = None
        self.actions: list[str] = []
        self.runs = 0
        self.failure = ''
        self.closed = False

    def pid(self) -> str | None:
        """Return the simulated host identity."""
        return self.process

    def environment(self, bundle: Path, machine: str, channel: str) -> dict:
        """Supply deterministic environment metadata without accessing an installed app."""
        assert bundle == Path('/fixture/Excel.app')
        return {'platform': 'mac', 'excel_bitness': 64, 'excel_build': 'fixture', 'excel_channel': channel,
                'architecture': 'arm64', 'os_build': 'fixture', 'machine': machine,
                'distribution': 'local guarded diagnostic fixture', 'macro_policy': 'existing Excel automation security',
                'timing_method': driver.MACRO_WALL_TIME}

    def call(self, action: str, workbook: Path, expected_identity: str = '') -> str:
        """Refuse alternate paths and write only the emulated session workbook."""
        assert workbook == self.directory / driver.SESSION
        self.actions.append(action)
        if action == 'open':
            assert digest(workbook) == expected_identity
            self.process = '71'
            return 'wrong policy' if self.failure == 'open' else 'policy-unchanged|/fixture/Excel.app'
        if action == 'run':
            self.runs += 1
            if self.failure == 'pid':
                self.process = '72'
            return 'returned'
        if action == 'profile' and self.runs == 0:
            schema = driver.GUARDED_PROFILE if self.arm == 'funding-proof' else driver.PROFILE
            return '' if self.failure == 'trust' else f'{schema}|unavailable'
        if action == 'profile':
            if self.failure == 'receipt':
                return 'incomplete receipt'
            schema = driver.GUARDED_PROFILE if self.arm == 'funding-proof' else driver.PROFILE
            return f'{schema}|complete|mac-raw|{self.runs}|' + '0.000001|' * 10 + '1532|501|33|3|1535|1|0'
        if action == 'guarded':
            return f'GuardedCalculation/v1|{self.runs}|2|0'
        if action == 'proof':
            return 'FundingProofCounts/v1|1538|1538|2|2|1536|0|0|0|0'
        if action == 'save':
            phase, repeat = driver.WARM_SEQUENCE[self.runs - 1]
            raw = read_json(self.directory / f'raw-{phase}-{repeat:02}.json')
            assert raw['profile'].split('|')[3] == str(self.runs)
            if self.failure == 'save':
                raise ValueError('Unavailable native save')
            workbook.write_text(f'capture {self.runs}', encoding='ascii')
            return driver.CONFIRMED_SAVE
        if action == 'close':
            prefix = 'evidence' if self.arm == 'funding-proof' else 'profile-evidence'
            assert (self.directory / f'{prefix}-warm-03.json').is_file()
            assert expected_identity == self.process and not self.closed
            self.closed = True
            if self.failure == 'close-pid':
                self.process = '72'
            return 'closed'
        if action == 'quit-empty':
            assert self.closed and expected_identity == self.process
            self.process = None
            return 'quit-requested'
        return f'{action}-{self.runs}'


class MacDriverTests(unittest.TestCase):
    def setUp(self) -> None:
        """Use tiny regular files and mock only the costly package and financial validators."""
        self.root = Path(self.enterContext(TemporaryDirectory())).resolve()
        self.source = self.root / 'template'
        self.source.mkdir()
        for name in (driver.INPUT, 'fixture.json', 'reference.json', 'reference.xlsx', driver.CONTROL,
                     driver.PROOF, driver.BANK):
            (self.source / name).write_text(name, encoding='ascii')
        self.destination = self.root / 'output'
        self.fixture = self.enterContext(patch.object(driver, 'full_scale_fixture', return_value={'pairs': 1000}))
        self.reference = self.enterContext(patch.object(driver, 'reference_file'))
        self.proof_package = self.enterContext(patch.object(driver, 'verify_candidate'))
        self.control_package = self.enterContext(patch.object(driver, 'check_candidate'))
        self.proof_replay = self.enterContext(patch.object(driver, 'replay_capture', side_effect=self.reconcile))
        self.control_replay = self.enterContext(patch.object(driver, 'record_control', side_effect=self.reconcile_control))

    def reconcile(self, directory: Path, capture: Path, raw: dict, **context: object) -> dict:
        """Verify each input and model the existing receipt-bound financial replay."""
        del context
        self.assertEqual(directory, capture.parent)
        count = int(raw['profile'].split('|')[3])
        return {'capture_sha256': digest(capture), 'profile': {'batch_serial': count}}

    def reconcile_control(self, directory: Path, name: str, timing: dict, profile: str) -> dict:
        """Use the same lightweight capture checks for the full-first arm."""
        del timing
        return self.reconcile(directory, directory / name, {'profile': profile})

    def execute(self, excel: MemoryExcel, destination: Path | None = None) -> list[dict]:
        """Run the orchestration without launching or querying native Excel."""
        with redirect_stdout(StringIO()):
            return driver.run_session(self.source, destination or self.destination, excel.arm,
                                      machine='fixture Mac', channel='not observed', excel=excel)

    def test_records_both_arms_with_exact_predecessors(self) -> None:
        """Verify the canonical receipt sequence and leave the input bundle byte-identical."""
        for arm in driver.ARMS:
            with self.subTest(arm=arm):
                directory = self.root / arm
                excel = MemoryExcel(directory, arm)
                results = self.execute(excel, directory)
                self.assertEqual(len(results), 4)
                self.assertEqual(excel.actions[:2], ['open', 'profile'])
                self.assertEqual(excel.actions[-2:], ['close', 'quit-empty'])
                self.assertIsNone(excel.pid())
                self.assertEqual(excel.actions.count('save'), 4)
                proof = arm == 'funding-proof'
                prefix = 'evidence' if proof else 'profile-evidence'
                candidate = driver.PROOF if proof else driver.CONTROL
                preceding = digest(self.source / candidate)
                opened = read_json(directory / 'api-open.json')
                self.assertEqual(opened['input_sha256'], preceding)
                self.assertIs(opened['temporary_automation_policy'], False)
                for phase, repeat in driver.WARM_SEQUENCE:
                    label = f'{phase}-{repeat:02}'
                    raw = read_json(directory / f'raw-{label}.json')
                    self.assertEqual((raw['starting_sha256'], raw['excel_pid']), (preceding, '71'))
                    self.assertEqual('proof' in raw, proof)
                    self.assertEqual(bool(raw['counts']), proof)
                    preceding = digest(directory / f'capture-{label}.xlsm')
                    self.assertEqual(read_json(directory / f'save-{label}.json')['capture_sha256'], preceding)
                    self.assertEqual(read_json(directory / f'{prefix}-{label}.json')['capture_sha256'], preceding)
                self.assertEqual(digest(directory / candidate), digest(self.source / candidate))
        self.control_package.assert_called_once_with(self.source)
        self.proof_package.assert_called_once_with(self.source)
        self.assertEqual((self.control_replay.call_count, self.proof_replay.call_count), (4, 4))

    def test_invalid_package_or_existing_excel_never_opens(self) -> None:
        """Reject an unverified workbook and an already running Excel process before dispatch."""
        excel = MemoryExcel(self.destination, 'funding-proof')
        self.proof_package.side_effect = ValueError('unverified package')
        with self.assertRaisesRegex(ValueError, 'unverified package'):
            self.execute(excel)
        self.assertFalse(self.destination.exists())
        excel.process = '88'
        with self.assertRaisesRegex(ValueError, 'Excel closed'):
            self.execute(excel)
        self.assertEqual(excel.actions, [])

    def test_failed_open_trust_pid_receipt_or_save_stops_without_retrying(self) -> None:
        """Retain raw evidence on a failed receipt or save and never repeat a batch."""
        for failure, message in (('open', 'confirm opening'), ('trust', 'fresh, accessible'),
                                 ('pid', 'changed Excel processes'), ('receipt', 'guarded calculation stage profile'),
                                 ('save', 'native save')):
            with self.subTest(failure=failure):
                directory = self.root / failure
                excel = MemoryExcel(directory, 'funding-proof')
                excel.failure = failure
                with self.assertRaisesRegex(ValueError, message):
                    self.execute(excel, directory)
                self.assertEqual(excel.runs, int(failure in ('pid', 'receipt', 'save')))
                self.assertEqual((directory / 'raw-cold-01.json').exists(), failure in ('receipt', 'save'))
                if failure == 'receipt':
                    self.assertNotIn('save', excel.actions)
                self.assertFalse((directory / 'capture-cold-01.xlsm').exists())
                self.assertNotIn('close', excel.actions)
                self.assertNotIn('quit-empty', excel.actions)

    def test_working_copy_must_match_the_immediate_saved_predecessor(self) -> None:
        """Detect a changed working file even when the first independent replay passed."""
        excel = MemoryExcel(self.destination, 'funding-proof')

        def replay_then_change(directory: Path, capture: Path, raw: dict, **context: object) -> dict:
            """Simulate a separate edit after the first completed evidence check."""
            record = self.reconcile(directory, capture, raw, **context)
            (directory / driver.SESSION).write_text('different saved workbook', encoding='ascii')
            return record

        self.proof_replay.side_effect = replay_then_change
        with self.assertRaisesRegex(ValueError, 'verified predecessor'):
            self.execute(excel)
        self.assertEqual(excel.runs, 1)
        self.assertNotIn('close', excel.actions)

    def test_handoff_refuses_a_different_process_after_close(self) -> None:
        """Leave the replacement process alone and preserve all verified files."""
        excel = MemoryExcel(self.destination, 'funding-proof')
        excel.failure = 'close-pid'
        with self.assertRaisesRegex(ValueError, 'changed Excel processes'):
            self.execute(excel)
        self.assertEqual(excel.actions[-1], 'close')
        self.assertEqual(excel.pid(), '72')
        self.assertTrue((self.destination / 'evidence-warm-03.json').is_file())

    def test_handoff_waits_for_exit_and_times_out_if_excel_stays_running(self) -> None:
        """Wait only for the original process, propagating a failed handoff without closing anything else."""
        working = self.destination / driver.SESSION
        for remaining, elapsed, error in ((None, 1, ''), ('71', driver.EXIT_TIMEOUT, 'handoff timeout'),
                                           ('72', 1, 'Another Excel process')):
            with self.subTest(remaining=remaining):
                excel = Mock(spec=driver.MacExcel)
                excel.pid.side_effect = ['71', '71', '71', remaining]
                excel.call.side_effect = ['closed', 'quit-requested']
                with (patch.object(driver, 'monotonic', side_effect=(0, elapsed)),
                      patch.object(driver, 'sleep') as sleep):
                    if error:
                        with self.assertRaisesRegex(ValueError, error):
                            driver.close_session(excel, working, '71')
                    else:
                        driver.close_session(excel, working, '71')
                        self.assertEqual(excel.pid.call_count, 4)
                        sleep.assert_called_once()
                self.assertEqual(excel.call.call_args_list, [call('close', working, '71'),
                                                           call('quit-empty', working, '71')])

    def test_applescript_checks_paths_and_restores_only_save_alerts(self) -> None:
        """Guard the tracked host code without compiling it against a live Excel installation."""
        source = driver.SCRIPT.read_text(encoding='utf-8')
        self.assertNotIn('set automation security', source.casefold())
        self.assertNotIn('System Events', source)
        path_check = source.index('actualPath is not expectedPath')
        self.assertLess(path_check, source.index('run VB Macro'))
        self.assertLess(path_check, source.index('save book'))
        self.assertLess(path_check, source.index('close book saving no'))
        self.assertIn('my requireExcelProcess(item 4 of argv)\n                close book saving no', source)
        quit_branch = source.split('if actionName is "quit-empty" then', 1)[1].split('return "quit-requested"', 1)[0]
        self.assertLess(quit_branch.index('(count workbooks) is not 0'), quit_branch.index('            quit'))
        self.assertLess(quit_branch.index('requireExcelProcess'), quit_branch.index('            quit'))
        self.assertIn('if actionName is "run" then set timeoutSeconds to 7200', source)
        self.assertNotIn('with timeout of 7200', source)
        save = source[source.index('set alertsBefore to display alerts'):]
        self.assertEqual(save.count('set display alerts to alertsBefore'), 2)
        self.assertLess(save.index('on error saveText'), save.index('set display alerts to alertsBefore'))
        self.assertIn('error saveText number saveNumber', save)
