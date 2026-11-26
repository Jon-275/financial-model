"""Replay miniature evidence bundles without Excel or workbook reconciliation work."""

from contextlib import nullcontext
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import desktop_funding_proof_replay as checker
from desktop_benchmark import digest, read_json, write_json
from package_macro import SOURCE


FIXTURE = {'fixture': 'full-0-125', 'pairs': 1000, 'seed': 2026, 'source_sha256': 'a' * 64}
ENVIRONMENT = {'platform': 'mac', 'excel_bitness': 64, 'excel_build': 'fixture build',
               'excel_channel': 'not observed', 'os_build': 'fixture macOS', 'architecture': 'arm64',
               'machine': 'fixture Mac', 'distribution': 'local guarded diagnostic fixture',
               'macro_policy': 'per-file Excel prompt', 'timing_method': checker.MACRO_WALL_TIME}


def replace_json(path: Path, **changes: object) -> None:
    """Alter one recorded claim independently of its artifact."""
    record = {**read_json(path), **changes}
    path.unlink()
    write_json(path, record)


class MatchedFundingReplayTests(unittest.TestCase):
    def setUp(self) -> None:
        """Provide independent JSON, byte identities and typed native receipts for eight runs."""
        folder = self.enterContext(TemporaryDirectory())
        self.control, self.candidate = Path(folder) / 'control', Path(folder) / 'candidate'
        for directory in (self.control, self.candidate):
            directory.mkdir()
            write_json(directory / 'fixture.json', FIXTURE)
            (directory / 'reference.xlsx').write_bytes(b'synthetic independent reference')
        self.package = self.candidate / checker.CANDIDATE
        self.package.write_bytes(b'synthetic natural runner package')
        self.reference = self.candidate / 'reference.xlsx'
        self.open_receipt = self.candidate / 'api-open.json'
        write_json(self.open_receipt, {'input_sha256': digest(self.package), 'automation_policy_restored': True,
                                      'temporary_automation_policy': False, 'display_state': 'unlocked'})
        self.enterContext(patch.object(checker, 'load_fixture', return_value=FIXTURE))
        self.enterContext(patch.object(checker, 'verify_candidate', return_value=self.package))
        self.enterContext(patch.object(checker, 'reference_file', return_value=self.reference))
        self.reconcile = self.enterContext(patch.object(checker, 'compare_capture', side_effect=self.reconcile_file))
        control_package = self.control / 'desktop-profile-input.xlsm'
        control_package.write_bytes(b'synthetic full-first package')
        starting = digest(control_package)
        self.controls = []
        for i, ((phase, repeat), label) in enumerate(zip(checker.WARM_SEQUENCE, checker.LABELS, strict=True)):
            capture = self.control / f'capture-{label}.xlsm'
            capture.write_text(f'control capture {i}', encoding='ascii')
            timing = {**ENVIRONMENT, 'phase': phase, 'repeat': repeat, 'seconds': (80, 70, 60, 90)[i]}
            control = {**FIXTURE, 'run': timing, 'profile': {'batch_serial': 100 + i},
                       'payload': f'control profile {i}', 'candidate_sha256': digest(control_package),
                       'capture_sha256': digest(capture), 'fixture_sha256': digest(self.control / 'fixture.json'),
                       'reference_sha256': digest(self.reference)}
            receipt = {'excel_pid': '11', 'starting_sha256': starting, 'timing': timing,
                       'profile': control['payload'], 'counts': ''}
            self.controls.append((control, receipt))
            starting = control['capture_sha256']
        self.control_replay = self.enterContext(patch.object(checker, 'replay_warm_arm', return_value=self.controls))

        starting = digest(self.package)
        for i, ((phase, repeat), label) in enumerate(zip(checker.WARM_SEQUENCE, checker.LABELS, strict=True)):
            capture = self.candidate / f'capture-{label}.xlsm'
            capture.write_text(f'candidate capture {i}', encoding='ascii')
            claim = self.candidate / f'evidence-{label}.json'
            serial = 200 + i
            timing = {**ENVIRONMENT, 'phase': phase, 'repeat': repeat, 'seconds': (100, 42, 35, 49)[i]}
            profile = f'RunnerGuardedProfile/v1|complete|mac-raw|{serial}|' + '1|' * 10 + '1532|501|33|3|1535|1|0'
            receipt = {'excel_pid': '12', 'starting_sha256': starting, 'timing': timing, 'profile': profile,
                       'counts': f'GuardedCalculation/v1|{serial}|2|0',
                       'proof': 'FundingProofCounts/v1|1538|1538|2|2|1536|0|0|0|0'}
            write_json(self.candidate / f'raw-{label}.json', receipt)
            replay = checker.replay_capture(self.candidate, capture, receipt, fixture=FIXTURE,
                                            candidate=self.package, reference=self.reference)
            write_json(claim, replay)
            starting = digest(capture)
            write_json(self.candidate / f'save-{label}.json', {'capture_sha256': starting,
                       'method': 'standard Excel save', 'display_state': 'not observed'})
        self.reconcile.reset_mock()

    def reconcile_file(self, directory: Path, fixture: dict, reference: Path, capture: Path, *,
                       candidate: Path, expected_reused: int) -> tuple[str, float]:
        """Stand in for the existing independent financial field comparator."""
        self.assertEqual((directory, fixture, reference, candidate, expected_reused),
                         (self.candidate, FIXTURE, self.reference, self.package, 501))
        index = checker.LABELS.index(capture.stem.removeprefix('capture-'))
        return f'PASS reconciled capture {index}', 200.0 + index

    def test_replays_files_and_keeps_cold_and_api_metadata_separate(self) -> None:
        """Exclude cold timing from the median and preserve API handling as its own receipt."""
        output = checker.replay(self.control, self.candidate)
        self.assertEqual(output['warm_median_seconds'], {'full_first': 70, 'funding_proof': 42})
        self.assertEqual(output['warm_improvement_percent'], 40)
        self.assertEqual(output['run_environment']['macro_policy'], 'per-file Excel prompt')
        self.assertIs(output['candidate_excel_api']['open']['temporary_automation_policy'], False)
        self.assertEqual(len(output['arms']['funding_proof']), 4)
        self.assertEqual(self.reconcile.call_count, 4)
        self.control_replay.assert_called_once_with(
            [self.control / f'profile-evidence-{label}.json' for label in checker.LABELS], guarded=False)

    def test_rejects_changed_candidate_process_and_saved_predecessor(self) -> None:
        """Reject either a different open session or a break in its ordered capture chain."""
        raw = self.candidate / 'raw-warm-02.json'
        original = read_json(raw)
        for delta, expected in (({'excel_pid': '43'}, 'same Excel process'),
                                ({'starting_sha256': 'b' * 64}, 'preceding saved capture')):
            with self.subTest(delta=delta):
                replace_json(raw, **(original | delta))
                with self.assertRaisesRegex(ValueError, expected):
                    checker.replay(self.control, self.candidate)

    def test_rejects_other_scale_or_a_shared_cold_process(self) -> None:
        """Require the pinned experiment and independent cold starts for the two arms."""
        with patch.object(checker, 'load_fixture', return_value={**FIXTURE, 'pairs': 100}):
            with self.assertRaisesRegex(ValueError, '1,000-pair fixture'):
                checker.replay(self.control, self.candidate)
        for _, receipt in self.controls:
            receipt['excel_pid'] = '12'
        with self.assertRaisesRegex(ValueError, 'own Excel process'):
            checker.replay(self.control, self.candidate)

    def test_rejects_receipt_or_saved_batch_disagreement(self) -> None:
        """Detect edits to waterfall counts and a mismatch in the financial batch identity."""
        raw = self.candidate / 'raw-cold-01.json'
        original = read_json(raw)
        replace_json(raw, proof='FundingProofCounts/v1|1536|1536|0|0|1536|0|0|0|0')
        with self.assertRaisesRegex(ValueError, 'verified runner attempts'):
            checker.replay(self.control, self.candidate)
        replace_json(raw, **original)
        self.reconcile.side_effect = None
        self.reconcile.return_value = 'PASS other capture', 99.0
        with self.assertRaisesRegex(ValueError, 'saved batch identity'):
            checker.replay(self.control, self.candidate)

    def test_rejects_mislabeled_capture_or_mismatched_environment(self) -> None:
        """Require the catalogue chronology and common configured environment after replay."""
        raw = self.candidate / 'raw-warm-01.json'
        claim = self.candidate / 'evidence-warm-01.json'
        original = read_json(raw)['timing']
        for delta, expected in (({'repeat': 2}, 'File order'), ({'excel_build': 'other build'}, 'different excel_build')):
            with self.subTest(delta=delta):
                timing = original | delta
                replace_json(raw, timing=timing)
                replace_json(claim, run=timing)
                with self.assertRaisesRegex(ValueError, expected):
                    checker.replay(self.control, self.candidate)

    def test_rejects_unbound_api_receipts(self) -> None:
        """Verify both the opening policy flag and save hashes against replayed files."""
        rows = [read_json(self.candidate / f'evidence-{label}.json') for label in checker.LABELS]
        replace_json(self.open_receipt, automation_policy_restored=1)
        with self.assertRaisesRegex(ValueError, 'API open receipt'):
            checker.check_api_receipts(self.candidate, rows)
        replace_json(self.open_receipt, automation_policy_restored=True)
        replace_json(self.candidate / 'save-warm-03.json', capture_sha256='c' * 64)
        with self.assertRaisesRegex(ValueError, 'API save receipt'):
            checker.check_api_receipts(self.candidate, rows)


class ProofCandidatePackageTests(unittest.TestCase):
    def test_checks_natural_overlay_and_unmodified_bank_parts(self) -> None:
        """Verify a lightweight ZIP boundary and feed the real expected natural VBA transform."""
        source = SOURCE.read_text(encoding='ascii')
        expected = checker.instrument(checker.guarded_overlay(source, mode='natural'))
        with TemporaryDirectory() as folder:
            directory = Path(folder)
            bank, candidate = (directory / name for name in (checker.BANK, checker.CANDIDATE))
            for path, vba in ((bank, b'bank'), (candidate, b'candidate')):
                with ZipFile(path, 'w') as archive:
                    archive.writestr(checker.VBA_PART, vba)
                    archive.writestr('worksheet.xml', b'identical')
            control = {checker.RUNNER: (), 'Sheet1': ('preserved',)}

            def project(path: Path) -> object:
                """Return the archived source side without opening a heavy workbook."""
                text = source if path == bank else candidate_source
                return nullcontext(SimpleNamespace(get_module=lambda name: text))

            candidate_source = expected
            with (patch.object(checker, 'verify_package', return_value=bank) as verify_bank,
                  patch.object(checker, 'vba_identity', return_value=(control, ())),
                  patch.object(checker, 'ExcelFile', side_effect=project)):
                self.assertEqual(checker.verify_candidate(directory), candidate)
                verify_bank.assert_called_once_with(directory / checker.INPUT, bank)
                candidate_source = expected + "\n' unexpected edit"
                with self.assertRaisesRegex(ValueError, 'reviewed natural VBA overlay'):
                    checker.verify_candidate(directory)
                with ZipFile(candidate, 'w') as archive:
                    archive.writestr(checker.VBA_PART, b'candidate')
                    archive.writestr('worksheet.xml', b'other')
                with self.assertRaisesRegex(ValueError, 'only the verified bank'):
                    checker.verify_candidate(directory)
