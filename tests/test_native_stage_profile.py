"""Exercise the fixture overlay and the local stage evidence boundary."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from pyopenvba import ExcelFile

import desktop_benchmark as benchmark
import desktop_evidence
import desktop_stage_profile as profile


PAYLOAD = ('RunnerStageProfile/v1|complete|mac-raw|50000.125|1|1|30|3|4|2|.4|.6|1|1'
           '|200|33|33|3|203|12|8')


class StageProfileTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        """Prepare a checked reference and a diagnostic copy of the real fixture."""
        cls.temporary = TemporaryDirectory()
        cls.addClassCleanup(cls.temporary.cleanup)
        cls.directory = Path(cls.temporary.name) / 'example-95-100'
        source = benchmark.ROOT / benchmark.SOURCE
        benchmark.prepare(source, 'HEAD', benchmark.digest(source), cls.directory, 'example-95', 100)
        benchmark.make_reference(cls.directory, 'unused')
        profile.prepare(cls.directory)

    def test_overlay_changes_only_the_expected_vba(self) -> None:
        original = self.directory / benchmark.INPUT
        candidate = self.directory / profile.CANDIDATE
        with ZipFile(original) as left, ZipFile(candidate) as right:
            self.assertEqual(left.namelist(), right.namelist())
            for name in left.namelist():
                with self.subTest(part=name):
                    if name == profile.VBA_PART:
                        self.assertNotEqual(left.read(name), right.read(name))
                    else:
                        self.assertEqual(left.read(name), right.read(name))
        with ExcelFile(original) as book:
            runner = book.get_module(profile.RUNNER)
        self.assertEqual(desktop_evidence.vba_identity(candidate)[0].keys(),
                         desktop_evidence.vba_identity(original)[0].keys())
        self.assertEqual(profile.check_candidate(self.directory), candidate)
        self.assertNotIn('ReadMarketRiskProfile', runner)
        with self.assertRaises(FileExistsError):
            profile.prepare(self.directory)

    def test_overlay_rejects_absent_or_ambiguous_runner_anchors(self) -> None:
        with ExcelFile(self.directory / benchmark.INPUT) as book:
            runner = book.get_module(profile.RUNNER).replace('\r\n', '\n')
        missing = runner.replace('        selector.Value2 = code', '        selector.Value2 = 0')
        ambiguous = runner.replace('        selector.Value2 = code', '        selector.Value2 = code\n' * 2)
        for changed in (missing, ambiguous):
            with self.assertRaisesRegex(ValueError, 'diagnostic anchor'):
                profile.instrument(changed)

    def test_candidate_rejects_an_edit_outside_the_vba_project(self) -> None:
        altered = self.directory / 'changed-input.xlsm'
        benchmark.write_values(self.directory / profile.CANDIDATE, altered, {'Inputs': {'D53': 80}})
        with self.assertRaisesRegex(ValueError, 'outside its VBA'):
            profile.check_candidate(self.directory, path=altered)

    def test_counts_and_exclusive_timing_partition(self) -> None:
        data = profile.parse(PAYLOAD, pairs=100, platform='mac', seconds=44.1)
        self.assertEqual(data['seconds'], 44)
        self.assertEqual(data['counters']['computed'], 200)
        self.assertEqual(data['counters']['retry_repeated_ending'], 8)
        self.assertEqual(data['stages']['restore'], 1)

    def test_invalid_or_partial_profile_cannot_be_reconciled(self) -> None:
        for payload, message in (
            (PAYLOAD + '|junk', 'fields'),
            (PAYLOAD.replace('|complete|', '|cancelled|'), 'complete'),
            ('RunnerStageProfile/v1|unavailable', 'unavailable'),
            (PAYLOAD.replace('|30|', '|NaN|'), 'durations'),
            (PAYLOAD.replace('|30|', '|1e999|'), 'durations'),
            (PAYLOAD.replace('|30|', '|-1|'), 'durations'),
            (PAYLOAD.replace('|30|', '|90|'), 'wall time'),
            (PAYLOAD.replace('|mac-raw|', '|windows-qpc|'), 'clock'),
            (PAYLOAD.replace('|203|', '|202|'), 'counts'),
            (PAYLOAD.replace('|203|', '|203.1|'), 'counts'),
            (PAYLOAD.replace('|12|8', '|196|8'), 'counts'),
            (PAYLOAD.replace('|33|33|', '|33|32|'), 'counts'),
            (PAYLOAD.rsplit('|12|8', 1)[0] + '|0|0', 'work performed'),
        ):
            with self.subTest(payload=payload):
                with self.assertRaisesRegex(ValueError, message):
                    profile.parse(payload, pairs=100, platform='mac', seconds=44.1)

    def test_evidence_binds_source_capture_and_report_timestamp(self) -> None:
        copy = self.directory / 'profile-capture.xlsm'
        benchmark.write_values(self.directory / profile.CANDIDATE, copy, {'Risk_Data': {'Z7': 50000.125}})
        timing = {
            'platform': 'mac', 'phase': 'cold', 'repeat': 1, 'seconds': 45,
            'excel_build': 'synthetic', 'excel_channel': 'synthetic', 'excel_bitness': 64,
            'os_build': 'synthetic', 'architecture': 'arm64', 'machine': 'synthetic',
            'distribution': 'profile fixture', 'macro_policy': 'test', 'timing_method': 'test',
        }
        with patch.object(profile, 'compare_capture', return_value=('233 outcomes.', 50000.125)) as compare:
            with self.assertRaisesRegex(ValueError, '64-bit'):
                profile.record(self.directory, copy.name, {**timing, 'excel_bitness': 32}, PAYLOAD)
            record = profile.record(self.directory, copy.name, timing, PAYLOAD)
            expected_input = self.directory / profile.CANDIDATE
            self.assertEqual(compare.call_args.kwargs['candidate'], expected_input)
            file = self.directory / 'profile-evidence-cold.json'
            benchmark.write_json(file, record)
            profile.check(file)
            with self.assertRaisesRegex(ValueError, 'batch timestamp'):
                profile.record(self.directory, copy.name, timing, PAYLOAD.replace('50000.125', '50001.0'))
            bad = self.directory / 'profile-evidence-wrong-source.json'
            benchmark.write_json(bad, {**record, 'candidate_sha256': '0' * 64})
            with self.assertRaisesRegex(ValueError, 'replay'):
                profile.check(bad)
            with self.assertRaisesRegex(ValueError, 'Duplicate'):
                profile.check_all([file, file])


if __name__ == '__main__':
    unittest.main()
