"""Test fixture provenance and fail-closed desktop evidence."""

from __future__ import annotations

import contextlib
from io import StringIO
from pathlib import Path
from shutil import copyfile
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import desktop_benchmark as benchmark
import desktop_evidence as evidence
from xlsx_read import Workbook


class DesktopBenchmarkTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        """Build a real 100-pair input from the tracked workbook once."""
        cls.source = benchmark.ROOT / benchmark.SOURCE
        cls.temporary = TemporaryDirectory()
        cls.addClassCleanup(cls.temporary.cleanup)
        cls.directory = Path(cls.temporary.name) / 'case'
        cls.source_hash = benchmark.digest(cls.source)
        benchmark.prepare(cls.source, 'HEAD', cls.source_hash, cls.directory, 'example-95', 100)
        benchmark.make_reference(cls.directory, 'unused-for-published-reference')
        cls.capture = cls.directory / 'capture.xlsm'
        copyfile(cls.directory / benchmark.INPUT, cls.capture)
        cls.run_metadata = {
            'platform': 'mac', 'phase': 'cold', 'repeat': 1, 'seconds': 60,
            'excel_build': 'test-build', 'excel_channel': 'test-channel', 'excel_bitness': 64,
            'os_build': 'test-os', 'architecture': 'arm64', 'machine': 'synthetic test host',
            'distribution': 'fictional release', 'macro_policy': 'test policy',
            'timing_method': 'test observation',
        }

    def test_pins_source_and_preserves_embedded_runner(self) -> None:
        fixture = benchmark.load_fixture(self.directory)
        self.assertEqual((fixture['pairs'], fixture['seed'], fixture['outcomes']), (100, 2026, 233))
        self.assertEqual(benchmark.digest(self.directory / 'reference.xlsm'), self.source_hash)
        self.assertEqual(benchmark.digest(self.source), self.source_hash)
        with Workbook(self.directory / benchmark.INPUT) as book:
            self.assertEqual(book['Risk_Settings']['B29'].value, 100)
        with ZipFile(self.source) as source, ZipFile(self.directory / benchmark.INPUT) as fixture_book:
            self.assertEqual(source.read('xl/vbaProject.bin'), fixture_book.read('xl/vbaProject.bin'))

    def test_wrong_release_hash_does_not_create_fixture(self) -> None:
        output = self.directory / 'rejected'
        with self.assertRaisesRegex(ValueError, 'pinned source'):
            benchmark.prepare(self.source, 'HEAD', '0' * 64, output, 'example-95', 100)
        self.assertFalse(output.exists())

    def test_prior_batch_cannot_be_recorded_as_a_native_run(self) -> None:
        with self.assertRaisesRegex(ValueError, 'newly completed'):
            evidence.evidence(self.directory, self.capture.name, self.run_metadata)

    def test_replays_complete_evidence_and_rejects_wrong_or_missing_data(self) -> None:
        with patch.object(evidence, 'compare_capture', return_value='233 outcomes.') as compare:
            record = evidence.evidence(self.directory, self.capture.name, self.run_metadata)
            cases = (
                ('valid', {}, None),
                ('partial', {'checks': {'fresh_capture': 'pass'}}, 'both passing checks'),
                ('wrong-fixture', {'fixture': 'extended-125'}, 'artifact hashes'),
                ('wrong-candidate', {'source_sha256': '0' * 64}, 'artifact hashes'),
                ('wrong-capture', {'capture_sha256': '0' * 64}, 'artifact hashes'),
                ('wrong-reference', {'reference_sha256': '0' * 64}, 'artifact hashes'),
                ('unbound-copy', {'capture': '../capture.xlsm'}, 'inside the fixture directory'),
                ('false-summary', {'comparison': '2,033 outcomes.'}, 'differs from the replay'),
                ('unknown-time', {'run': {**self.run_metadata, 'seconds': -1}}, 'positive, finite'),
            )
            for name, changes, message in cases:
                with self.subTest(name=name):
                    path = self.directory / f'{name}.json'
                    benchmark.write_json(path, {**record, **changes})
                    if message:
                        with self.assertRaisesRegex(ValueError, message):
                            evidence.check(path)
                    else:
                        with contextlib.redirect_stdout(StringIO()):
                            _, capture_hash = evidence.check(path)
                        self.assertEqual(capture_hash, benchmark.digest(self.capture))
            self.assertEqual(compare.call_count, 3)

    def test_reused_capture_does_not_count_as_another_repeat(self) -> None:
        with patch.object(evidence, 'compare_capture', return_value='233 outcomes.'):
            record = evidence.evidence(self.directory, self.capture.name, self.run_metadata)
            first, second = (self.directory / f'evidence-repeat-{number}.json' for number in (1, 2))
            benchmark.write_json(first, record)
            benchmark.write_json(second, {**record, 'run': {**self.run_metadata, 'phase': 'warm'}})
            with patch.object(sys, 'argv', ['desktop_evidence.py', 'check', str(first), str(second)]):
                with contextlib.redirect_stdout(StringIO()), contextlib.redirect_stderr(StringIO()) as error:
                    with self.assertRaises(SystemExit) as failed:
                        evidence.main()
            self.assertEqual(failed.exception.code, 1)
            self.assertIn('reused capture', error.getvalue())


if __name__ == '__main__':
    unittest.main()
