"""Verify the bounded scratch overlay and guarded profile evidence."""

from pathlib import Path
import unittest

from desktop_guarded_calculation import (ENVIRONMENT, GUARDED_BRANCH, GUARDED_PROFILE, matched_results,
                                         overlay, read_profile)
from desktop_stage_profile import STAGES, instrument


RUNNER = Path(__file__).resolve().parents[1] / 'tools' / 'vba' / 'RunMarketRisk.bas'


def profile(balance: int = 1) -> str:
    """Build a complete fictional 100-pair profile with a selected retry count."""
    times = [int(bool(balance)) if stage == 'ordered_calculation' else 1 for stage in STAGES]
    return '|'.join(map(str, (GUARDED_PROFILE, 'complete', 'mac-raw', 46200, *times,
                              187, 46, 33, 3, 190, balance, 0)))


class GuardedCalculationTests(unittest.TestCase):
    def test_overlay_preserves_publication_and_the_existing_ordered_routine(self) -> None:
        source = RUNNER.read_text(encoding='ascii')
        natural = overlay(source, mode='natural')
        baseline = instrument(source)
        self.assertIn(GUARDED_BRANCH.replace('mProfileFull', 'mGuardedFirst'), natural)
        self.assertIn('    For attempt = 1 To 3\n', natural)
        self.assertIn('            If attempt = 3 Then Exit For\n', natural)
        self.assertEqual(natural.split('    If publishing Then\n')[1].split('End Sub')[0],
                         baseline.split('    If publishing Then\n')[1].split('End Sub')[0])
        self.assertEqual(natural.split('Private Sub RecalculateTrialYears')[1].split('End Sub')[0],
                         baseline.split('Private Sub RecalculateTrialYears')[1].split('End Sub')[0])
        self.assertNotIn('mGuardedCanary = mGuardedCanary + 1', natural)
        self.assertIn('If primary And code = 1 And attempt <= 2 Then', overlay(source, mode='canary'))

    def test_profile_requires_plausible_fallback_counts_and_same_batch(self) -> None:
        good = read_profile(profile(), 'GuardedCalculation/v1|46200|2|0', seconds=30, mode='natural')
        self.assertEqual(good['counters']['smart_attempts'], 190)
        self.assertEqual(good['counters']['full_fallbacks'], 2)
        self.assertNotIn('full_attempts', good['counters'])
        read_profile(profile(), 'GuardedCalculation/v1|46200|1|2', seconds=30, mode='canary')
        for counts, mode in (
                ('GuardedCalculation/v1|46201|2|0', 'natural'),
                ('GuardedCalculation/v1|46200|0|0', 'natural'),
                ('GuardedCalculation/v1|46200|191|0', 'natural'),
                ('GuardedCalculation/v1|46200|1|2', 'natural'),
                ('GuardedCalculation/v1|46200|1|1', 'canary')):
            with self.subTest(counts=counts, mode=mode), self.assertRaises(ValueError):
                read_profile(profile(), counts, seconds=30, mode=mode)
        with self.assertRaises(ValueError):
            read_profile(profile(0), 'GuardedCalculation/v1|46200|1|2', seconds=30, mode='canary')

    def test_matched_results_reject_mixed_environment_and_reused_batches(self) -> None:
        def evidence(phase: str, timestamp: float, mode: str | None) -> dict:
            return {
                'mode': mode, 'fixture_sha256': 'same fixture', 'source_sha256': 'same source',
                'reference_sha256': 'same reference', 'candidate_sha256': mode,
                'capture_sha256': str(timestamp), 'profile': {'batch_serial': timestamp},
                'comparison': 'PASS', 'run': dict.fromkeys(ENVIRONMENT, 'same environment')
                | {'phase': phase, 'repeat': 1, 'seconds': 30},
            }

        controls = [evidence('cold', 46200.1, None), evidence('warm', 46200.2, None)]
        natural = [evidence('cold', 46200.3, 'natural'), evidence('warm', 46200.4, 'natural')]
        canary = evidence('cold', 46200.5, 'canary')
        self.assertEqual(len(matched_results(controls, natural, canary)['matched_initial_observations']), 2)
        canary['run']['excel_build'] = 'different build'
        with self.assertRaisesRegex(ValueError, 'different excel_build'):
            matched_results(controls, natural, canary)
        canary['run']['excel_build'] = controls[0]['run']['excel_build']
        natural[1]['profile']['batch_serial'] = controls[0]['profile']['batch_serial']
        with self.assertRaisesRegex(ValueError, 'reused a saved capture or batch timestamp'):
            matched_results(controls, natural, canary)


if __name__ == '__main__':
    unittest.main()
