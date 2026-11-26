"""Verify the bounded scratch overlay and guarded profile evidence."""

from pathlib import Path
import unittest

from desktop_guarded_calculation import (ENVIRONMENT, GUARDED_BRANCH, GUARDED_PROFILE, MACRO_WALL_TIME,
                                         MACRO_WALL_TIME_LABELS, matched_results, matched_warm_results,
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

    def test_profile_counts_reconcile_to_the_declared_full_scale_fixture(self) -> None:
        times = [1] * len(STAGES)
        payload = '|'.join(map(str, (GUARDED_PROFILE, 'complete', 'mac-raw', 46200, *times,
                                     1532, 501, 33, 3, 1535, 1, 0)))
        counts = 'GuardedCalculation/v1|46200|4|0'
        result = read_profile(payload, counts, seconds=30, mode='natural', pairs=1000)
        self.assertEqual(result['counters']['smart_attempts'], 1535)
        with self.assertRaisesRegex(ValueError, 'counts do not reconcile'):
            read_profile(payload, counts, seconds=30, mode='natural', pairs=100)

    def test_matched_results_reject_mixed_environment_and_reused_batches(self) -> None:
        def evidence(phase: str, timestamp: float, mode: str | None) -> dict:
            method = MACRO_WALL_TIME if mode is None else MACRO_WALL_TIME_LABELS[1]
            return {
                'fixture': 'full-0-125', 'pairs': 1000,
                'mode': mode, 'fixture_sha256': 'same fixture', 'source_sha256': 'same source',
                'reference_sha256': 'same reference', 'candidate_sha256': mode,
                'capture_sha256': str(timestamp), 'profile': {'batch_serial': timestamp},
                'comparison': 'PASS', 'run': dict.fromkeys(ENVIRONMENT, 'same environment')
                | {'phase': phase, 'repeat': 1, 'seconds': 30, 'timing_method': method},
            }

        controls = [evidence('cold', 46200.1, None), evidence('warm', 46200.2, None)]
        natural = [evidence('cold', 46200.3, 'natural'), evidence('warm', 46200.4, 'natural')]
        canary = evidence('cold', 46200.5, 'canary')
        result = matched_results(controls, natural, canary)
        self.assertEqual(result['experiment'], 'full-0-125, 1000 fictional pairs')
        self.assertEqual(len(result['matched_initial_observations']), 2)
        self.assertEqual(result['environment']['timing_method'], MACRO_WALL_TIME)
        controls[0]['run']['timing_method'] = 'monotonic time through file save'
        with self.assertRaisesRegex(ValueError, 'macro wall time through return'):
            matched_results(controls, natural, canary)
        controls[0]['run']['timing_method'] = MACRO_WALL_TIME
        canary['run']['excel_build'] = 'different build'
        with self.assertRaisesRegex(ValueError, 'different excel_build'):
            matched_results(controls, natural, canary)
        canary['run']['excel_build'] = controls[0]['run']['excel_build']
        natural[1]['profile']['batch_serial'] = controls[0]['profile']['batch_serial']
        with self.assertRaisesRegex(ValueError, 'reused a saved capture or batch timestamp'):
            matched_results(controls, natural, canary)

    def test_warm_medians_require_three_distinct_replays_in_each_original_session(self) -> None:
        """Require exact timing, the preceding capture and the original Excel session."""
        def arm(name: str, pid: str, start: int, times: list[int]) -> list[tuple[dict, dict]]:
            """Create evidence and timing receipts for a fictional Excel process."""
            result = []
            prior = f'{name}-input'
            for index, seconds in enumerate(times):
                timing = dict.fromkeys(ENVIRONMENT, 'same') | {
                    'platform': 'mac', 'excel_bitness': 64, 'phase': 'warm' if index else 'cold',
                    'repeat': index or 1, 'seconds': seconds, 'timing_method': MACRO_WALL_TIME,
                    'distribution': 'test fixture',
                }
                capture = f'{name}-capture-{index}'
                evidence = {
                    'fixture': 'full-0-125', 'pairs': 1000, 'fixture_sha256': 'fixture',
                    'source_sha256': 'source', 'reference_sha256': 'reference',
                    'candidate_sha256': f'{name}-input', 'capture_sha256': capture,
                    'comparison': 'PASS', 'profile': {'batch_serial': start + index}, 'run': timing,
                }
                if name == 'control':
                    evidence['payload'] = 'control profile'
                else:
                    evidence.update(mode='natural', profile_text='candidate profile', count_text='fallback count')
                receipt = {
                    'excel_pid': pid, 'starting_sha256': prior, 'timing': dict(timing),
                    'profile': evidence.get('payload') or evidence['profile_text'],
                    'counts': evidence.get('count_text', ''),
                }
                result.append((evidence, receipt))
                prior = capture
            return result

        control = arm('control', '110', 46200, [1000, 120, 98, 100])
        candidate = arm('candidate', '210', 46210, [999, 75, 86, 72])
        result = matched_warm_results(control, candidate)
        self.assertEqual(result['warm_median_seconds'], {'control': 100, 'guarded': 75})
        self.assertEqual(result['warm_improvement_percent'], 25)
        self.assertEqual(result['windows_status'], 'unverified')
        with self.assertRaisesRegex(ValueError, 'three warm'):
            matched_warm_results(control[:-1], candidate)
        candidate[2][1]['excel_pid'] = '211'
        with self.assertRaisesRegex(ValueError, 'same Excel process'):
            matched_warm_results(control, candidate)
        candidate[2][1]['excel_pid'] = '210'
        candidate[2][1]['starting_sha256'] = candidate[0][0]['capture_sha256']
        with self.assertRaisesRegex(ValueError, 'preceding saved capture'):
            matched_warm_results(control, candidate)
        candidate[2][1]['starting_sha256'] = candidate[1][0]['capture_sha256']
        candidate[2][0]['pairs'] = 100
        with self.assertRaisesRegex(ValueError, 'single allowed pair count'):
            matched_warm_results(control, candidate)
        candidate[2][0]['pairs'] = 1000
        candidate[2][1]['timing']['seconds'] = 70
        with self.assertRaisesRegex(ValueError, 'timing receipt'):
            matched_warm_results(control, candidate)


if __name__ == '__main__':
    unittest.main()
