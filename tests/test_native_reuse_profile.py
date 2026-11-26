"""Check the optional reuse timing partition and its evidence boundary."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from pyopenvba.apps.excel import ExcelApplication

import desktop_reuse_profile as reuse
import desktop_stage_profile as stage
from desktop_benchmark import ROOT, write_json


STAGE_PAYLOAD = ('RunnerStageProfile/v1|complete|mac-raw|50000.125|1|1|30|3|4|2|.4|.6|1|1'
                 '|200|33|33|3|203|12|8')
REUSE_PAYLOAD = 'RunnerReuseProfile/v1|mac-raw|50000.125|.32|113|37|.08|37|33'


class ReuseProfileTests(unittest.TestCase):
    def test_parse_matches_both_calls_to_one_regular_profile(self) -> None:
        """Partition the original reuse seconds and reconcile the calls and passes."""
        profile = stage.parse(STAGE_PAYLOAD, pairs=100, platform='mac', seconds=44.1)
        self.assertEqual(reuse.parse(REUSE_PAYLOAD, pairs=100, profile=profile), {
            'fixed_path': {'seconds': .32, 'calls': 113, 'passes': 37},
            'partner_admission': {'seconds': .08, 'calls': 37, 'passes': 33},
        })
        self.assertEqual(profile['stages']['reuse_verification'], .4)

    def test_parse_rejects_invalid_receipts_and_reconciliation(self) -> None:
        """Fail closed for malformed times or counts and receipts from another batch."""
        profile = stage.parse(STAGE_PAYLOAD, pairs=100, platform='mac', seconds=44.1)
        for payload, message in (
            (REUSE_PAYLOAD + '|extra', 'fields'),
            ('RunnerReuseProfile/v1|unavailable', 'unavailable'),
            (REUSE_PAYLOAD.replace('mac-raw', 'windows-qpc'), 'clock'),
            (REUSE_PAYLOAD.replace('50000.125', '50000.126'), 'batch'),
            (REUSE_PAYLOAD.replace('50000.125', 'NaN'), 'decimal'),
            (REUSE_PAYLOAD.replace('|.32|', '|1E+999|'), 'finite'),
            (REUSE_PAYLOAD.replace('|.32|', '|-.32|'), 'nonnegative'),
            (REUSE_PAYLOAD.replace('|113|', '|113.0|'), 'integers'),
            (REUSE_PAYLOAD.replace('|113|', '|112|'), 'counts'),
            (REUSE_PAYLOAD.replace('|113|37|', '|113|114|'), 'counts'),
            (REUSE_PAYLOAD.replace('|.08|37|33', '|.08|36|33'), 'counts'),
            (REUSE_PAYLOAD.replace('|.08|37|33', '|.08|37|32'), 'counts'),
            (REUSE_PAYLOAD.replace('|.32|', '|.31|'), 'partition'),
        ):
            with self.subTest(payload=payload), self.assertRaisesRegex(ValueError, message):
                reuse.parse(payload, pairs=100, profile=profile)

    def test_zero_partner_calls_require_zero_partner_time(self) -> None:
        """A completed fixture may have no fixed paths eligible to be reused."""
        profile = stage.parse(STAGE_PAYLOAD.replace('|200|33|', '|233|0|').replace('|203|', '|236|'),
                              pairs=100, platform='mac', seconds=44.1)
        payload = 'RunnerReuseProfile/v1|mac-raw|50000.125|.4|113|0|0|0|0'
        parsed = reuse.parse(payload, pairs=100, profile=profile)
        self.assertEqual(parsed['partner_admission'], {'seconds': 0, 'calls': 0, 'passes': 0})
        with self.assertRaisesRegex(ValueError, 'counts'):
            reuse.parse(payload.replace('|.4|113|0|0|', '|.38|113|0|.02|'), pairs=100, profile=profile)

    def test_overlay_counts_exactly_the_two_existing_calls(self) -> None:
        """Keep the original decisions, clock and stage payload in the diagnostic."""
        source = (ROOT / 'tools/vba/RunMarketRisk.bas').read_text(encoding='ascii')
        regular = stage.instrument(source)
        candidate = reuse.instrument(source)
        for assignment, function, receipt in (
            ('reused', 'CanReuseFixedTrial(book, batch, ids, pairedRow, code, context, width)',
             'RecordReuseCheck ReusePartnerAdmission, reused'),
            ('fixedUnchanged(index)', 'FixedPathAllowsReuse(cutProof, captured, book)',
             'RecordReuseCheck ReuseFixedPath, fixedUnchanged(index)'),
        ):
            call = f'                {assignment} = {function}\n'
            anchor = '                MoveMarketRiskProfile ProfileReuseVerification\n' + call
            self.assertEqual(candidate.count(anchor + f'                {receipt}\n'), 1)
        for name in ('RunnerClockTicks', 'ReadMarketRiskProfile'):
            anchor = f'Function {name}('
            start = regular.index(anchor)
            definition = regular[start:regular.index('End Function', start)]
            self.assertIn(definition, candidate)
        self.assertNotIn('ReadReuseVerificationProfile', regular)
        self.assertIn('If mRunning Or mProfileOutcome <> "complete" Then', candidate)
        self.assertIn('Erase mReuseCalls', candidate)
        self.assertEqual(candidate.count('.Calculate'), regular.count('.Calculate'))
        self.assertEqual(candidate.count('DoEvents'), regular.count('DoEvents'))

    def test_macro_partitions_existing_clock_intervals_and_counts_success(self) -> None:
        """Execute the new VBA procedures with a deterministic stage clock."""
        app = ExcelApplication()
        harness = '''
Private Const ProfileReuseVerification As Long = 1
Private Const ProfilePreparation As Long = 2
Private Const PROFILE_CLOCK As String = "mac-raw"
Private mProfileActive As Long, mProfileSeconds(1 To 2) As Double
Private mProfileOutcome As String, mProfileSerial As Double, mRunning As Boolean
Private Sub MoveMarketRiskProfile(ByVal targetStage As Long)
    mProfileSeconds(mProfileActive) = mProfileSeconds(mProfileActive) + 0.125
    mProfileActive = targetStage
End Sub
Public Sub SimulateReuseProfile()
    mProfileActive = ProfileReuseVerification
    RecordReuseCheck ReuseFixedPath, True
    mProfileActive = ProfileReuseVerification
    RecordReuseCheck ReuseFixedPath, False
    mProfileActive = ProfileReuseVerification
    RecordReuseCheck ReusePartnerAdmission, True
    mProfileOutcome = "complete"
    mProfileSerial = 50000.125
End Sub
'''
        app.add_module(reuse.DECLARATIONS + harness + reuse.PROCEDURES, name='ReuseTest')
        self.assertEqual(app.run('ReadReuseVerificationProfile'), f'{reuse.PROFILE}|unavailable')
        app.run('SimulateReuseProfile')
        fields = app.run('ReadReuseVerificationProfile').split('|')
        self.assertEqual(fields[:3], [reuse.PROFILE, 'mac-raw', '50000.125'])
        self.assertEqual(list(map(float, fields[3:])), [.25, 2, 1, .125, 1, 1])

    def test_companion_evidence_replays_with_the_extra_runner(self) -> None:
        """Use existing independent reconciliation with this candidate's exact VBA."""
        profile = stage.parse(STAGE_PAYLOAD, pairs=100, platform='mac', seconds=44.1)
        saved_stage = {'schema': stage.EVIDENCE, 'pairs': 100, 'profile': profile, 'payload': STAGE_PAYLOAD,
                       'run': {'seconds': 44.1}, 'capture': 'saved.xlsm'}
        with TemporaryDirectory() as temp, patch.object(stage, 'record', return_value=saved_stage) as record:
            directory = Path(temp)
            evidence = reuse.record(directory, 'saved.xlsm', saved_stage['run'], STAGE_PAYLOAD, REUSE_PAYLOAD)
            self.assertEqual(evidence['stage'], saved_stage)
            self.assertIs(record.call_args.kwargs['transform'], reuse.instrument)
            valid = directory / 'valid.json'
            write_json(valid, evidence)
            self.assertEqual(reuse.check(valid), evidence)
            invalid = directory / 'invalid.json'
            write_json(invalid, {**evidence, 'reuse': {}})
            with self.assertRaisesRegex(ValueError, 'replay'):
                reuse.check(invalid)
