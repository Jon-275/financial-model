"""Recalculate fictional California v2 tax fixtures and reconcile both annual models."""

from argparse import ArgumentParser, Namespace
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from fixtures import by_age, case, out
from xlsx_acceptance import run_case, scenario_source
from xlsx_read import Workbook

from desktop_policy_preflight import verify_pairing
from native_model_acceptance import validate_trial
from risk_engine.annual import AnnualModel, reconcile
from risk_engine.scenario import load_scenario
from state_policy import CALIFORNIA_POLICY, WORKBOOK_STATE_POLICY_NATIVE_SCHEMA


HISTORICAL_CASES = (
    'payroll_caps_each_joint_earner_separately',
    'payroll_single_uses_single_additional_medicare_threshold',
    'capital_loss_released_proportionally_and_used_across_years',
    'federal_and_california_opening_carryforwards_stay_separate',
    'california_senior_credit_and_phaseout_use_federal_agi',
    'published_2027_federal_and_2026_california_floors',
    'published_2027_payroll_cap_by_earner',
)
STRESS_CASE = 'california_published_surtax_at_threshold'
FIXTURES = Path(__file__).with_name('phase2-fixtures.json')


def threshold_cases() -> list[dict]:
    """Isolate the model's published California exemption credit and surtax."""
    # One senior in a joint household gets $316 + $158. Above the modeled
    # threshold, 3 exemptions lose $6 per $2,500 of excess federal AGI.
    age = {'D8': 64, 'D9': 62, 'D10': 65, 'D53': 65}
    thresholds = []
    for label, income, credit in (('below', 521536, 474), ('at', 521561, 474), ('above', 521586, 473.82)):
        thresholds.append(case(
            f'california_published_credit_{label}_threshold', {**age, 'D52': income},
            {'Tax_Tables': {'B219': 1}},
            checks=by_age(65, AK=income, CD=credit),
        ))
    # Published CA standard deduction is $11,800 in this fictional fixture.
    # The common fixture zeros ordinary CA rates, isolating the 1% surtax.
    for label, taxable, surtax in (('below', 999900, 0), ('at', 1000000, 0), ('above', 1000100, 1)):
        thresholds.append(case(
            f'california_published_surtax_{label}_threshold', {**age, 'D52': taxable + 11800},
            {'Tax_Tables': {'B219': 1, 'B44': .01, 'B45': 1000000}},
            checks=by_age(65, AS=taxable, AT=surtax, CD=0),
        ))
    return thresholds


def plan_cases() -> tuple[dict, dict[str, dict]]:
    """Preserve the checked-in goldens and activate one native trial per fixture."""
    manifest = json.loads(FIXTURES.read_text())
    historical = {item['name']: item for item in manifest['cases']}
    cases = {}
    for fixture in (*(historical[name] for name in HISTORICAL_CASES), *threshold_cases()):
        name = fixture['name']
        trial = 2201 if name == STRESS_CASE else 1
        overrides = fixture['overrides']
        cases[name] = {
            **fixture,
            'overrides': {**overrides,
                          'Risk_Trials': {**overrides.get('Risk_Trials', {}), 'B3': trial},
                          'Risk_Settings': {**overrides.get('Risk_Settings', {}), 'B31': 0, 'B32': 0}},
            'expects': [*fixture['expects'], out('Inputs!B136', 'CA'), out('Trial_Inputs!B136', 'CA'),
                        out('Market_Paths!J3', 'OK'), out('Risk_Trials!C7', 1)],
        }
    return manifest['common_overrides'], cases


def select_cases(cases: dict[str, dict], names: list[str]) -> list[dict]:
    """Select exact cases, keeping intentional execution order."""
    if missing := sorted(set(names) - cases.keys()):
        raise ValueError(f'Unknown California parity case: {", ".join(missing)}')
    if len(names) != len(set(names)):
        raise ValueError('A California parity case was repeated.')
    return [cases[name] for name in names] if names else list(cases.values())


def verify_candidate(path: Path) -> None:
    """Require an independently bootstrapped explicit v2 source before recalculation."""
    with Workbook(path) as workbook:
        if 'Risk_Data' not in workbook.parts:
            raise ValueError('Use an explicit California v2 test workbook.')
        data = workbook['Risk_Data']
        if data['Z2'].value != WORKBOOK_STATE_POLICY_NATIVE_SCHEMA or data['Z2'].formula is not None:
            raise ValueError('Use an explicit California v2 test workbook.')
        if data['Z1'].value != 'current' or data['Z6'].value != 'complete':
            raise ValueError('The California candidate needs a completed reference batch.')
    verify_pairing(path)


def verify_annual_result(path: Path) -> None:
    """Compare Calc's base and inspected trial ledgers to the independent annual engine."""
    scenario = load_scenario(path)
    if scenario.state_policy is not CALIFORNIA_POLICY:
        raise AssertionError('The recalculated scenario did not use the California policy contract.')
    model = AnnualModel(scenario)
    result = reconcile(model, model.run(model.baseline_returns, audit=True))
    print(f'PASS base plan: {result["values_compared"]:,} independent annual values. '
          f'Largest absolute difference {result["max_difference"]:.8f}.', flush=True)
    validate_trial(path, path)


def main(argv: list[str] | None = None) -> int:
    """Run each selected workbook case in an isolated Calc profile."""
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path, help='Disposable, bootstrapped explicit California v2 workbook')
    parser.add_argument('--case', action='append', default=[], help='Run only this exact case. Repeat for several.')
    parser.add_argument('--soffice', default='soffice')
    parser.add_argument('--timeout', type=int, default=180)
    options = parser.parse_args(argv)
    shared, matrix = plan_cases()
    try:
        cases = select_cases(matrix, options.case)
        verify_candidate(options.workbook)
    except ValueError as error:
        parser.error(str(error))
    with TemporaryDirectory(prefix='california-policy-parity-') as scratch:
        directory = Path(scratch)
        settings = Namespace(workbook=scenario_source(options.workbook, directory), scratch=directory,
                             compact=True, phase1=True, keep=None, soffice=options.soffice,
                             timeout=options.timeout, tolerance=.01)
        failed = 0
        for fixture in cases:
            name = fixture['name']
            try:
                run_case(settings, fixture, shared, {}, verify_result=verify_annual_result)
            except (AssertionError, ValueError, RuntimeError) as error:
                print(f'FAIL {name}: {error}', flush=True)
                failed += 1
            else:
                print(f'PASS {name}', flush=True)
    print(f'{len(cases) - failed}/{len(cases)} California v2 parity cases passed.')
    return int(bool(failed))


if __name__ == '__main__':
    raise SystemExit(main())
