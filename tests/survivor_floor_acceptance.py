"""Check that the survivor medical stress and survivor Social Security use separate inputs."""
import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from fixtures import COMMON
from native_capture_acceptance import prepare_trial_copy
from risk_engine.native_annual import NativeAnnualModel, price_path
from risk_engine.scenario import load_scenario
from workbook_recalc import recalculate
from xlsx_acceptance import prepare_xlsx
from xlsx_read import Workbook


def equal(label, actual, expected):
    if actual != expected:
        raise AssertionError(f'{label}: expected {expected!r}, found {actual!r}')


def verify(source, office):
    with TemporaryDirectory() as tmp:
        root = Path(tmp)
        clean = root / ('clean' + source.suffix)
        prepare_trial_copy(source, clean)
        for case, (combined, medical_share) in enumerate(((20000, .75), (0, .60))):
            folder = root / str(case)
            output = folder / 'out'
            output.mkdir(parents=True)
            ssa = .60 * combined
            gross = 18000 * medical_share
            inputs = {**COMMON['Inputs'], 'D8': 66, 'D9': 66, 'D10': 67, 'D53': 68, 'D54': 66,
                      'D43': 4000, 'D44': combined, 'D45': 67, 'D55': .60, 'D57': medical_share,
                      'D101': 0, 'B132': 0}
            changes = {'Inputs': inputs, 'Risk_Trials': {'B3': 2206}, 'Risk_Settings': {'B29': 1},
                       'Risk_Data': {'Z6': 'running'}}
            prepared = folder / ('test' + source.suffix)
            prepare_xlsx(clean, prepared, changes, compact=True)
            calculated = recalculate(prepared, output, folder / 'profile', office, 600)
            scenario = load_scenario(calculated)
            period = price_path(NativeAnnualModel(scenario), (0, 0), end_age=68, medical_floor=18000)
            path = period.run(period.baseline_returns, audit=True)
            equal('Python annual survivor benefit in dollars', scenario.inputs[55], ssa)
            equal('Python survivor healthcare share', scenario.inputs[57], medical_share)
            with Workbook(calculated) as book:
                inputs, health = book['Inputs'], book['Trial_Healthcare']
                equal('Selected trial', book['Risk_Trials']['B3'].value, 2206)
                equal('Valid trial', book['Risk_Trials']['C7'].value, 1)
                for k, expected in {'D55': .60, 'B55': ssa, 'D57': medical_share, 'B57': medical_share}.items():
                    equal(f'Inputs!{k}', inputs[k].value, expected)
                for row, year in enumerate(path.years, 6):
                    equal('Survivor filing', book['Trial_Projection'][f'F{row}'].value, 'Single')
                    equal('Social Security received', book['Trial_Projection'][f'N{row}'].value, ssa)
                    equal('Medical cost before reserve', health[f'AG{row}'].value, 4000 * medical_share)
                    equal('Medical cash reserve', health[f'AH{row}'].value, (18000 - 4000) * medical_share)
                    equal('Gross medical costs', health[f'P{row}'].value, gross)
                    equal('Python gross medical costs', year.audit['Healthcare']['P'], gross)
            print(f'PASS ages 67–68: Social Security {ssa:,.0f}, survivor healthcare share {medical_share:.0%}, '
                  f'gross medical costs {gross:,.0f} in Excel formulas and Python.', flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--soffice', default='soffice')
    args = parser.parse_args()
    verify(args.workbook, args.soffice)
