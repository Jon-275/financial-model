"""Synthetic tests for the Phase 2 tax policy, in the Phase 1 runner format."""
import json
import sys
from pathlib import Path

# Works beside the Phase 1 fixtures in the repository or in the Phase 2
# integration workspace, where those fixtures are in the tests subdirectory.
workspace_tests = Path(__file__).resolve().parent / 'tests'
if (workspace_tests / 'fixtures.py').is_file():
    sys.path.insert(0, str(workspace_tests))
from fixtures import COMMON as PHASE1_COMMON, by_age, case, survives

COMMON = {
    **PHASE1_COMMON,
    'Inputs': {**PHASE1_COMMON['Inputs'], 'D101': 0, 'D126': .5, 'D127': 0, 'D128': 0},
    'Tax_Tables': {**PHASE1_COMMON['Tax_Tables'], 'B190': 3000, 'B192': 504411,
                   'B193': 252203, 'B194': 2500, 'B195': 6},
}

PAYROLL_2026 = {'B50': 0.062, 'B51': 184500, 'B52': 0.0145, 'B53': 0.009, 'B54': 250000, 'B191': 200000}
FEDERAL_2026_JOINT = {
    **{f'B{r}': rate for r, rate in enumerate((.10, .12, .22, .24, .32, .35, .37), 6)},
    **{f'C{r}': floor for r, floor in enumerate((0, 24800, 100800, 211400, 403550, 512450, 768700), 6)},
    'B17': 0, 'B18': .15, 'B19': .20, 'B23': 32200,
    'C17': 0, 'C18': 98900, 'C19': 613700,
    'B80': 140200, 'B81': 1000000, 'B82': .50, 'B83': .26, 'B84': .28, 'B85': 244500,
}

CASES = [
    case('payroll_caps_each_joint_earner_separately', {
        'D10': 56, 'D11': 2, 'D12': 300000, 'D126': .9,
    }, {'Tax_Tables': PAYROLL_2026}, checks=(
        # Primary 270k caps at 184.5k, spouse 30k does not. The household
        # pays base Medicare on 300k and additional Medicare on 50k.
        *by_age(55, BT=270000, BU=30000, AR=18099, Z=0), survives('YES'),
    )),
    case('payroll_single_uses_single_additional_medicare_threshold', {
        'D10': 56, 'D11': 2, 'D12': 300000, 'D54': 54, 'D126': .9,
    }, {'Tax_Tables': PAYROLL_2026}, checks=(
        # The single filer is one earner, with additional Medicare on 100k.
        *by_age(55, F='Single', BT=300000, BU=0, AR=16689, Z=0), survives('YES'),
    )),
    case('capital_loss_released_proportionally_and_used_across_years', {
        'D25': 25000, 'D51': 40000, 'D52': 10000, 'D53': 56,
        'A75': 'Synthetic retirement expense', 'B75': 2027, 'C75': 18600,
    }, checks=(
        # Sell 10k against 16k basis. Federal tax is $1,400 on $7k AGI in
        # both years under the common flat tax; state uses none of the
        # claimed loss because income before the loss is below its standard deduction.
        *by_age(55, V=10000, AE=-6000, BV=-3000, BW=3000, BX=-3000, BY=6000,
                CE=16000, CF=0, BD=24000, AK=7000, AV=1400, Z=0),
        *by_age(56, V=0, AE=0, BV=-3000, BW=0, BX=-3000, BY=6000,
                BD=32600, AK=7000, AV=1400, Z=0), survives('YES'),
    )),
    case('capital_loss_below_standard_deduction_carries_in_full', {
        'D25': 25000, 'D51': 40000, 'D52': 10000, 'D37': 20000,
    }, {'Tax_Tables': {'B23': 32200}}, checks=(
        *by_age(55, V=10000, AE=-6000, BV=-3000, BW=6000, BX=-3000, BY=6000,
                AK=7000, AV=0, Z=0), survives('YES'),
    )),
    case('federal_and_california_opening_carryforwards_stay_separate', {
        'D25': 20000, 'D51': 4000, 'D37': 10000, 'D127': 5000, 'D128': 12000,
    }, checks=(
        # Federal: draw - 20%*(80%*draw-5,000) = 10,000.
        # California: entire remaining net loss carries at no taxable income.
        *by_age(55, V=9000/.84, AE=.8*9000/.84, BV=.8*9000/.84-5000, BW=0,
                BX=-3000, BY=12000-.8*9000/.84, Z=0), survives('YES'),
    )),
    case('amt_preserves_preferential_gain_rates_and_compares_total_regular_tax', {
        'D25': 850000, 'D51': 0, 'D52': 350000, 'D37': 1500000,
    }, {'Tax_Tables': FEDERAL_2026_JOINT}, checks=(
        # Mandatory spending exceeds cash and available assets: sell all.
        # AMT exemption $40,200; 850k preferred plus 309.8k ordinary AMT
        # tax base. Gains stack above regular ordinary taxable of $317,800.
        *by_age(55, V=850000, AE=850000, AK=1200000, AN=61468, AO=155205,
                BZ=1159800, CA=850000, CB=237059, AQ=20386, AV=237059, Z=537059),
        survives('NO'),
    )),
    case('regular_gain_worksheet_caps_tax_at_all_ordinary_rate', {
        'D25': 500, 'D51': 0, 'D52': 132200, 'D37': 200000,
    }, {'Tax_Tables': FEDERAL_2026_JOINT}, checks=(
        # $100k regular taxable ordinary is above the 0% gain band but the
        # additional $500 is still in the 12% ordinary band. Use $60 not $75.
        *by_age(55, V=500, AE=500, AN=11504, AO=60, AQ=0, AV=11564), survives('NO'),
    )),
    case('california_senior_credit_and_phaseout_use_federal_agi', {
        'D8': 64, 'D9': 62, 'D10': 65, 'D52': 506911, 'D53': 65,
    }, checks=(
        # 2025 joint threshold $504,411. Three exemptions (two personal,
        # one senior) at $153 lose $6 each for one $2,500 increment.
        *by_age(65, AK=506911, CC=506911, CD=441), survives('YES'),
    )),
]

if __name__ == '__main__':
    print(json.dumps({'common_overrides': COMMON, 'cases': CASES}, indent=2))
