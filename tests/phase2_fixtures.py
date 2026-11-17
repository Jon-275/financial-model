"""Generate the combined regression and independently calculated Phase 2 cases."""
import json
import re

from fixtures import COMMON, CASES as LEGACY, by_age, case, out, survives
from health_fixtures import CASES as HEALTH
from targeting_fixtures import CASES as TARGETING
from tax_fixtures import CASES as TAX

BOUNDARIES=[
    case('legacy_medical_surcharge_allows_negative_joint_magi',
         {'D8':64,'D9':62,'D10':65,'D53':67,'D25':1000,'D51':2000,'D37':1000},
         checks=(*by_age(65,AK=-1000),out('Healthcare!L8',0),out('Cash_Flow!Q8',2000),
                 out('Planning Metrics!B26',2028),out('Summary!B8','First shortfall at age 66 (2028).'),out('Planning Metrics!B25','NO - cash shortfall'))),
    case('legacy_medical_surcharge_allows_negative_single_magi',
         {'D8':64,'D9':62,'D10':65,'D53':67,'D54':64,'D25':1000,'D51':2000,'D37':1000,'D56':1},
         checks=(*by_age(65,AK=-1000),out('Healthcare!L8',0),out('Cash_Flow!Q8',2000),
                 out('Planning Metrics!B26',2028),out('Summary!B8','First shortfall at age 66 (2028).'),out('Planning Metrics!B25','NO - cash shortfall'))),
    case('complete_brokerage_liquidation_removes_its_cost_basis',
         {'D8':60,'D9':58,'D10':61,'D53':62,'D25':4246.77694910289,'D51':8493.55389820578,
          'D26':50000.123,'D37':100000,'D48':2,'D49':1000000},
         {'Asset_Allocation':{**{f'D{row}':.01 for row in range(6,11)},
                              **{f'E{row}':.02 for row in range(6,11)}}},
         checks=(*by_age(61,tolerance=0,BC=0,BD=0,BE=0),*by_age(62,tolerance=0,BC=0,BD=0),
                 out('Planning Metrics!B26',2027),out('Summary!B8','First shortfall at age 61 (2027).'),out('Planning Metrics!B25','NO - cash shortfall'))),
]

PUBLISHED=[
    case('published_2027_federal_and_2026_california_floors',{'D52':133200},
        {'Tax_Tables':{'B219':1,**{f'B{r}':v for r,v in enumerate((.10,.12,.22,.24,.32,.35,.37),6)},
                       **{f'B{r}':v for r,v in enumerate((.01,.02,.04,.06,.08,.093,.103,.113,.123),30)}}},
        checks=(*by_age(55,AL=33200,AM=100000,AN=11488,AS=121400,AT=3986.24),survives('YES'))),
    case('published_2027_payroll_cap_by_earner',{'D10':56,'D11':2,'D12':300000,'D126':.9},
        {'Tax_Tables':{'B219':1,'B50':.062,'B52':.0145,'B53':.009}},
        checks=(*by_age(55,BT=270000,BU=30000,AR=18582.60,AU=4200),survives('YES'))),
    case('opening_year_moves_calendar_and_preserves_tax_year',
         {'B130': 2027, 'D10': 56, 'D29': 2027, 'D12': 30000, 'D14': .025, 'D49': 100000},
         {'Tax_Tables': {'B219': 1}},
         checks=(out('Projection!B5', 2027), out('Projection!B6', 2028),
                 *by_age(55, G=1.025, H=1.050625, AL=34030, U=136530), survives('YES'))),
    case('joint_roth_uses_2027_income_phaseout',
         {'D10': 56, 'D12': 255000, 'D35': 20000, 'D91': .5},
         checks=(*by_age(55, M=8600), out('Roth_Access!AN6', 4300), out('Roth_Access!AO6', 4300))),
    case('joint_roth_stops_at_upper_income_limit',
         {'D10': 56, 'D12': 260000, 'D35': 6000},
         checks=(*by_age(55, M=0), out('Roth_Access!AN6', 0))),
    case('fee_reduces_account_growth',
         {'D25': 1000, 'D51': 1000, 'B131': .01},
         checks=(*by_age(55, BC=990), out('Inputs!B18', -.01))),
    case('hsa_uses_its_own_mix_in_both_stages',
         {'D10': 56, 'D15': 1, 'D16': 2, 'D27': 1000, 'D31': 1000, 'D53': 56, 'B131': .01},
         {'Asset_Allocation': {'B6': .10, 'B7': .05, 'B8': .02,
                               'D26': 1, 'D27': 0, 'D28': 0, 'D29': 0, 'D30': 0,
                               'E26': 0, 'E27': 1, 'E28': 0, 'E29': 0, 'E30': 0,
                               'D35': 1, 'D36': 0, 'D37': 0, 'D38': 0, 'D39': 0,
                               'E35': 0, 'E36': 0, 'E37': 1, 'E38': 0, 'E39': 0}},
         checks=(out('Inputs!B23', .06425), out('Inputs!B24', .0494),
                 out('Inputs!B134', .0395), out('Inputs!B135', .0098),
                 *by_age(55, BF=1089, BG=1039.5), *by_age(56, BF=1185.921, BG=1049.6871))),
    case('healthcare_budget_switches_each_spouse_at_65',
         {'D8': 64, 'D9': 62, 'D10': 65, 'D53': 65, 'D25': 20000, 'D51': 20000,
          'D42': 20000, 'D43': 10000, 'B132': .05},
         checks=(*by_age(65, BA=15750, BC=4250), survives('YES'))),
    case('survivor_benefit_tracks_combined_social_security',
         {'D44': 40000, 'D45': 55, 'D54': 54, 'D55': .60},
         checks=(*by_age(55, N=24000), out('Inputs!B55', 24000))),
    case('one_time_expense_instruction_row_requires_review', {'B83': 2031, 'C83': 5000},
         checks=(out('Inputs!B70', 0), out('Summary!B3', 'CHECK PLAN'),
                 out('Summary!C3', 'Move one-time expenses to Inputs rows 75–82'),
                 {'cell': 'Summary!C6', 'blank': True})),
    case('unavailable_workplace_roth_is_reported', {'D27': 2000, 'D28': 2, 'D37': 1000},
         checks=(out('Summary!B3', 'NO'), out('Summary!B5', 2000), out('Summary!B6', 2000))),
    case('qualified_workplace_roth_is_available', {'D8': 60, 'D9': 58, 'D10': 61, 'D53': 61,
                                                  'D27': 2000, 'D28': 2, 'D93': 2020, 'D95': 60},
         checks=(out('Summary!B3', 'YES'), out('Summary!B5', 2000), out('Summary!B6', 0))),
    case('subcent_shortfall_keeps_funded_headline', {'D37': .009},
         checks=({'cell': 'Projection!BR6', 'equals': .009, 'tolerance': 0},
                 out('Summary!B3', 'YES'), out('Summary!B4', 'None'))),
    case('exact_cent_shortfall_reaches_dashboard', {'D37': .01},
         checks=({'cell': 'Projection!BR6', 'equals': .01, 'tolerance': 0},
                 out('Summary!B3', 'NO'), out('Summary!B4', 2027))),
    case('wage_growth_includes_inflation_and_stops_at_retirement',
        {'D10':57,'D12':30000,'D13':.1,'D14':.025,'D53':57},
        {'Tax_Tables':{'B219':1}},
        checks=(*by_age(55,G=1.025,I=33000,AU=462),*by_age(56,I=36300,AU=508.2),*by_age(57,I=0,AU=0))),
    case('working_year_cost_uses_roth_ira_basis',
         {'D10':56,'D12':10000,'D36':9000,'D27':2000,'D30':1500},
         checks=(*by_age(55,Y=1000,BF=1000,Z=0),
                 out('Roth_Access!AC6',500),out('Cash_Flow!Q6',0),survives('YES'))),
    case('partial_roth_basis_leaves_visible_spending_obligation',
         {'D10':56,'D12':10000,'D36':9000,'D27':2000,'D30':500},
         checks=(*by_age(55,Y=500,BF=1500,BJ=1000,Z=500),
                 out('Cash_Flow!Q6',500),out('Cash_Flow!AU6',500),survives('NO'))),
    case('unpaid_spending_rolls_forward_without_double_counting',
         {'D37':100,'D53':57},
         checks=(out('Cash_Flow!Q6',100),out('Cash_Flow!Q7',200),out('Cash_Flow!Q8',300),
                 out('Cash_Flow!AR8',200),out('Cash_Flow!AU8',100),out('Cash_Flow!AV8',300),
                 *by_age(57,BJ=-300),survives('NO'))),
    case('later_cash_repays_prior_spending_before_deposits',
         {'D37':100,'D53':57,'D52':2000,'D87':56,'D88':56},
         checks=(out('Cash_Flow!Q6',100),out('Cash_Flow!AR7',100),
                 out('Cash_Flow!AS7',100),out('Cash_Flow!AT7',100),out('Cash_Flow!Q7',0),
                 out('Cash_Flow!AU7',0),out('Cash_Flow!AV8',100),
                 *by_age(56,BB=1400),*by_age(57,BC=1300),survives('NO'))),
    case('confirmed_qualified_workplace_roth_available_while_working',
         {'D8':60,'D9':58,'D10':62,'D53':61,'D12':2000,'D36':3000,
          'D27':2000,'D28':2,'D93':2020,'D95':60},
         checks=(*by_age(61,Y=1400,BF=600,Z=0),survives('YES'))),
]


REPORT = [
    case('report_explains_missing_retirement_age', {'D10': None},
         checks=(out('Inputs!B70', 0),
                 out('Summary!B8', 'Review the plan inputs or calculation'),
                 out('Summary!B9', 'Enter your retirement age on Inputs'),
                 {'cell': 'Summary!B12', 'blank': True})),
    case('report_explains_missing_retirement_spending', {'D37': None},
         checks=(out('Inputs!B70', 0),
                 out('Summary!B9', 'Enter retirement spending on Inputs'))),
]

def dashboard_case(existing):
    """Carry the saved-horizon and late-shortfall cases to the new report."""
    replacements = {
        'age125_matches_zero_return_simulation': (
            'age125_dashboard', out('Report_Data!A3', 125), out('Report_Data!B3', 1000),
            {'cell': 'Report_Data!A4', 'blank': True},
        ),
        'full_horizon_to125_matches_zero_return_simulation': (
            'full_horizon_to125_dashboard', out('Report_Data!A105', 125), out('Report_Data!B105', 29000),
            {'cell': 'Report_Data!A106', 'blank': True},
        ),
        'late_shortfall_reaches_summary_and_risk': (
            'late_shortfall_reaches_dashboard', {'cell': 'Summary!B3', 'prefix': 'NO'}, out('Summary!B4', 2094),
        ),
    }
    if existing['name'] not in replacements:
        return existing
    name, *assertions = replacements[existing['name']]
    obsolete = ('Monte_Carlo', 'Sequence_Risk')
    overrides = {sheet: changes for sheet, changes in existing['overrides'].items() if sheet not in obsolete}
    expects = [rule for rule in existing['expects'] if rule.get('cell', '').split('!')[0] not in obsolete]
    return {'name': name, 'overrides': overrides, 'expects': [*expects, *assertions]}


def report_addresses(existing):
    """Use Planning Metrics for legacy calculation assertions."""
    for rule in existing['expects']:
        if match := re.fullmatch(r'Summary!([ABC])(\d+)', rule.get('cell', '')):
            rule['cell'] = 'Planning Metrics!' + match[1] + match[2]
    return existing


if __name__=='__main__':
    synthetic = {
        **COMMON,
        'Inputs': {'B4': COMMON['Inputs']['B4'], 'B131': 0, 'B132': '', **COMMON['Inputs']},
        'Tax_Tables': {**COMMON['Tax_Tables'], 'B219': 0},
    }
    print(json.dumps({'common_overrides':synthetic,
                      'cases':[report_addresses(existing) for existing in
                               [dashboard_case(case) for case in LEGACY] + HEALTH + TARGETING + TAX + PUBLISHED] + BOUNDARIES + REPORT},indent=2))
