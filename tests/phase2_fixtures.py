"""Generate the combined regression and independently calculated Phase 2 cases."""
import json

from fixtures import COMMON, CASES as LEGACY, by_age, case, out, survives
from health_fixtures import CASES as HEALTH
from targeting_fixtures import CASES as TARGETING
from tax_fixtures import CASES as TAX

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


if __name__=='__main__':
    synthetic = {
        **COMMON,
        'Inputs': {'B4': COMMON['Inputs']['B4'], 'B131': 0, 'B132': '', **COMMON['Inputs']},
        'Tax_Tables': {**COMMON['Tax_Tables'], 'B219': 0},
    }
    print(json.dumps({'common_overrides':synthetic,
                      'cases':[dashboard_case(existing) for existing in LEGACY]+HEALTH+TARGETING+TAX+PUBLISHED},indent=2))
