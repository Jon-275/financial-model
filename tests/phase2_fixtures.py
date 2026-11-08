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
    synthetic={**COMMON,'Tax_Tables':{**COMMON['Tax_Tables'],'B219':0}}
    print(json.dumps({'common_overrides':synthetic,
                      'cases':[dashboard_case(existing) for existing in LEGACY]+HEALTH+TARGETING+TAX+PUBLISHED},indent=2))
