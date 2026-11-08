"""Synthetic healthcare checks, composed with the Phase 1 flat-tax fixtures."""

FPL_2027_JOINT = 21_640


def expected_credit(magi, actual, benchmark):
    """Independent ACA calculation; FPL is fixed in these zero-inflation cases."""
    ratio = magi / FPL_2027_JOINT
    if not 1 <= ratio <= 4:
        return 0
    if ratio < 1.33:
        rate = .0215
    elif ratio < 1.5:
        rate = .0323 + (ratio - 1.33) / .17 * .0107
    elif ratio < 2:
        rate = .0430 + (ratio - 1.5) / .5 * .0248
    elif ratio < 2.5:
        rate = .0678 + (ratio - 2) / .5 * .0188
    elif ratio < 3:
        rate = .0866 + (ratio - 2.5) / .5 * .0156
    else:
        rate = .1022
    return min(actual, max(0, benchmark - magi * rate))


def out(address, value):
    return {'cell': address, 'equals': value}


def case(name, inputs, *checks):
    return {'name': name, 'overrides': {'Inputs': {'D101': 1, **inputs}}, 'expects': list(checks)}


def aca(name, magi, actual=20_000, benchmark=15_000):
    return case(name, {'D102': 1, 'D106': actual, 'D116': 1, 'D119': benchmark, 'D52': magi},
                out('Healthcare!O6', FPL_2027_JOINT), out('Healthcare!AC6', magi),
                out('Healthcare!R6', expected_credit(magi, actual, benchmark)),
                out('Cash_Flow!AC6', expected_credit(magi, actual, benchmark)))


MIXED = {
    'D8': 64, 'D9': 62, 'D10': 65, 'D53': 67, 'D31': 5000, 'D52': 50000,
    'D102': 1, 'D103': 1, 'D106': 8000, 'D107': 7000, 'D108': 3000, 'D109': 3000,
    'D110': 1200, 'D111': 800, 'D112': 1000, 'D113': 1000, 'D114': 500, 'D115': 400,
    'D116': 1, 'D117': 1, 'D118': 16000, 'D119': 8000, 'D120': 10000,
    'D121': 1, 'D122': 220000, 'D123': 0,
}

NEAR_CLIFF_MAGI = 58_500 / (.8 - .1022)
NEAR_CLIFF_DRAW = NEAR_CLIFF_MAGI - 75_000

CASES = [
    case('health_mixed_medicare_and_marketplace', MIXED,
         out('Healthcare!D6', 'Medicare'), out('Healthcare!E6', 'Marketplace'),
         out('Healthcare!M6', 7000), out('Healthcare!N6', 10000), out('Healthcare!O6', FPL_2027_JOINT),
         out('Healthcare!L6', 1148.4), out('Healthcare!P6', 15248.4), out('Healthcare!Q6', 6148.4),
         out('Projection!P6', 5000), out('Healthcare!R6', expected_credit(50000, 7000, 10000)),
         out('Healthcare!Y8', 50000), out('Healthcare!L8', 0), out('Healthcare!M8', 0),
         out('Healthcare!E8', 'Medicare'), out('Healthcare!P8', 10900), out('Summary!B25', 'YES')),
    case('health_irmaa_prior_return_filing_status', {**MIXED, 'D54': 64, 'D53': 65},
         out('Projection!F6', 'Single'), out('Healthcare!Z6', 'MFJ'),
         out('Healthcare!E6', 'Not covered'), out('Healthcare!L6', 1148.4)),
    case('health_hsa_cobra_and_qualified_cost_only', {
        'D31': 10000, 'D52': 30000, 'D102': 3, 'D103': 2, 'D106': 4000, 'D107': 3000,
        'D112': 1000, 'D113': 200, 'D114': 100, 'D115': 50,
    }, out('Healthcare!D6', 'COBRA'), out('Healthcare!AB6', 4000), out('Healthcare!P6', 8350),
       out('Healthcare!Q6', 5200), out('Projection!P6', 5200), out('Projection!BG6', 4800),
       out('Healthcare!R6', 0), out('Healthcare!AF6', 3150)),
    case('health_hsa_owner_under65_excludes_spouse_medicare', {
        'D8': 63, 'D9': 64, 'D10': 64, 'D53': 64, 'D31': 5000, 'D52': 30000,
        'D102': 2, 'D106': 4000, 'D109': 3000, 'D111': 800, 'D112': 1000, 'D113': 200,
        'D121': 1, 'D122': 220000,
    }, out('Healthcare!D6', 'Other'), out('Healthcare!E6', 'Medicare'), out('Healthcare!L6', 1148.4),
       out('Healthcare!AB6', 0), out('Healthcare!P6', 10148.4), out('Projection!P6', 1200)),
    case('health_missing_medicare_lookback_blocks_decision', {**MIXED, 'D122': -1},
         out('Healthcare!T6', 'Missing IRMAA MAGI'), {'cell': 'Summary!B25', 'prefix': 'Check healthcare'}),
    case('health_missing_spouse_only_benchmark_blocks_decision', {**MIXED, 'D120': 0},
         out('Healthcare!T6', 'Missing ACA quote'), out('Healthcare!O6', 0),
         {'cell': 'Summary!B25', 'prefix': 'Check healthcare'}),
    case('health_aca_magi_adds_back_social_security_and_muni_interest', {
        'D52': 30000, 'D44': 6000, 'D45': 55, 'D46': 5000, 'D102': 1,
        'D106': 10000, 'D116': 1, 'D119': 12000,
    }, out('Healthcare!AC6', 41000), out('Healthcare!R6', expected_credit(41000, 10000, 12000))),
    aca('health_aca_below_income_floor', FPL_2027_JOINT - .01),
    aca('health_aca_at_income_floor', FPL_2027_JOINT, actual=10000),
    aca('health_aca_below_133_percent_band', FPL_2027_JOINT * 1.33 - .01),
    aca('health_aca_at_133_percent_band', FPL_2027_JOINT * 1.33),
    aca('health_aca_at_400_percent_cliff', FPL_2027_JOINT * 4),
    aca('health_aca_one_cent_above_400_percent', FPL_2027_JOINT * 4 + .01),
    case('health_aca_eligible_draw_funds_when_maximum_draw_fails', {
        'D25': 20000, 'D51': 0, 'D37': 58500, 'D52': 75000,
        'D102': 1, 'D106': 20000, 'D116': 1, 'D119': 20000,
    }, out('Projection!V6', NEAR_CLIFF_DRAW), out('Projection!AV6', .2 * NEAR_CLIFF_MAGI),
       out('Healthcare!AC6', NEAR_CLIFF_MAGI), out('Healthcare!R6', 20000 - .1022 * NEAR_CLIFF_MAGI),
       out('Projection!Z6', 0), out('Summary!B25', 'YES')),
]

NO_FEDERAL={f'B{r}':0 for r in (*range(6,13),17,18,19)}
CASES.extend([
    case('health_brokerage_loss_can_cross_into_aca_eligibility',{
        'D25':8000,'D51':16000,'D37':1000,'D52':87560,'D86':0,
        'D102':1,'D106':10000,'D116':1,'D119':20000,
    },out('Projection!V6',1000),out('Projection!AE6',-1000),out('Healthcare!AC6',86560),
      out('Healthcare!R6',10000),out('Projection!Z6',0)),
    case('health_funding_resolves_133_percent_credit_discontinuity',{
        'D25':1900,'D51':0,'D37':1100,'D52':27000,'D86':0,
        'D102':1,'D106':10000,'D116':1,'D119':10000,
    },out('Projection!V6',(1100+.0215*27000)/(1-.0215)),out('Summary!B25','YES')),
])
for fixture in CASES[-2:]: fixture['overrides']['Tax_Tables']=NO_FEDERAL
