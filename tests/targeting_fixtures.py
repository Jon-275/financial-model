"""Independent one-year targeting cases, layered on the Phase 1 flat-tax fixture."""


def result(cell, value):
    return {'cell': f'Income_Targeting!{cell}', 'equals': value}


def blank(cell):
    return {'cell': f'Income_Targeting!{cell}', 'blank': True}


def case(name, inputs, request=0, age=0, tax=None, checks=()):
    overrides = {
        'Inputs': inputs,
        'Income_Targeting': {'B5': age, 'B21': request},
    }
    if tax:
        overrides['Tax_Tables'] = tax
    return {'name': name, 'overrides': overrides, 'expects': list(checks)}


FLAT_SINGLE = {**{f'B{r}': .10 for r in range(144, 151)}, 'B171': 0,
               **{f'B{r}': 0 for r in range(160, 169)}}
NO_FEDERAL = {f'B{r}': 0 for r in (*range(6, 13), 17, 18, 19)}

CASES = [
    case('targeting_equal_active_conversion_matches_all_cash_flows', {
        'D25': 10000, 'D51': 10000, 'D26': 30000, 'D37': 6000, 'D47': 5000,
    }, request=5000, checks=(
        result('B9', 55), result('B24', 5000), result('C24', 5000), result('D24', 0),
        result('B25', 1000), result('C25', 1000), result('D25', 0),
        result('B29', 7000), result('C29', 7000), result('D29', 0),
        result('B28', 0), result('C28', 0), result('B31', 'Funded'), result('C31', 'Funded'),
    )),
    case('targeting_resolves_tax_funded_draws', {
        'D25': 10000, 'D51': 10000, 'D26': 30000, 'D37': 6000,
    }, request=5000, checks=(
        result('B24', 0), result('C24', 5000), result('B29', 6000), result('C29', 7000),
        result('C25', 1000), result('D29', 1000), result('C28', 0), result('C31', 'Funded'),
    )),
    case('targeting_uses_single_filer_tax_schedule', {
        'D54': 54, 'D25': 10000, 'D51': 10000, 'D26': 30000,
    }, request=5000, tax=FLAT_SINGLE, checks=(
        result('B7', 'Single'), result('C24', 5000), result('C25', 500),
        result('C29', 500), result('C31', 'Funded'),
    )),
    case('targeting_unfunded_conversion_keeps_explicit_cash_gap', {'D26': 30000}, request=5000, checks=(
        result('B24', 0), result('C24', 5000), result('C25', 1000), result('C28', 1000),
        result('C31', 'Cash shortfall'), blank('D24'), blank('D25'),
    )),
    case('targeting_required_distribution_not_convertible', {
        'D8': 74, 'D9': 72, 'D10': 75, 'D53': 75, 'D26': 24600, 'D25': 50000, 'D51': 50000,
    }, request=24600, checks=(
        result('B11', 1000), result('B34', 23600), result('C24', 23600), result('C25', 4920),
        result('C29', 4920), result('C31', 'Request not fully funded'), result('A92', 23600),
        result('B92', 23600), result('C92', 'Funded'), blank('D24'),
    )),
    case('targeting_rejects_age_after_end_of_plan', {}, age=56, checks=(
        result('B6', 0), result('B7', 'Check retirement age'), blank('C24'), blank('A42'), blank('S42'),
    )),
    case('targeting_unavailable_with_missing_healthcare_quote', {
        'D101': 1, 'D102': 1, 'D106': 6000, 'D116': 1, 'D119': 0,
    }, request=5000, checks=(
        result('F5', 'Missing ACA quote'), result('C31', 'Check healthcare'),
        blank('C24'), blank('A42'), blank('S42'),
    )),
    case('targeting_aca_cliff_preserves_earliest_fundable_branch', {
        # $70k taxable noncash + $12k sale = $82k of MAGI, under 400% FPL
        # ($86,560 in 2027). Full $10k credit makes $12k draw feasible.
        # Adding the requested $5k conversion crosses the cliff at the
        # minimum fundable draw. Preserving the credit at the cliff leaves a
        # smaller gap than exhausting the taxable account above it.
        'D25': 21000, 'D51': 0, 'D26': 10000, 'D37': 12000, 'D52': 70000, 'D86': 0,
        'D101': 1, 'D102': 1, 'D106': 10000, 'D116': 1, 'D119': 20000,
    }, request=5000, tax=NO_FEDERAL, checks=(
        result('B26', 10000), result('B28', 0), result('B29', 12000),
        result('C24', 5000), result('C26', 10000), result('C28', 440),
        result('C29', 11560), result('C31', 'Cash shortfall'),
        result('C42', 'Funded'), result('K42', 10000), result('H42', 12000),
    )),
]
