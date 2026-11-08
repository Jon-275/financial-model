"""Generate synthetic spreadsheet fixtures with independently calculated outputs.

The 20% flat tax and zero return rates are test assumptions, not tax guidance.
The only inputs supplied to the test runner are model scenario and tax assumptions.
"""

import json
from pathlib import Path


COMMON = {
    "Inputs": {
        "B4": 1, "D8": 54, "D9": 52, "D10": 55, "D11": 1, "D12": 0, "D13": 0, "D14": 0,
        "D25": 0, "D26": 0, "D27": 0, "D28": 1, "D29": 2020, "D30": 0, "D31": 0,
        "D32": 0, "D33": 0, "D34": 0, "D35": 0, "D36": 0, "D37": 0, "D38": 0, "D39": 0,
        "D42": 0, "D43": 0, "D44": 0, "D46": 0, "D47": 0, "D48": 1, "D50": 55,
        "D51": 0, "D52": 0, "D53": 55, "D54": 999, "D58": 999, "D67": 75, "D68": 0,
        "D86": 1, "D87": 0, "D88": 125, "D89": 1, "D90": 1, "D91": 1, "D92": 0,
        "D93": 0, "D94": 0, "D95": 126, "D96": 126, "D97": 1, "D98": 1,
    },
    "Asset_Allocation": {f"{col}{r}": 0 for col in "BCDE" for r in range(6, 11)},
    "Tax_Tables": {
        **{f"B{r}": 0.2 for r in (*range(6, 13), 17, 18, 19)},
        **{f"B{r}": 0 for r in (*range(30, 39), 23, 24, 44, 46, 50, 52, 53, 74, 83, 84)},
    },
}


def by_age(age, **expected):
    return [{"age": age, "column": column, "equals": amount} for column, amount in expected.items()]


def out(cell, value):
    return {"cell": cell, "equals": value}


def survives(answer):
    return {"cell": "Summary!B25", "prefix": answer}


def case(name, inputs=None, other=None, checks=()):
    return {"name": name, "overrides": {"Inputs": inputs or {}, **(other or {})}, "expects": list(checks)}


CASES = [
    case("cash_other_income_while_working", {"D10": 56, "D12": 10000, "D35": 6000, "D36": 22000, "D52": 30000}, checks=(
        *by_age(55, O=30000, AV=8000, M=6000, BB=4000, Z=0, BO=10000), survives("YES"),
    )),
    case("noncash_taxable_income", {"D10": 56, "D12": 10000, "D35": 6000, "D36": 22000, "D52": 30000, "D86": 0}, checks=(
        *by_age(55, O=30000, AV=8000, M=0, BB=0, Z=20000, BO=0), survives("NO"),
    )),
    case("working_cash_deficit", {"D10": 56, "D12": 10000, "D35": 6000, "D36": 9000}, checks=(
        *by_age(55, AV=2000, M=0, BB=0, Z=1000, BF=0), survives("NO"),
    )),
    case("working_brokerage_draw_funds_required_cash", {"D10": 56, "D12": 10000, "D25": 3000, "D35": 6000, "D36": 9000, "D51": 3000}, checks=(
        *by_age(55, V=1000, AV=2000, M=0, BB=0, Z=0, BC=2000, BF=0), survives("YES"),
    )),
    case("ira_savings_limited_by_taxable_wages", {"D10": 56, "D12": 2000, "D32": 1000, "D35": 5000, "D52": 50000}, checks=(
        *by_age(55, J=1000, AV=10200, M=1000, BB=39800, BF=1000, Z=0), survives("YES"),
    )),
    case("income_window_includes_boundaries", {"D10": 57, "D52": 30000, "D53": 58, "D87": 56, "D88": 57}, checks=(
        *by_age(55, O=0, AV=0, BB=0), *by_age(56, O=30000, AV=6000, BB=24000),
        *by_age(57, O=30000, AV=6000, BB=24000), *by_age(58, O=0, AV=0, BB=0, BC=48000), survives("YES"),
    )),
    case("same_year_grossup_first_and_final_retirement", {"D8": 60, "D9": 58, "D10": 61, "D26": 30000, "D37": 10000, "D53": 62}, checks=(
        *by_age(61, W=12500, AV=2500, Z=0, BE=17500), *by_age(62, W=12500, AV=2500, Z=0, BE=5000), survives("YES"),
    )),
    case("grossup_hits_account_limit", {"D8": 60, "D9": 58, "D10": 61, "D26": 12000, "D37": 10000, "D53": 61}, checks=(
        *by_age(61, W=12000, AV=2400, Z=400, BE=0), survives("NO"),
    )),
    case("grossup_crosses_progressive_bracket", {"D8": 60, "D9": 58, "D10": 61, "D26": 30000, "D37": 11000, "D53": 61}, {
        "Tax_Tables": {"B6": .1, **{f"B{r}": .3 for r in range(7, 13)}, "C7": 10000},
    }, checks=(
        # Above $10,000: tax = $1,000 + 30% of the excess. Then 70% of draw + $2,000 must equal $11,000.
        *by_age(61, W=9000 / .7, AV=9000 / .7 - 11000, Z=0), survives("YES"),
    )),
    case("realized_gain_grossup", {"D25": 25000, "D51": 12500, "D37": 10000}, checks=(
        *by_age(55, V=10000 / .9, AE=5000 / .9, AV=1000 / .9, BD=12500 - 5000 / .9, Z=0), survives("YES"),
    )),
    case("tax_exempt_income_cash_share", {"D37": 3000, "D46": 10000, "D89": .5}, checks=(
        *by_age(55, AG=10000, AV=0, Z=0, BB=2000), survives("YES"),
    )),
    case("hsa_covers_eligible_healthcare_once", {"D25": 2000, "D31": 3000, "D42": 4000, "D51": 2000}, checks=(
        *by_age(55, BA=4000, P=3000, V=1000, AV=0, Z=0, BC=1000, BG=0, BO=-4000), survives("YES"),
    )),
    case("prior_tax_arrears_paid_before_deposit", {"D8": 60, "D9": 58, "D10": 61, "D52": 30000, "D53": 62, "D86": 0, "D87": 61, "D88": 61, "D44": 12000, "D45": 62}, checks=(
        *by_age(61, AV=6000, Z=6000, BO=0), *by_age(62, AV=0, Z=0, BB=6000, BC=6000, BO=6000), survives("NO"),
    )),
    case("exact_cash_funding", {"D25": 1000, "D51": 1000, "D37": 1000}, checks=(
        *by_age(55, V=1000, AV=0, Z=0, BC=0), survives("YES"),
    )),
    case("subdollar_shortfall_is_visible", {"D25": 1000, "D51": 1000, "D37": 1000.5}, checks=(
        *by_age(55, V=1000, AV=0, Z=.5, BC=0), survives("NO"),
    )),
    case("workplace_roth_is_not_ira_basis", {"D27": 40000, "D28": 2, "D30": 10000, "D37": 15000}, checks=(
        *by_age(55, X=0, Y=0, Z=15000, BF=40000), out("Summary!B19", 0), survives("NO"),
    )),
    case("workplace_access_at55_is_not_qualified", {"D27": 40000, "D28": 2, "D37": 15000, "D93": 2020, "D95": 55}, checks=(
        *by_age(55, X=0, Y=0, Z=15000, BF=40000), out("Summary!B19", 0), survives("NO"),
    )),
    case("verified_workplace_roth_qualifies_at60", {"D8": 59, "D9": 57, "D10": 60, "D27": 40000, "D28": 2, "D37": 10000, "D53": 60, "D93": 2020, "D95": 60}, checks=(
        *by_age(60, X=40000, Y=10000, AV=0, Z=0, BF=30000), out("Roth_Access!W6", 10000), survives("YES"),
    )),
    case("ira_regular_basis_excludes_earnings", {"D27": 40000, "D28": 1, "D30": 10000, "D37": 15000}, checks=(
        *by_age(55, X=10000, Y=10000, Z=5000, BF=30000), out("Summary!B19", 10000), survives("NO"),
    )),
    case("roth_owners_have_independent_eligibility", {"D8": 59, "D9": 57, "D10": 60, "D27": 40000, "D30": 0, "D37": 25000, "D53": 60, "D90": .5, "D92": 2025}, checks=(
        *by_age(60, X=20000, Y=20000, Z=5000, BF=20000), {"cell": "Summary!B19", "blank": True}, survives("NO"),
    )),
    case("primary_only_roth_share_does_not_spill", {"D9": 44, "D10": 56, "D12": 50000, "D35": 20000, "D91": 1}, checks=(
        *by_age(55, AV=10000, M=8600, BB=31400, Z=0), out("Roth_Access!Y6", 8600), out("Roth_Access!Z6", 0), survives("YES"),
    )),
    case("roth_split_respects_the_lower_individual_limit", {"D9": 44, "D10": 56, "D12": 50000, "D35": 20000, "D91": .5}, checks=(
        *by_age(55, AV=10000, M=15000, BB=25000, Z=0), out("Roth_Access!Y6", 7500), out("Roth_Access!Z6", 7500), survives("YES"),
    )),
    case("conversion_is_internal_and_uses_fifth_tax_year", {
        "D8": 49, "D9": 47, "D10": 50, "D25": 2000, "D26": 30000, "D47": 10000, "D50": 50,
        "D51": 2000, "D53": 55, "A75": "Spending just before maturation", "B75": 2031, "C75": 6000,
        "A76": "Spending at maturation", "B76": 2032, "C76": 6000,
    }, checks=(
        *by_age(50, V=2000, AA=10000, AV=2000, Z=0, BE=20000, BF=10000, BO=-2000),
        *by_age(54, X=0, Y=0, Z=6000, BF=10000), *by_age(55, X=10000, Y=6000, Z=0, BF=4000), survives("NO"),
    )),
    case("optional_conversion_does_not_create_unpayable_tax", {"D26": 30000, "D47": 10000}, checks=(
        # No brokerage, income, or unlocked distributions can pay the tax; skip the optional conversion.
        *by_age(55, W=0, AA=0, AV=0, Z=0, BE=30000, BF=0, BO=0), survives("YES"),
    )),
    case("rmd_reinvests_only_net_cash", {"D8": 74, "D9": 72, "D10": 75, "D26": 88000, "D53": 75}, checks=(
        *by_age(75, W=88000 / 24.6, AV=17600 / 24.6, BB=70400 / 24.6, BO=-17600 / 24.6, Z=0), survives("YES"),
    )),
    case("rmd_uses_spouse_owners_age", {"D8": 76, "D9": 73, "D10": 77, "D26": 88000, "D53": 78, "D97": 2}, checks=(
        *by_age(77, D=74, S=0, W=0, BE=88000),
        *by_age(78, D=75, S=88000 / 24.6, W=88000 / 24.6, AV=17600 / 24.6, BB=70400 / 24.6, BO=-17600 / 24.6, Z=0), survives("YES"),
    )),
    case("age125_matches_zero_return_simulation", {"D8": 124, "D9": 122, "D10": 125, "D25": 2000, "D51": 2000, "D37": 1000, "D53": 125}, {
        "Sequence_Risk": {"B5": 0, "B7": 0},
    }, checks=(
        *by_age(125, BH=1000), out("Monte_Carlo!L26", 1000), out("Monte_Carlo!L27", None), out("Sequence_Risk!G16", 1000),
        out("Inputs!B70", 2), out("Projection!C7", None), survives("YES"),
    )),
    case("full_horizon_to125_matches_zero_return_simulation", {"D8": 22, "D9": 20, "D25": 100000, "D51": 100000, "D37": 1000, "D53": 125}, {
        "Sequence_Risk": {"B5": 0, "B7": 0},
    }, checks=(
        *by_age(54, BH=100000), *by_age(55, BH=99000), *by_age(125, BH=29000),
        out("Monte_Carlo!L128", 29000), out("Sequence_Risk!G118", 29000),
        out("Inputs!B70", 104), out("Projection!C109", None), survives("YES"),
    )),
    case("late_shortfall_reaches_summary_and_risk", {
        "D8": 22, "D9": 20, "D25": 100000, "D51": 100000, "D37": 1000, "D53": 95,
        "A75": "Late required outflow", "B75": 2094, "C75": 65000,
    }, checks=(
        *by_age(89, BH=65000, Z=0), *by_age(90, BH=0, Z=1000), *by_age(95, BH=0, Z=1000),
        out("Summary!B26", 2094), out("Summary!B27", 1000), out("Monte_Carlo!B15", "Base plan unfunded"),
        out("Sequence_Risk!B148", "Base plan unfunded"), survives("NO"),
    )),
    case("invalid_age126", {"D53": 126}, checks=(out("Inputs!B70", 0), {"cell": "Projection!C5", "blank": True})),
    case("scenario_selector_uses_active_cash_controls", {
        **{"E" + address[1:]: value for address, value in COMMON["Inputs"].items() if address.startswith("D")},
        "B4": 2, "E10": 56, "E12": 10000, "E36": 10000, "E52": 30000, "E86": .5,
    }, checks=(out("Inputs!B86", .5), *by_age(55, O=30000, AV=8000, BB=7000, Z=0), survives("YES"))),
]


if __name__ == "__main__":
    print(json.dumps({"common_overrides": COMMON, "cases": CASES}, indent=2))
