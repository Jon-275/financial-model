"""Verify independent accounting identities against the saved workbook."""

from __future__ import annotations

import argparse
import math
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from xlsx_read import Workbook, col_letter, q, sheet_parts


# Match the workbook's funding decision independently of comparison tolerance.
SHORTFALL_THRESHOLD = 0.01


REQUIRED_CASH_HEADERS = {
    "cash income", "required spending", "hsa paid", "opening tax payable", "current tax", "tax paid",
    "ending tax payable", "spending paid", "unpaid spending", "roth contribution", "taxable deposit",
    "cash residual", "essential shortfall",
}


def norm(value: str) -> str:
    return " ".join(value.lower().split())


def sheet_columns(sheet, row: int, expected: set[str]) -> dict[str, str]:
    if row not in sheet.rows:
        raise AssertionError(f"No header row at {sheet.title}!{row}")
    actual = {}
    for column, cell in sheet.rows[row].items():
        if isinstance(cell.value, str):
            header = norm(cell.value)
            if header in expected and header in actual:
                raise AssertionError(f"Duplicate header {cell.value!r} in {sheet.title}!{row}")
            actual[header] = col_letter(column)
    if missing := expected - actual.keys():
        raise AssertionError(f"{sheet.title} is missing headers: {', '.join(sorted(missing))}")
    return actual


def expect_number(sheet, coordinate: str) -> float:
    value = sheet[coordinate].value
    if not isinstance(value, (int, float)) or isinstance(value, bool) or not math.isfinite(value):
        raise AssertionError(f"{sheet.title}!{coordinate}: expected a finite number, got {value!r}")
    return value


def equal(label: str, actual: float, expected: float, tolerance: float) -> None:
    # Dollar tolerance is absolute; scale only enough to allow binary floating-point roundoff.
    allowance = tolerance + 32 * max(math.ulp(float(actual)), math.ulp(float(expected)))
    if abs(actual - expected) > allowance:
        raise AssertionError(f"{label}: expected {expected:.6f}, found {actual:.6f} (difference {actual-expected:+.6f})")


def require_blank(sheet, coordinates) -> None:
    for coordinate in coordinates:
        if sheet[coordinate].value not in (None, ""):
            raise AssertionError(f"{sheet.title}!{coordinate}: result beyond plan horizon: {sheet[coordinate].value!r}")


def check_chart_structure(workbook: Path) -> None:
    # The cached charts should still be present; this is not a visual Excel check.
    with ZipFile(workbook) as archive:
        modern = "Risk_Settings" in sheet_parts(archive)
        charts = sorted(name for name in archive.namelist() if name.startswith("xl/charts/chart") and name.endswith(".xml"))
        expected = 4 if modern else 5
        if len(charts) != expected:
            raise AssertionError(f"Expected {expected} chart objects, found {len(charts)}")
        for path in charts:
            root = ET.fromstring(archive.read(path))
            for element in root.iter():
                if not element.tag.endswith("}f"):
                    continue
                formula = element.text or ""
                if "#REF!" in formula or (modern and any(
                    token in formula for token in ("Monte_Carlo", "Sequence_Risk", "RMTrials")
                )):
                    raise AssertionError(f"Broken chart formula in {path}")


def check_probe_formula_encoding(workbook: Path) -> None:
    # Excel for Mac drops shared formulas with cross-sheet references on open.
    with ZipFile(workbook) as archive:
        parts = sheet_parts(archive)
        for name in ('Funding_Calc', 'Targeting_Calc'):
            if name not in parts:
                continue
            with archive.open(parts[name]) as source:
                for _, element in ET.iterparse(source, events=('end',)):
                    if element.tag == q('f') and element.get('t') == 'shared' and '!' in (element.text or ''):
                        raise AssertionError(f'{name} contains a shared formula with a cross-sheet reference')
                    element.clear()


def premium_credit(year: int, magi: float, poverty: float, actual: float, benchmark: float) -> float:
    if min(poverty, actual, benchmark) <= 0 or not poverty <= magi <= 4 * poverty:
        return 0.0
    ratio = magi / poverty
    bands = (((1, 1.33, .0215, .0215), (1.33, 1.5, .0323, .0430), (1.5, 2, .0430, .0678),
              (2, 2.5, .0678, .0866), (2.5, 3, .0866, .1022), (3, 4, .1022, .1022)) if year >= 2027 else
             ((1, 1.33, .021, .021), (1.33, 1.5, .0314, .0419), (1.5, 2, .0419, .066),
              (2, 2.5, .066, .0844), (2.5, 3, .0844, .0996), (3, 4, .0996, .0996)))
    left, right, low, high = next(x for x in bands if ratio < x[1] or x[1] == 4)
    contribution = magi * (low + (ratio - left) * (high - low) / (right - left))
    return min(actual, max(0, benchmark - contribution))


def validate_model(wb: Workbook, tolerance: float = 0.01) -> dict:
    modern = "Risk_Settings" in wb.parts
    legacy_risk = ("Monte_Carlo", "Sequence_Risk")
    risk_sheets = ("Risk_Settings", "Report_Data") if modern else legacy_risk
    for name in ("Inputs", "Projection", "Cash_Flow", "Roth_Access", "Funding_Calc", "Checks", "Summary", *risk_sheets):
        if name not in wb.parts:
            raise AssertionError(f"Missing worksheet: {name}")
    if modern:
        if any(name in wb.parts for name in legacy_risk):
            raise AssertionError("Retired risk worksheets should be absent")
        from dashboard_acceptance import validate_dashboard
        validate_dashboard(wb, tolerance)

    inputs, projection, cash, summary, roth, checks = (wb[name] for name in ("Inputs", "Projection", "Cash_Flow", "Summary", "Roth_Access", "Checks"))
    cols = sheet_columns(cash, 4, REQUIRED_CASH_HEADERS)
    has_health = "Healthcare" in wb.parts
    health = wb['Healthcare'] if has_health else None
    roth_headers = {"roth accessible", "roth closing", "primary ira opening", "spouse ira opening", "primary workplace opening", "spouse workplace opening", "primary ira draw", "spouse ira draw", "primary workplace draw", "spouse workplace draw", "primary savings", "spouse savings", "primary conversion", "spouse conversion"}
    roth_cols = sheet_columns(roth, 4, roth_headers)
    check_col = sheet_columns(checks, 4, {"reconciliation"})["reconciliation"]
    years = expect_number(inputs, "B70")
    if years != int(years) or not 0 <= years <= 126:
        raise AssertionError(f"Invalid cached horizon count: {years}")
    years = int(years)
    if years == 0:
        require_blank(projection, ("C5", "C6"))
        if summary["B25"].value == "YES":
            raise AssertionError("Invalid horizon displayed as a funded plan")
        return {"modeled_years": 0}

    opening_age = expect_number(inputs, "B8")
    end_age = expect_number(inputs, "B53")
    retirement_age = expect_number(inputs, "B10")
    cash_share = expect_number(inputs, "B86")
    exempt_cash_share = expect_number(inputs, "B89")
    equal("Modeled horizon", years, end_age - opening_age + 1, 0)
    ending_row = years + 4
    forecasts = [(projection, 5)]
    if not modern:
        forecasts.extend(((wb["Monte_Carlo"], 25), (wb["Sequence_Risk"], 15)))
    for sheet, start in forecasts:
        equal(f"{sheet.title} opening age", expect_number(sheet, f"C{start}"), opening_age, 0)
        equal(f"{sheet.title} horizon age", expect_number(sheet, f"C{start + years - 1}"), end_age, 0)
        # Result table may begin shortly after the forecast grid; only inspect the 126-year grid.
        require_blank(sheet, (f"C{row}" for row in range(start + years, start + 126)))

    previous_tax_payable = 0.0
    max_shortfall = 0.0
    first_shortfall = None
    first_retirement_row = None
    healthcare_review = False
    for row in range(6, ending_row + 1):
        def p(column: str) -> float:
            return expect_number(projection, f"{column}{row}")

        def cf(header: str) -> float:
            return expect_number(cash, f"{cols[header]}{row}")

        def ra(header: str) -> float:
            return expect_number(roth, f"{roth_cols[header]}{row}")

        age = p("C")
        label = f"Age {age:g}"
        equal(f"{label} age roll-forward", age, opening_age + row - 5, 0)
        if age >= retirement_age and first_retirement_row is None:
            first_retirement_row = row

        for header in REQUIRED_CASH_HEADERS - {"cash income", "cash residual"}:
            if cf(header) < -tolerance:
                raise AssertionError(f"{label} negative {header}: {cf(header):.6f}")
        for column in ("J", "K", "L", "M", "P", "V", "W", "X", "Y", "Z", "AA", "AV", "BB", "BC", "BD", "BE", "BF", "BG", "BH"):
            if p(column) < -tolerance:
                raise AssertionError(f"{label} negative Projection!{column}{row}: {p(column):.6f}")

        spending = p("AW") + p("AX") + p("BA")
        income = p("I") + p("N") + p("O") * cash_share + p("AG") * exempt_cash_share - p("J") - p("K")
        if health is not None:
            h = lambda c: expect_number(health, f'{c}{row}')
            magi = p('AK')+p('AG')+p('N')-p('AI')
            credit = premium_credit(p('B'),magi,h('O'),h('M'),h('N'))
            equal(f'{label} gross healthcare',p('BA'),h('P'),tolerance)
            equal(f'{label} ACA household MAGI',h('AC'),magi,tolerance)
            equal(f'{label} ACA credit from statutory bands',h('R'),credit,tolerance)
            equal(f'{label} ACA credit booked once',expect_number(cash,f'AC{row}'),credit,tolerance)
            equal(f'{label} net healthcare',h('S'),h('P')-credit,tolerance)
            equal(f'{label} HSA draw',p('P'),min(expect_number(projection,f'BG{row-1}'),h('Q')),tolerance)
            status = health[f'T{row}'].value
            healthcare_review |= status not in ('OK','Legacy budget')
            income += credit
        equal(f"{label} cash income", cf("cash income"), income, tolerance)
        equal(f"{label} required spending", cf("required spending"), spending, tolerance)
        equal(f"{label} HSA paid", cf("hsa paid"), p("P"), tolerance)
        if p("P") > p("BA") + tolerance:
            raise AssertionError(f"{label} HSA pays more than total modeled healthcare")
        equal(f"{label} current tax", cf("current tax"), p("AV"), tolerance)
        equal(f"{label} opening tax payable", cf("opening tax payable"), previous_tax_payable, tolerance)
        tax_due = previous_tax_payable + p("AV")
        equal(f"{label} ending tax payable", cf("ending tax payable"), tax_due - cf("tax paid"), tolerance)
        equal(f"{label} unpaid spending", cf("unpaid spending"), spending - cf("spending paid") - p("P"), tolerance)
        equal(f"{label} Roth contribution", cf("roth contribution"), p("M"), tolerance)
        equal(f"{label} taxable deposit", cf("taxable deposit"), p("BB"), tolerance)

        residual = income + p("V") + p("W") + p("Y") - cf("tax paid") - cf("spending paid") - p("M") - p("BB")
        equal(f"{label} independently recomputed cash residual", residual, 0, tolerance)
        equal(f"{label} displayed cash residual", cf("cash residual"), residual, tolerance)
        shortfall = cf("ending tax payable") + cf("unpaid spending")
        equal(f"{label} essential shortfall", cf("essential shortfall"), shortfall, tolerance)
        equal(f"{label} Projection shortfall", p("Z"), shortfall, tolerance)
        if shortfall > tolerance and (p("M") > tolerance or p("BB") > tolerance):
            raise AssertionError(f"{label} cash went to elective savings while essential costs were unpaid")
        if cf("ending tax payable") > tolerance and cf("spending paid") > tolerance:
            raise AssertionError(f"{label} cash spending was paid before taxes")
        if p("Y") > p("X") + tolerance:
            raise AssertionError(f"{label} Roth withdrawal exceeds available Roth assets")
        equal(f"{label} Roth accessible", p("X"), ra("roth accessible"), tolerance)
        equal(f"{label} Roth closing", p("BF"), ra("roth closing"), tolerance)
        equal(f"{label} Roth opening balances", expect_number(projection, f"BF{row-1}"), sum(ra(h) for h in ("primary ira opening", "spouse ira opening", "primary workplace opening", "spouse workplace opening")), tolerance)
        equal(f"{label} Roth draws by account", p("Y"), sum(ra(h) for h in ("primary ira draw", "spouse ira draw", "primary workplace draw", "spouse workplace draw")), tolerance)
        equal(f"{label} Roth savings by owner", p("M"), ra("primary savings") + ra("spouse savings"), tolerance)
        equal(f"{label} Roth conversions by owner", p("AA"), ra("primary conversion") + ra("spouse conversion"), tolerance)
        if p("Y") > expect_number(projection, f"BF{row-1}") + tolerance:
            raise AssertionError(f"{label} Roth withdrawal exceeds the opening balance")
        if p("V") > expect_number(projection, f"BC{row-1}") + tolerance:
            raise AssertionError(f"{label} taxable withdrawal exceeds the opening balance")
        if p("W") + p("AA") > expect_number(projection, f"BE{row-1}") + tolerance:
            raise AssertionError(f"{label} traditional withdrawals and conversions exceed the opening balance")
        equal(f"{label} portfolio balance", p("BH"), p("BC") + p("BE") + p("BF") + p("BG"), tolerance)
        external_flow = p("J") + p("K") + p("L") + p("M") + p("BB") - p("V") - p("W") - p("Y") - p("P")
        equal(f"{label} portfolio cash flow", p("BO"), external_flow, tolerance)
        if checks[f"{check_col}{row}"].value != "OK":
            raise AssertionError(f"{label} workbook reconciliation flag is {checks[f'{check_col}{row}'].value!r}")

        previous_tax_payable = cf("ending tax payable")
        if first_shortfall is None and shortfall >= SHORTFALL_THRESHOLD:
            first_shortfall = p("B")
        max_shortfall = max(max_shortfall, shortfall)

    equal("Summary maximum shortfall", expect_number(summary, "B27"), max_shortfall, tolerance)
    if healthcare_review:
        if summary['B25'].value != 'Check healthcare inputs':
            raise AssertionError('Incomplete healthcare assumptions did not block a funded headline')
    elif first_shortfall is not None:
        if not str(summary["B25"].value).startswith("NO"):
            raise AssertionError("Summary shows survival despite unpaid required cash")
        equal("Summary first shortfall year", expect_number(summary, "B26"), first_shortfall, 0)
    elif summary["B25"].value != "YES":
        raise AssertionError(f"Summary does not show a funded plan: {summary['B25'].value!r}")
    if first_retirement_row is not None and expect_number(projection, f"C{first_retirement_row}") < 60:
        row = first_retirement_row
        bridge = expect_number(projection, f"BC{row - 1}") + expect_number(projection, f"X{row}")
        equal("Summary bridge funds", expect_number(summary, "B19"), bridge, tolerance)
    else:
        require_blank(summary, ("B19",))
    return {"modeled_years": years - 1, "end_age": end_age, "largest_shortfall": max_shortfall, "ending_tax_payable": previous_tax_payable}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--tolerance", type=float, default=0.01)
    args = parser.parse_args()
    from xlsx_acceptance import scan_errors
    with Workbook(args.workbook) as wb:
        scan_errors(wb)
        result = validate_model(wb, args.tolerance)
    check_chart_structure(args.workbook)
    check_probe_formula_encoding(args.workbook)
    print(f"PASS {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
