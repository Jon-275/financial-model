"""Verify independent accounting identities against the saved workbook."""

from __future__ import annotations

import argparse
import math
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from dashboard_acceptance import planning_metrics
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
        parts = sheet_parts(archive)
        modern = "Risk_Settings" in parts
        charts = sorted(name for name in archive.namelist() if name.startswith("xl/charts/chart") and name.endswith(".xml"))
        expected = 3 if "Annual Detail" in parts else 4 if modern else 5
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


def check_release_package(workbook: Path) -> None:
    with ZipFile(workbook) as archive:
        book = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = book.findall(f"{q('sheets')}/{q('sheet')}")
        hidden = [sheet.get("name") for sheet in sheets if sheet.get("state", "visible") != "visible"]
        if not sheets or hidden:
            raise AssertionError(f"Release workbook must expose all worksheets. Hidden tabs: {hidden}")
        for view in book.findall(f"{q('bookViews')}/{q('workbookView')}"):
            if view.get("showSheetTabs", "1") in ("0", "false"):
                raise AssertionError("Release workbook must display its sheet tabs.")
        discarded = {"docProps/custom.xml", "docMetadata/LabelInfo.xml", "xl/calcChain.xml"}
        if set(archive.namelist()) & discarded:
            raise AssertionError("Release workbook contains inherited document metadata.")
        for part in ("[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels"):
            for entry in ET.fromstring(archive.read(part)):
                target = (entry.get("PartName") or entry.get("Target") or "").lstrip("/")
                if part == "xl/_rels/workbook.xml.rels" and target == "calcChain.xml":
                    target = "xl/calcChain.xml"
                if target in discarded:
                    raise AssertionError(f"Release workbook still references removed metadata in {part}.")
        if len(ET.fromstring(archive.read("docProps/core.xml"))) or book.find("{*}revisionPtr") is not None:
            raise AssertionError("Release workbook contains document history or author properties.")


def check_plan_print_layout(workbook: Path) -> None:
    with ZipFile(workbook) as archive:
        parts = sheet_parts(archive)
        if 'Annual Detail' not in parts:
            return
        book = ET.fromstring(archive.read('xl/workbook.xml'))
        tabs = [sheet.get('name') for sheet in book.find(q('sheets'))]
        names = {(tabs[int(name.get('localSheetId'))], name.get('name')): name.text
                 for name in book.find(q('definedNames')) if name.get('localSheetId') is not None}
        consolidated = 'Risk_Data' in parts
        if 'Conversion Detail' in parts:
            expected_tabs = ['Start Here', 'Inputs', 'Summary', 'Asset_Allocation', 'Income_Targeting',
                             'Risk_Settings', 'Risk Results', 'Tax_Tables', 'Annual Detail', 'Appendix']
            if 'Planning Metrics' in parts:
                expected_tabs.insert(3, 'Planning Metrics')
            assert tabs[:len(expected_tabs)] == expected_tabs
            assert all(tabs.index(name) > tabs.index('Appendix') for name in ('Conversion Detail', 'Risk_Data'))
        elif consolidated:
            assert tabs[:9] == ['Start Here', 'Inputs', 'Summary', 'Asset_Allocation', 'Income_Targeting',
                                 'Risk_Settings', 'Risk Results', 'Annual Detail', 'Appendix']
            assert tabs.index('Risk_Data') > tabs.index('Appendix')
        else:
            assert tabs[:5] == ['Start Here', 'Inputs', 'Summary', 'Annual Detail', 'Appendix']
        for sheet, part in parts.items():
            view = ET.fromstring(archive.read(part)).find(f'{q("sheetViews")}/{q("sheetView")}')
            assert view.get('zoomScale') == '100' and view.get('showGridLines') == '0', sheet
        printed = {'Summary', 'Annual Detail'} if consolidated else {'Inputs', 'Summary', 'Annual Detail'}
        assert {sheet for sheet, name in names if name == '_xlnm.Print_Area'} == printed
        if not consolidated:
            assert names.get(('Inputs', '_xlnm.Print_Area')) == 'Inputs!$L$136:$S$194'
            input_breaks = ET.fromstring(archive.read(parts['Inputs'])).find(q('rowBreaks'))
            assert [(point.get('id'), point.get('man')) for point in input_breaks] == [('162', '1')]
        assert names.get(('Summary', '_xlnm.Print_Area')) == f'Summary!$B$2:$I${157 if consolidated else 92}'
        assert names.get(('Annual Detail', '_xlnm.Print_Titles')) == "'Annual Detail'!$5:$5"
        root = ET.fromstring(archive.read(parts['Summary']))
        if 'Planning Metrics' in parts:
            link = root.find(f"{q('hyperlinks')}/{q('hyperlink')}[@ref='K34']")
            assert link is not None and link.get('location') == "'Planning Metrics'!A1"
            addresses = {cell.get('r') for cell in root.iter(q('c'))}
            for address in (f'{column}{row}' for column in 'MNO' for row in range(1, 51)):
                assert address not in addresses, address
        breaks = root.find(q('rowBreaks'))
        expected_breaks = ('32', '64', '96', '128') if consolidated else ('32', '64')
        assert [(point.get('id'), point.get('man')) for point in breaks] == [(row, '1') for row in expected_breaks]
        assert root.find(q('pageSetup')).get('orientation') == 'portrait'
        if 'Risk Results' in parts:
            views = root.find(f"{q('dataValidations')}/{q('dataValidation')}[@sqref='K18']")
            assert views is not None and views.get('type') == 'list'
            assert views.findtext(q('formula1')) == '"Today’s dollars,Future dollars"'
        if 'Conversion Detail' in parts:
            check_schedule_layout(archive, parts)


def check_schedule_layout(archive, parts):
    sheets = {
        'Annual Detail': ('B6:J131', '$B6', 'D131'), 'Risk Results': ('B58:K183', '$B58', 'F183'),
        'Projection': ('A5:CF130', '$A5', 'B130'), 'Roth_Access': ('A5:AP130', '$A5', 'E130'),
        'Cash_Flow': ('A5:AQ130', '$A5', 'C130'), 'Healthcare': ('A5:AF130', '$A5', 'C130'),
        'Checks': ('A5:M130', '$A5', 'C130'), 'Conversion Detail': ('A5:V55', '$A5', 'D55'),
        'Report_Data': ('A2:N127', '$A2', 'A127'), 'Risk_Data': ('A5:I130', '$A5', 'C130'),
        'Funding_Calc': ('A2:AO15376', '$A2', 'D15376'),
        'Targeting_Calc': ('A2:AP6397', '$A2', 'D6397'),
    }
    cash = ET.fromstring(archive.read(parts['Cash_Flow']))
    if cash.find(f".//{q('c')}[@r='AR4']") is not None:
        sheets['Cash_Flow'] = ('A5:AW130', '$A5', 'C130')
    styles = ET.fromstring(archive.read('xl/styles.xml')).find(q('cellXfs'))
    for sheet, (area, key, address) in sheets.items():
        raw = ET.fromstring(archive.read(parts[sheet]))
        fmt = raw.find(f"{q('conditionalFormatting')}[@sqref='{area}']/{q('cfRule')}/{q('formula')}")
        assert fmt is not None and fmt.text == f'AND({key}<>"",MOD(ROW(),2)=1)', sheet
        if sheet in ('Funding_Calc', 'Targeting_Calc'):
            assert raw.find(q('autoFilter')).get('ref') == area.replace('A2:', 'A1:'), sheet
        cell = raw.find(f".//{q('c')}[@r='{address}']")
        assert cell is not None, (sheet, address)
        style = styles[int(cell.get('s', '0'))]
        assert style.get('fillId', '0') == style.get('borderId', '0') == '0', (sheet, address)


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

    inputs, projection, cash, roth, checks = (
        wb[name] for name in ("Inputs", "Projection", "Cash_Flow", "Roth_Access", "Checks")
    )
    summary, metric_col = planning_metrics(wb)
    status_address, first_gap, max_gap, bridge_value = (
        f'{metric_col}{row}' for row in (25, 26, 27, 19)
    )
    cols = sheet_columns(cash, 4, REQUIRED_CASH_HEADERS)
    has_health = "Healthcare" in wb.parts
    health = wb['Healthcare'] if has_health else None
    account_types = ('roth ira', 'workplace roth') if 'Risk_Data' in wb.parts else ('ira', 'workplace')
    owners = tuple(f'{owner} {kind}' for kind in account_types for owner in ('primary', 'spouse'))
    roth_openings = tuple(f'{owner} opening' for owner in owners)
    roth_draws = tuple(f'{owner} draw' for owner in owners)
    roth_headers = {"roth accessible", "roth closing", "primary savings", "spouse savings",
                    "primary conversion", "spouse conversion", *roth_openings, *roth_draws}
    roth_cols = sheet_columns(roth, 4, roth_headers)
    check_col = sheet_columns(checks, 4, {"reconciliation"})["reconciliation"]
    years = expect_number(inputs, "B70")
    if years != int(years) or not 0 <= years <= 126:
        raise AssertionError(f"Invalid cached horizon count: {years}")
    years = int(years)
    if years == 0:
        require_blank(projection, ("C5", "C6"))
        if summary[status_address].value == "YES":
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
    previous_spending_payable = 0.0
    cumulative_unpaid_spending = 0.0
    carries_spending = cash['AR4'].value == 'Opening unpaid spending'
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

        current_spending = p("AW") + p("AX") + p("BA")
        spending = current_spending + (previous_spending_payable if carries_spending else 0)
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
        if carries_spending:
            opening = expect_number(cash, f'AR{row}')
            repaid = expect_number(cash, f'AS{row}')
            current_paid = cf("spending paid") - repaid
            newly_unpaid = max(0, current_spending - p("P") - current_paid)
            cumulative_unpaid_spending += newly_unpaid
            equal(f'{label} opening spending obligations', opening, previous_spending_payable, tolerance)
            equal(f'{label} prior spending repaid first', repaid, min(opening, cf("spending paid")), tolerance)
            equal(f'{label} current spending paid', expect_number(cash, f'AT{row}'), current_paid, tolerance)
            equal(f'{label} newly unpaid spending', expect_number(cash, f'AU{row}'), newly_unpaid, tolerance)
            equal(f'{label} cumulative new obligations', expect_number(cash, f'AV{row}'), cumulative_unpaid_spending, tolerance)
            equal(f'{label} obligation roll-forward', opening + newly_unpaid - repaid, cf("unpaid spending"), tolerance)
            equal(f'{label} net resources after obligations', p('BJ'), p('BH')-p('BI')-cf("ending tax payable")-cf("unpaid spending"), tolerance)
            previous_spending_payable = cf("unpaid spending")
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
        equal(f"{label} Roth opening balances", expect_number(projection, f"BF{row-1}"), sum(ra(h) for h in roth_openings), tolerance)
        equal(f"{label} Roth draws by account", p("Y"), sum(ra(h) for h in roth_draws), tolerance)
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

    equal("Summary maximum shortfall", expect_number(summary, max_gap), max_shortfall, tolerance)
    if healthcare_review:
        if summary[status_address].value != 'Check healthcare inputs':
            raise AssertionError('Incomplete healthcare assumptions did not block a funded headline')
    elif first_shortfall is not None:
        if not str(summary[status_address].value).startswith("NO"):
            raise AssertionError("Summary shows survival despite unpaid required cash")
        equal("Summary first shortfall year", expect_number(summary, first_gap), first_shortfall, 0)
    elif summary[status_address].value != "YES":
        raise AssertionError(f"Summary does not show a funded plan: {summary[status_address].value!r}")
    if first_retirement_row is not None and expect_number(projection, f"C{first_retirement_row}") < 60:
        row = first_retirement_row
        bridge = expect_number(projection, f"BC{row - 1}") + expect_number(projection, f"X{row}")
        equal("Summary bridge funds", expect_number(summary, bridge_value), bridge, tolerance)
    else:
        require_blank(summary, (bridge_value,))
    return {"modeled_years": years - 1, "end_age": end_age, "largest_shortfall": max_shortfall, "ending_tax_payable": previous_tax_payable}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--tolerance", type=float, default=0.01)
    parser.add_argument("--release", action="store_true", help="Check published sheet visibility and document metadata.")
    args = parser.parse_args()
    if args.release:
        check_release_package(args.workbook)
        check_plan_print_layout(args.workbook)
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
