"""Independent workbook acceptance runner. Never writes to the source workbook."""

from __future__ import annotations

import argparse
import concurrent.futures
import json
import math
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
from xml.etree import ElementTree as ET
import zipfile

from xlsx_read import Workbook, address_parts, parse_preserving_namespaces, q, sheet_parts

ERRS = frozenset({"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"})

def patch_xml(sheet_xml: bytes, overrides: dict, last_row: int | None = None) -> bytes:
    root = parse_preserving_namespaces(sheet_xml)
    data = root.find(q("sheetData"))
    if last_row is not None:
        # Solver source banks share rows with trials but remain dependencies
        # even when a temporary test removes inactive annual trial groups.
        local_sources = root.find(f".//{q('c')}[@r='BE2']/{q('f')}") is not None
        for element in list(data):
            if int(element.get('r')) > last_row:
                if local_sources:
                    for cell in list(element):
                        if address_parts(cell.get('r'))[1] < 56:
                            element.remove(cell)
                if not len(element):
                    data.remove(element)
    rows = {int(r.get("r")): r for r in data.findall(q("row"))}
    cells = {c.get("r"): c for c in data.iter(q("c"))}
    for address, value in overrides.items():
        row_number, col_number = address_parts(address)
        cell = cells.get(address)
        if cell is None:
            if row_number not in rows:
                row = ET.Element(q("row"), r=str(row_number))
                nxt = next((old for old in data if int(old.get("r")) > row_number), None)
                if nxt is None:
                    data.append(row)
                else:
                    data.insert(list(data).index(nxt), row)
                rows[row_number] = row
            row = rows[row_number]
            cell = ET.Element(q("c"), r=address)
            nxt = next((old for old in row if address_parts(old.get("r"))[1] > col_number), None)
            if nxt is None:
                row.append(cell)
            else:
                row.insert(list(row).index(nxt), cell)
        if cell.find(q("f")) is not None:
            raise ValueError(f"Fixture attempted to overwrite a formula cell: {address}")
        for child in list(cell):
            cell.remove(child)
        cell.attrib.pop("t", None)
        if isinstance(value, bool):
            cell.set("t", "b")
            ET.SubElement(cell, q("v")).text = "1" if value else "0"
        elif isinstance(value, (int, float)) and math.isfinite(value):
            ET.SubElement(cell, q("v")).text = str(value)
        elif isinstance(value, str):
            cell.set("t", "inlineStr")
            ET.SubElement(ET.SubElement(cell, q("is")), q("t")).text = value
        elif value is not None:
            raise ValueError(f"Unsupported value at {address}: {value!r}")
    # LibreOffice may otherwise reuse the original XLSX formula cache on import,
    # even if a test changed literal inputs. Clear every formula result.
    for cell in cells.values():
        if cell.find(q("f")) is not None:
            cache = cell.find(q("v"))
            if cache is not None:
                cell.remove(cache)
            cell.attrib.pop("t", None)
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def prepare_xlsx(source: Path, destination: Path, overrides: dict, compact=False) -> None:
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as copy:
        parts = sheet_parts(original)
        if missing := overrides.keys() - parts.keys():
            raise ValueError(f"Unknown sheets: {sorted(missing)}")
        keep_rows = {}
        if compact and 'Targeting_Calc' in parts:
            # Only inactive funding groups are removed. Comparison fixtures retain the entire sweep.
            with Workbook(source) as model:
                inps = model['Inputs']
                changes = overrides.get('Inputs',{})
                selected = changes.get('B4',inps['B4'].value)
                if isinstance(selected,(int,float)) and selected in (1,2,3):
                    col = 'DEF'[int(selected)-1]
                    start = changes.get(f'{col}8',inps[f'{col}8'].value)
                    end = changes.get(f'{col}53',inps[f'{col}53'].value)
                    if all(isinstance(v,(int,float)) and int(v)==v for v in (start,end)) and 0<=start<=end<=125:
                        projection=model['Projection']
                        anchors=[re.search(r'Funding_Calc!\$[A-Z]+\$(\d+)',projection[f'V{r}'].formula or '') for r in (6,7)]
                        if all(anchors):
                            first,second=(int(hit[1]) for hit in anchors)
                            keep_rows[parts['Funding_Calc']]=first+(int(end-start)-1)*(second-first)
            if 'Income_Targeting' not in overrides:
                overrides={**overrides,'Income_Targeting':{'B5':126}}
                keep_rows[parts['Targeting_Calc']]=1
        sheet_overrides = {parts[name]: changes for name, changes in overrides.items()}
        worksheet_parts = set(parts.values())
        for item in original.infolist():
            payload = original.read(item.filename)
            if item.filename in worksheet_parts:
                payload = patch_xml(payload, sheet_overrides.get(item.filename, {}),keep_rows.get(item.filename))
            elif item.filename == "xl/workbook.xml":
                root = parse_preserving_namespaces(payload)
                calc = root.find(q("calcPr"))
                if calc is not None:
                    calc.set("calcMode", "auto")
                    calc.set("fullCalcOnLoad", "1")
                    calc.set("forceFullCalc", "1")
                    payload = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
            copy.writestr(item, payload)


def recalculate(source: Path, target_dir: Path, profile_dir: Path, executable: str, timeout: int) -> Path:
    result = subprocess.run(
        [executable, f"-env:UserInstallation={profile_dir.as_uri()}", "--headless", "--convert-to", "xlsx", "--outdir", str(target_dir), str(source)],
        capture_output=True, text=True, timeout=timeout, check=False,
    )
    calculated = target_dir / source.name
    if result.returncode or not calculated.is_file():
        raise RuntimeError(f"Recalculation failed: {result.stdout.strip()} {result.stderr.strip()}")
    return calculated


def cell_value(wb, address: str):
    sheet, coord = address.rsplit("!", 1)
    return wb[sheet][coord].value


def row_value(wb, spec: dict):
    sheet = wb[spec.get("sheet", "Projection")]
    age_column = spec.get("age_column", "C")
    first_row = spec.get("first_row", 5)
    last_row = spec.get("last_row", min(sheet.max_row, first_row + 126))
    matches = [row for row in range(first_row, last_row + 1) if sheet[f"{age_column}{row}"].value == spec["age"]]
    if len(matches) != 1:
        raise AssertionError(f"{sheet.title}: expected one row for age {spec['age']}, found {matches}")
    return sheet[f"{spec['column']}{matches[0]}"].value


def check_condition(wb, expected: dict, dollars_tolerance: float) -> None:
    key = expected.get("cell") or f"{expected.get('sheet', 'Projection')} age {expected['age']} col {expected['column']}"
    actual = cell_value(wb, expected["cell"]) if "cell" in expected else row_value(wb, expected)
    if "equals" in expected:
        target = expected["equals"]
        if isinstance(target, (int, float)) and not isinstance(target, bool):
            ok = isinstance(actual, (int, float)) and not isinstance(actual, bool) and abs(actual - target) <= expected.get("tolerance", dollars_tolerance)
        else:
            ok = actual == target
    elif "between" in expected:
        lower, upper = expected["between"]
        ok = isinstance(actual, (int, float)) and lower <= actual <= upper
    elif "prefix" in expected:
        ok = isinstance(actual, str) and actual.startswith(expected["prefix"])
    elif expected.get("blank"):
        ok = actual in (None, "")
    elif "same_as" in expected:
        target = cell_value(wb, expected["same_as"])
        ok = isinstance(actual, (int, float)) and isinstance(target, (int, float)) and abs(actual - target) <= expected.get("tolerance", dollars_tolerance)
    else:
        raise ValueError(f"No condition for {key}")
    if not ok:
        raise AssertionError(f"{key}: expected {expected}; got {actual!r}")


def check_rows(wb, rules: dict, tolerance: float) -> None:
    sheet = wb[rules.get("sheet", "Projection")]
    active = rules.get("active_column", "C")
    start = rules.get("first_flow_row", 6)
    end = rules.get("last_flow_row", min(sheet.max_row, start + 124))
    for row in range(start, end + 1):
        age = sheet[f"{active}{row}"].value
        if not isinstance(age, (int, float)):
            continue
        for column in rules.get("nonnegative", []):
            value = sheet[f"{column}{row}"].value
            if not isinstance(value, (int, float)) or value < -tolerance:
                raise AssertionError(f"{sheet.title}!{column}{row}, age {age}: expected nonnegative number, got {value!r}")
        for rule in rules.get("balance_equations", []):
            if rule.get("unless_column") and sheet[f"{rule['unless_column']}{row}"].value == rule["unless_equals"]:
                continue
            value = 0.0
            for coefficient, column in rule["terms"]:
                if not re.fullmatch(r"[A-Z]{1,3}", column):
                    raise ValueError(f"Invalid column {column!r}")
                item = sheet[f"{column}{row}"].value
                if not isinstance(item, (int, float)):
                    raise AssertionError(f"{sheet.title}!{column}{row} must be numeric, got {item!r}")
                value += coefficient * item
            if abs(value - rule.get("target", 0)) > rule.get("tolerance", tolerance):
                raise AssertionError(f"{sheet.title} row {row}, age {age}: {rule['name']} has residual {value:.6f}")


def scan_errors(wb) -> None:
    failures = []
    count = 0
    for sheet in wb:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type == "e" or (isinstance(cell.value, str) and cell.value in ERRS):
                    count += 1
                    if len(failures) < 12:
                        failures.append(f"{sheet.title}!{cell.coordinate}={cell.value}")
    if failures:
        raise AssertionError(f"Found {count} spreadsheet error cells: {', '.join(failures)}")


def run_case(args, case: dict, shared: dict, row_rules: dict) -> str:
    name = case["name"]
    if not re.fullmatch(r"[a-zA-Z0-9_-]+", name):
        raise ValueError(f"Unsafe fixture name {name!r}")
    with tempfile.TemporaryDirectory(dir=args.scratch, prefix=f"{name}-") as scratch:
        run_dir = Path(scratch).resolve()
        source = run_dir / args.workbook.name
        target = run_dir / "recalculated"
        target.mkdir()
        overrides = {sheet: {**shared.get(sheet, {}), **case.get("overrides", {}).get(sheet, {})} for sheet in shared.keys() | case.get("overrides", {}).keys()}
        prepare_xlsx(args.workbook, source, overrides,args.compact)
        result = recalculate(source, target, run_dir / "profile", args.soffice, args.timeout)
        with Workbook(result) as wb:
            scan_errors(wb)
            for rule in case.get("expects", []):
                check_condition(wb, rule, args.tolerance)
            check_rows(wb, row_rules | case.get("row_rules", {}), args.tolerance)
            if args.phase1:
                from model_acceptance import validate_model
                validate_model(wb, args.tolerance)
            if args.keep:
                args.keep.mkdir(parents=True, exist_ok=True)
                shutil.copy2(result, args.keep / f"{name}.xlsx")
    return name


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("fixtures", type=Path)
    parser.add_argument("--scratch", type=Path, default=Path(tempfile.gettempdir()))
    parser.add_argument("--keep", type=Path)
    parser.add_argument("--case", action="append")
    parser.add_argument("--soffice", default="soffice")
    parser.add_argument("--timeout", type=int, default=180)
    parser.add_argument("--jobs", type=int, default=2)
    parser.add_argument("--tolerance", type=float, default=0.01)
    parser.add_argument("--ledger", "--phase1", dest="phase1", action="store_true", help="Also reconcile each active year's cash, tax, account access and healthcare")
    parser.add_argument("--compact",action="store_true",help="Remove inactive phase-two solver groups in temporary fixtures only")
    args = parser.parse_args()
    fixtures = json.loads(args.fixtures.read_text())
    cases = [case for case in fixtures["cases"] if not args.case or case["name"] in args.case]
    if not cases:
        parser.error("No matching cases")
    args.scratch.mkdir(parents=True, exist_ok=True)
    failures = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=args.jobs) as executor:
        futures = {executor.submit(run_case, args, case, fixtures.get("common_overrides", {}), fixtures.get("row_rules", {})): case["name"] for case in cases}
        for future in concurrent.futures.as_completed(futures):
            name = futures[future]
            try:
                future.result()
            except Exception as error:
                print(f"FAIL {name}: {error}", flush=True)
                failures += 1
            else:
                print(f"PASS {name}", flush=True)
    print(f"{len(cases) - failures}/{len(cases)} scenarios passed")
    return int(bool(failures))


if __name__ == "__main__":
    raise SystemExit(main())
