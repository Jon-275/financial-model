"""Keep the authored Excel dashboard charts when assembling and caching the workbook."""

from copy import deepcopy
from posixpath import basename, dirname, join, normpath, relpath
import re

from lxml import etree as E

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
D = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"

_CHART_PREFIX = "'retirement-model.xlsx'!"
_CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
# Chart authoring uses fixed placeholders, which become dynamic names here.
_SERIES = {
    "RMChartPortfolioAge": ("A", "K2", 127),
    "RMChartPortfolioBalance": ("B", "K2", 127),
    "RMChartCashAge": ("E", "K4", 14),
    "RMChartCashIncome": ("F", "K4", 14),
    "RMChartCashWithdrawals": ("G", "K4", 14),
    "RMChartCashHealth": ("H", "K4", 14),
    "RMChartCashUnpaid": ("I", "K4", 14),
}
_RETIREMENT = {"RMChartRetirementAge": "A", "RMChartRetirementBalance": "B"}
_PLACEHOLDERS = {
    **{f"${column}$2:${column}${last}": name
       for name, (column, _, last) in _SERIES.items()},
    **{f"${column}$2": name for name, column in _RETIREMENT.items()},
}


def _read(archive, part):
    return E.fromstring(archive.read(part))


def _pack(element):
    return E.tostring(element, encoding="UTF-8", xml_declaration=True, standalone=True)


def _rels(part):
    return join(dirname(part), "_rels", basename(part) + ".rels")


def _target(part, relationship):
    path = relationship.get("Target")
    return path.lstrip("/") if path.startswith("/") else normpath(join(dirname(part), path))


def _relationship(relationships, identifier):
    return next(item for item in relationships if item.get("Id") == identifier)


def _drawing(archive, sheet_part):
    ref = _read(archive, sheet_part).find(X + "drawing").get(R + "id")
    drawing = _target(sheet_part, _relationship(_read(archive, _rels(sheet_part)), ref))
    return drawing, _read(archive, drawing), _read(archive, _rels(drawing))


def _chart_anchors(drawing):
    return [(anchor, chart) for anchor in drawing
            if (chart := anchor.find(".//" + C + "chart")) is not None]


def _add_names(workbook):
    names = workbook.find(X + "definedNames")
    for item in list(names):
        name = item.get("name", "")
        if name.startswith("RMPlan") or name in _SERIES or name in _RETIREMENT:
            names.remove(item)
    sheet = "Report_Data!"
    for name, (column, count, last) in _SERIES.items():
        first = f"{sheet}${column}$2"
        source = f"{first}:${column}${last}"
        c = f"{sheet}${count[0]}${count[1:]}"
        E.SubElement(names, X + "definedName", name=name).text = (
            f"{first}:INDEX({source},MAX(1,{c}))"
        )
    for name, column in _RETIREMENT.items():
        # Row 128 is physically empty. A retirement outside the horizon uses
        # that empty cell for both coordinates, so Excel plots no marker.
        source = f"{sheet}${column}$2:${column}$128"
        idx = f"{sheet}$K$5"
        E.SubElement(names, X + "definedName", name=name).text = (
            f"INDEX({source},IF(AND({idx}>=1,{idx}<={sheet}$K$2),{idx},127))"
        )


def _bind_chart(root, expected):
    changed = set()
    for formula in root.iter(C + "f"):
        if not formula.text:
            continue
        sheet, sep, address = formula.text.partition("!")
        if sep and sheet.strip("'") == "Report_Data":
            name = _PLACEHOLDERS[address]
            formula.text = _CHART_PREFIX + name
            changed.add(name)
    if changed != expected:
        raise ValueError(f"Dashboard chart series mismatch: {sorted(changed)}")
    root.find(".//" + C + "plotVisOnly").set("val", "0")


def install_dashboard(src, authored, workbook, content_types, changes, summary_part, authored_summary_part):
    """Replace the old Summary chart with the editable charts from Artifact Tool."""
    drawing_part, drawing, relationships = _drawing(src, summary_part)
    authored_part, authored_drawing, authored_rels = _drawing(authored, authored_summary_part)
    [(old_anchor, old_chart)] = _chart_anchors(drawing)
    old_rel = _relationship(relationships, old_chart.get(R + "id"))
    original_chart_part = _target(drawing_part, old_rel)
    drawing.remove(old_anchor)
    relationships.remove(old_rel)
    next_id = max((int(element.get("id")) for element in drawing.iter(D + "cNvPr")), default=0) + 1
    next_chart = 1 + max(
        int(match.group(1)) for part in src.namelist()
        if (match := re.fullmatch(r"xl/charts/chart(\d+)\.xml", part))
    )
    chart_parts = (original_chart_part, f"xl/charts/chart{next_chart}.xml")
    expected = (
        {name for name in _SERIES if name.startswith("RMChartPortfolio")} | _RETIREMENT.keys(),
        {name for name in _SERIES if name.startswith("RMChartCash")},
    )
    for i, ((anchor, chart), target, references) in enumerate(
        zip(_chart_anchors(authored_drawing), chart_parts, expected, strict=True), next_id
    ):
        raw_part = _target(authored_part, _relationship(authored_rels, chart.get(R + "id")))
        root = _read(authored, raw_part)
        _bind_chart(root, references)
        changes[target] = _pack(root)
        used_ids = {item.get("Id") for item in relationships}
        n = 1
        while f"rId{n}" in used_ids:
            n += 1
        identifier = f"rId{n}"
        E.SubElement(relationships, PKG + "Relationship", Id=identifier,
                     Type=R.strip("{}") + "/chart", Target=relpath(target, dirname(drawing_part)))
        frame = deepcopy(anchor)
        frame.find(".//" + C + "chart").set(R + "id", identifier)
        props = frame.find(".//" + D + "cNvPr")
        props.set("id", str(i))
        props.set("name", "Portfolio balance" if i == next_id else "How the plan pays its bills")
        drawing.append(frame)
        if not any(item.get("PartName") == "/" + target for item in content_types):
            E.SubElement(content_types, CT + "Override", PartName="/" + target,
                         ContentType=_CHART_TYPE)
    changes[drawing_part] = _pack(drawing)
    changes[_rels(drawing_part)] = _pack(relationships)
    _add_names(workbook)


def dashboard_chart_data(formula, workbook):
    """Return calculated points for a dashboard reference, or None for another chart."""
    if not formula.startswith(_CHART_PREFIX):
        return None
    name = formula.removeprefix(_CHART_PREFIX)
    if name in _SERIES:
        column, count, _ = _SERIES[name]
        sheet = workbook["Report_Data"]
        return [sheet[f"{column}{row}"].value for row in range(2, 2 + max(1, int(sheet[count].value or 0)))]
    if name in _RETIREMENT:
        sheet = workbook["Report_Data"]
        point = int(sheet["K5"].value or 0)
        if 1 <= point <= (sheet["K2"].value or 0):
            return [sheet[f"{_RETIREMENT[name]}{point + 1}"].value]
        return [None]
    return None
