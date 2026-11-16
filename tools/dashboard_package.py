"""Keep the authored Excel dashboard charts when assembling and caching the workbook."""

from copy import deepcopy
from posixpath import basename, dirname, join, normpath

from lxml import etree as E

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
D = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"

_CHART_PREFIX = "'retirement-model.xlsx'!"
# Chart authoring uses fixed placeholders, which become dynamic names here.
_SERIES = {
    "RMChartPortfolioAge": ("A", "K2", 127),
    "RMChartPortfolioBalance": ("B", "K2", 127),
}
_GUIDES = {
    "RMChartRetirementAge": "AA",
    "RMChartRetirementBalance": "AB",
    "RMChartShortfallAge": "AC",
    "RMChartShortfallBalance": "AD",
}
_PLACEHOLDERS = {
    **{f"${column}$2:${column}${last}": name
       for name, (column, _, last) in _SERIES.items()},
    **{f"${column}$2:${column}$3": name for name, column in _GUIDES.items()},
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
        if name.startswith("RMPlan") or name in _SERIES or name in _GUIDES:
            names.remove(item)
    sheet = "Report_Data!"
    for name, (column, count, last) in _SERIES.items():
        first = f"{sheet}${column}$2"
        source = f"{first}:${column}${last}"
        c = f"{sheet}${count[0]}${count[1:]}"
        E.SubElement(names, X + "definedName", name=name).text = (
            f"{first}:INDEX({source},MAX(1,{c}))"
        )
    for name, column in _GUIDES.items():
        row = 5 if name.startswith("RMChartRetirement") else 6
        position = f"{sheet}$K${row}"
        active = f"AND({position}>0,{position}<={sheet}$K$2)"
        # Excel may plot formulas returning empty text as zeros. Use the
        # physically empty range in rows 4–5 when a guide is inactive.
        source = f"{sheet}${column}$2:${column}$5"
        E.SubElement(names, X + "definedName", name=name).text = (
            f"INDEX({source},IF({active},1,3)):INDEX({source},IF({active},2,4))"
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


def install_dashboard(src, authored, workbook, changes, summary_part, authored_summary_part):
    """Replace the old Summary chart with the editable portfolio chart from Artifact Tool."""
    drawing_part, drawing, relationships = _drawing(src, summary_part)
    authored_part, authored_drawing, authored_rels = _drawing(authored, authored_summary_part)
    [(old_anchor, old_chart)] = _chart_anchors(drawing)
    old_rel = _relationship(relationships, old_chart.get(R + "id"))
    original_chart_part = _target(drawing_part, old_rel)
    drawing.remove(old_anchor)
    relationships.remove(old_rel)
    [(anchor, chart)] = _chart_anchors(authored_drawing)
    part = _target(authored_part, _relationship(authored_rels, chart.get(R + "id")))
    root = _read(authored, part)
    _bind_chart(root, _SERIES.keys() | _GUIDES.keys())
    changes[original_chart_part] = _pack(root)
    frame = deepcopy(anchor)
    frame.find(".//" + C + "chart").set(R + "id", old_rel.get("Id"))
    props = frame.find(".//" + D + "cNvPr")
    props.set("id", str(max((int(node.get("id")) for node in drawing.iter(D + "cNvPr")), default=0) + 1))
    props.set("name", "Projected portfolio by age")
    drawing.append(frame)
    relationships.append(old_rel)
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
    if name in _GUIDES:
        column = _GUIDES[name]
        sheet = workbook["Report_Data"]
        return [sheet[f"{column}{row}"].value for row in (2, 3)]
    return None
