"""Keep the authored Excel dashboard charts when assembling and caching the workbook."""

from copy import deepcopy
from posixpath import basename, dirname, join, normpath

from lxml import etree as E

from risk_presentation import CURRENT, DATA_SHEET, SHEET

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
D = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"

_CHART_PREFIX = "'Report_Data'!"
# Chart authoring uses fixed placeholders, which become dynamic names here.
_SERIES = {
    "RMChartPortfolioAge": ("A", "Q2", 127),
    "RMChartPortfolioBalance": ("AE", "Q2", 127),
}
_RISK_SERIES = {
    "RMChartRiskFloor": "AF",
    "RMChartRiskBand": "AG",
    "RMChartRiskMedian": "AH",
}
_PLACEHOLDERS = {
    **{f"${column}$2:${column}${last}": name
       for name, (column, _, last) in _SERIES.items()},
    **{f"${column}$2:${column}$127": name for name, column in _RISK_SERIES.items()},
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
    report_id = next(str(index) for index, sheet in enumerate(workbook.find(X + 'sheets'))
                     if sheet.get('name') == 'Report_Data')
    for item in list(names):
        name = item.get("name", "")
        if name.startswith(("RMPlan", "RMChart")):
            names.remove(item)
    sheet = "Report_Data!"
    for name, (column, count, last) in _SERIES.items():
        first = f"{sheet}${column}$2"
        source = f"{first}:${column}${last}"
        c = f"{sheet}${count[0]}${count[1:]}"
        E.SubElement(names, X + "definedName", name=name, localSheetId=report_id).text = (
            f"{first}:INDEX({source},MAX(1,{c}))"
        )
    for name, column in _RISK_SERIES.items():
        start = f'IF({CURRENT},1,127)'
        length = f'MAX(1,{sheet}$Q$2)'
        # The second half is physically empty. It keeps risks off the chart
        # while preserving the length of the common category axis.
        source = f"{sheet}${column}$2:${column}$253"
        E.SubElement(names, X + "definedName", name=name, localSheetId=report_id).text = (
            f"INDEX({source},{start}):INDEX({source},{start}+{length}-1)"
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


def rebind_dashboard_workbook(root, filename):
    """Keep chart names on their worksheet when a user renames the file."""
    del filename  # Retained for callers that also name the exported workbook.
    names = _SERIES.keys() | _RISK_SERIES.keys()
    updated = False
    for formula in root.iter(C + "f"):
        qualifier, separator, name = (formula.text or "").rpartition("!")
        local = separator and qualifier.strip("'").lower().endswith(".xlsx")
        if (local or qualifier.strip("'") == 'Report_Data') and not any(
                character in qualifier for character in "/\\") and name in names:
            reference = _CHART_PREFIX + name
            if formula.text != reference:
                formula.text = reference
                updated = True
    return updated


def _combine_charts(line_chart, area_chart):
    """Use the area series and line series on the same pair of Excel axes."""
    area = deepcopy(area_chart.find(f".//{C}plotArea/{C}areaChart"))
    line = line_chart.find(f".//{C}plotArea/{C}lineChart")
    grouping = area.find(C + "grouping")
    if grouping is None:
        grouping = E.Element(C + "grouping")
        area.insert(0, grouping)
    grouping.set("val", "stacked")
    for source, axis in zip(area.findall(C + "axId"), line.findall(C + "axId"), strict=True):
        source.set("val", axis.get("val"))
    line.addprevious(area)
    for index, series in enumerate(area.findall(C + "ser") + line.findall(C + "ser")):
        series.find(C + "idx").set("val", str(index))
        series.find(C + "order").set("val", str(index))

    # The transparent area starts the visible band at the 10th percentile.
    first = area.find(C + "ser")
    fill = first.find(C + "spPr")
    if fill is None:
        fill = E.Element(C + "spPr")
        first.find(C + "tx").addnext(fill)
    fill.clear()
    drawing_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    E.SubElement(fill, drawing_ns + "noFill")
    E.SubElement(E.SubElement(fill, drawing_ns + "ln"), drawing_ns + "noFill")
    legend = line_chart.find(f".//{C}chart/{C}legend")
    hidden = E.Element(C + "legendEntry")
    E.SubElement(hidden, C + "idx", val="0")
    E.SubElement(hidden, C + "delete", val="1")
    legend.find(C + "legendPos").addnext(hidden)

    # Each age is a category. A decade between labels keeps long plans legible.
    age_axis = line_chart.find(f".//{C}plotArea/{C}catAx")
    if (gridlines := age_axis.find(C + "majorGridlines")) is not None:
        age_axis.remove(gridlines)
    E.SubElement(age_axis, C + "tickLblSkip", val="10")


def install_dashboard(src, authored, workbook, changes, summary_part, authored_summary_part):
    """Replace the old Summary chart with the compiled editable portfolio chart."""
    drawing_part, drawing, relationships = _drawing(src, summary_part)
    authored_part, authored_drawing, authored_rels = _drawing(authored, authored_summary_part)
    [(old_anchor, old_chart)] = _chart_anchors(drawing)
    old_rel = _relationship(relationships, old_chart.get(R + "id"))
    original_chart_part = _target(drawing_part, old_rel)
    drawing.remove(old_anchor)
    relationships.remove(old_rel)
    anchors = _chart_anchors(authored_drawing)
    anchor, chart = anchors[0]
    root = _read(authored, _target(authored_part, _relationship(authored_rels, chart.get(R + "id"))))
    if len(anchors) == 2:
        area_chart = anchors[1][1]
        area = _read(authored, _target(authored_part, _relationship(authored_rels, area_chart.get(R + "id"))))
        _combine_charts(root, area)
    elif len(anchors) != 1 or root.find(f'.//{C}plotArea/{C}areaChart') is None:
        raise ValueError('Dashboard source requires the line and percentile-band charts.')
    _bind_chart(root, _SERIES.keys() | _RISK_SERIES.keys())
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
    name = formula.rsplit("!", 1)[-1]
    if name in _SERIES:
        column, count, _ = _SERIES[name]
        sheet = workbook["Report_Data"]
        return [sheet[f"{column}{row}"].value for row in range(2, 2 + max(1, int(sheet[count].value or 0)))]
    if name in _RISK_SERIES:
        column = _RISK_SERIES[name]
        sheet = workbook["Report_Data"]
        count = max(1, int(sheet["Q2"].value or 0))
        risk_sheet = DATA_SHEET if DATA_SHEET in workbook.sheetnames else SHEET
        if workbook[risk_sheet]["Z1"].value != "current":
            return [None] * count
        return [sheet[f"{column}{row}"].value for row in range(2, 2 + count)]
    return None
