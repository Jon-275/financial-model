"""Write compiled cells and native chart sources using public OOXML libraries."""

from collections import defaultdict
from copy import deepcopy
import json
from pathlib import Path
import sys
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as E
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
T = 'http://schemas.openxmlformats.org/package/2006/content-types'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
D = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'


def pack(root):
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def author(plan, destination):
    workbook = E.Element(f'{{{X}}}workbook', nsmap={None: X, 'r': R})
    sheets = E.SubElement(workbook, f'{{{X}}}sheets')
    relationships = E.Element(f'{{{P}}}Relationships', nsmap={None: P})
    types = E.Element(f'{{{T}}}Types', nsmap={None: T})
    E.SubElement(types, f'{{{T}}}Default', Extension='rels',
                 ContentType='application/vnd.openxmlformats-package.relationships+xml')
    E.SubElement(types, f'{{{T}}}Default', Extension='xml', ContentType='application/xml')
    payloads = {}
    for index, name in enumerate(plan['order'], 1):
        part = f'xl/worksheets/sheet{index}.xml'
        identifier = f'rId{index}'
        E.SubElement(sheets, f'{{{X}}}sheet', name=name, sheetId=str(index),
                     attrib={f'{{{R}}}id': identifier})
        E.SubElement(relationships, f'{{{P}}}Relationship', Id=identifier,
                     Type=R + '/worksheet', Target=f'worksheets/sheet{index}.xml')
        E.SubElement(types, f'{{{T}}}Override', PartName='/' + part,
                     ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        root = E.Element(f'{{{X}}}worksheet', nsmap={None: X, 'r': R})
        data = E.SubElement(root, f'{{{X}}}sheetData')
        rows = defaultdict(list)
        for address, value in plan['cells'].get(name, {}).items():
            column, row = coordinate_from_string(address)
            rows[row].append((column_index_from_string(column), address, value))
        for number, values in sorted(rows.items()):
            row = E.SubElement(data, f'{{{X}}}row', r=str(number))
            for _, address, value in sorted(values):
                cell = E.SubElement(row, f'{{{X}}}c', r=address)
                if isinstance(value, str):
                    if value.startswith('='):
                        E.SubElement(cell, f'{{{X}}}f').text = value[1:]
                    else:
                        cell.set('t', 'inlineStr')
                        E.SubElement(E.SubElement(cell, f'{{{X}}}is'),
                                     f'{{{X}}}t').text = value
                elif isinstance(value, bool):
                    cell.set('t', 'b')
                    E.SubElement(cell, f'{{{X}}}v').text = str(int(value))
                elif value is not None:
                    E.SubElement(cell, f'{{{X}}}v').text = str(value)
        if name == 'Summary' and 'Report_Data' in plan['order']:
            E.SubElement(root, f'{{{X}}}drawing', attrib={f'{{{R}}}id': 'rId1'})
            sheet_rels = E.Element(f'{{{P}}}Relationships', nsmap={None: P})
            E.SubElement(sheet_rels, f'{{{P}}}Relationship', Id='rId1',
                         Type=R + '/drawing', Target='../drawings/drawing1.xml')
            payloads[f'xl/worksheets/_rels/sheet{index}.xml.rels'] = pack(sheet_rels)
        payloads[part] = pack(root)
    if 'Report_Data' in plan['order']:
        source = E.parse(str(Path(__file__).with_name('dashboard.xml'))).getroot()
        chart = deepcopy(source.find(f'{{{C}}}chartSpace'))
        anchor = deepcopy(source.find(f'{{{D}}}twoCellAnchor'))
        if anchor is None:
            anchor = deepcopy(source.find(f'{{{D}}}oneCellAnchor'))
        if anchor is None:
            raise ValueError('Dashboard source has no supported chart frame.')
        anchor.find(f'.//{{{C}}}chart').set(f'{{{R}}}id', 'rId1')
        drawing = E.Element(f'{{{D}}}wsDr', nsmap={'xdr': D})
        drawing.append(anchor)
        drawing_rels = E.Element(f'{{{P}}}Relationships', nsmap={None: P})
        E.SubElement(drawing_rels, f'{{{P}}}Relationship', Id='rId1',
                     Type=R + '/chart', Target='../charts/chart1.xml')
        payloads['xl/drawings/drawing1.xml'] = pack(drawing)
        payloads['xl/drawings/_rels/drawing1.xml.rels'] = pack(drawing_rels)
        payloads['xl/charts/chart1.xml'] = pack(chart)
        for part, content_type in (
            ('xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
            ('xl/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'),
        ):
            E.SubElement(types, f'{{{T}}}Override', PartName='/' + part, ContentType=content_type)
    E.SubElement(types, f'{{{T}}}Override', PartName='/xl/workbook.xml',
                 ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml')
    root_rels = E.Element(f'{{{P}}}Relationships', nsmap={None: P})
    E.SubElement(root_rels, f'{{{P}}}Relationship', Id='rId1',
                 Type=R + '/officeDocument', Target='xl/workbook.xml')
    payloads.update({
        'xl/workbook.xml': pack(workbook),
        'xl/_rels/workbook.xml.rels': pack(relationships),
        '[Content_Types].xml': pack(types), '_rels/.rels': pack(root_rels),
    })
    destination.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(destination, 'x', ZIP_DEFLATED, compresslevel=6) as output:
        for name, data in payloads.items():
            output.writestr(name, data)


if __name__ == '__main__':
    plan_path, destination = map(Path, sys.argv[1:])
    author(json.loads(plan_path.read_text()), destination)
