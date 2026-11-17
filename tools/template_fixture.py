"""Create a synthetic build template while preserving its native Excel parts."""

import argparse
from copy import deepcopy
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as E
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from example_inputs import example_asset_cells, example_cells
import workbook_metadata


X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def parts(archive):
    workbook = E.fromstring(archive.read('xl/workbook.xml'))
    relationships = E.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    targets = {node.get('Id'): node.get('Target') for node in relationships}
    return {
        node.get('name'): target.lstrip('/') if target.startswith('/') else 'xl/' + target
        for node in workbook.find(X + 'sheets')
        for target in (targets[node.get(R + 'id')],)
    }


def fixture_plan(source):
    """Send the changed cell content through the normal artifact author."""
    cells = {'Inputs': example_cells(), 'Asset_Allocation': example_asset_cells()}
    with ZipFile(source) as archive:
        sheets = parts(archive)
        for name in ('Income_Targeting', 'Sequence_Risk', 'Monte_Carlo'):
            root = E.fromstring(archive.read(sheets[name]))
            cells[name] = {
                cell.get('r'): 0 for cell in root.iter(X + 'c')
                if cell.find(X + 'f') is None
                and cell.find(X + 'v') is not None
                and cell.get('t', 'n') == 'n'
            }
    return {'order': list(sheets), 'cells': cells}


def package(source, authored, destination):
    """Overlay authored values and discard every inherited financial cache."""
    edits = fixture_plan(source)['cells']
    with ZipFile(source) as original, ZipFile(authored) as staged:
        sheets, written = parts(original), parts(staged)
        strings = []
        if 'xl/sharedStrings.xml' in staged.namelist():
            strings = [
                ''.join(node.itertext())
                for node in E.fromstring(staged.read('xl/sharedStrings.xml'))
            ]
        changes = {}
        for name, part in sheets.items():
            root = E.fromstring(original.read(part))
            data = root.find(X + 'sheetData')
            rows = {int(row.get('r')): row for row in data}
            existing = {cell.get('r'): cell for cell in data.iter(X + 'c')}
            authored_cells = {
                cell.get('r'): cell
                for cell in E.fromstring(staged.read(written[name])).iter(X + 'c')
            }
            for address in edits.get(name, {}):
                column, row_number = coordinate_from_string(address)
                row = rows.get(row_number)
                if row is None:
                    row = rows[row_number] = E.Element(X + 'row', r=str(row_number))
                old = existing.get(address)
                replacement = deepcopy(authored_cells.get(address, E.Element(X + 'c', r=address)))
                if replacement.get('t') == 's':
                    value = strings[int(replacement.findtext(X + 'v'))]
                    replacement.remove(replacement.find(X + 'v'))
                    replacement.set('t', 'inlineStr')
                    E.SubElement(E.SubElement(replacement, X + 'is'), X + 't').text = value
                if old is not None:
                    if old.get('s') is not None:
                        replacement.set('s', old.get('s'))
                    row.replace(old, replacement)
                else:
                    row.append(replacement)
            for row in rows.values():
                row[:] = sorted(row, key=lambda cell: column_index_from_string(
                    coordinate_from_string(cell.get('r'))[0]))
            data[:] = [rows[number] for number in sorted(rows)]
            last_column = max(
                (coordinate_from_string(cell.get('r'))[0] for cell in data.iter(X + 'c')),
                key=column_index_from_string,
            )
            root.find(X + 'dimension').set('ref', f'A1:{last_column}{max(rows)}')
            for cell in data.iter(X + 'c'):
                if cell.find(X + 'f') is not None:
                    for cached in cell.findall(X + 'v'):
                        cell.remove(cached)
                    cell.attrib.pop('t', None)
            workbook_metadata.clear_revision_ids(root)
            changes[part] = workbook_metadata._pack(root)
        for part in original.namelist():
            if part.startswith('xl/charts/') and part.endswith('.xml'):
                chart = E.fromstring(original.read(part))
                for cache in chart.findall('.//' + C + 'numCache') + chart.findall('.//' + C + 'strCache'):
                    cache.getparent().remove(cache)
                changes[part] = workbook_metadata._pack(chart)
        book = E.fromstring(original.read('xl/workbook.xml'))
        app = E.fromstring(original.read('docProps/app.xml'))
        types = E.fromstring(original.read('[Content_Types].xml'))
        relationships = E.fromstring(original.read('xl/_rels/workbook.xml.rels'))
        workbook_metadata.make_sheets_visible(book)
        removed = workbook_metadata.clean_metadata(
            original, book, app, types, relationships, changes)
        workbook_metadata.compact_shared_strings(original, sheets.values(), changes)
        for part, root in (
            ('xl/workbook.xml', book), ('docProps/app.xml', app),
            ('[Content_Types].xml', types), ('xl/_rels/workbook.xml.rels', relationships),
        ):
            changes[part] = workbook_metadata._pack(root)
        destination.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(destination, 'x', ZIP_DEFLATED, compresslevel=6) as output:
            for part in original.namelist():
                if part not in removed:
                    output.writestr(part, changes.get(part, original.read(part)))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--authored', type=Path)
    args = parser.parse_args()
    if args.authored:
        package(args.source, args.authored, args.output)
    else:
        args.output.write_text(json.dumps(fixture_plan(args.source)) + '\n')


if __name__ == '__main__':
    main()
