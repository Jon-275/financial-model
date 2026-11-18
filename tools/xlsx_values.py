"""Write selected literals without resaving workbook objects."""
from math import isfinite
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED

from lxml import etree as E
from openpyxl.utils.cell import coordinate_from_string,column_index_from_string,get_column_letter

X='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def _read_xml_part(archive, part, *, worksheet=False):
    parser = E.XMLParser(resolve_entities=False, load_dtd=False, no_network=True, huge_tree=worksheet)
    with archive.open(part) as stream:
        return E.parse(stream, parser).getroot()


def sheet_parts(archive):
    book=_read_xml_part(archive, 'xl/workbook.xml')
    relationships=_read_xml_part(archive, 'xl/_rels/workbook.xml.rels')
    paths={r.get('Id'):r.get('Target') for r in relationships}
    return {s.get('name'):(lambda p:p.lstrip('/') if p.startswith('/') else 'xl/'+p)(paths[s.get(R+'id')])
            for s in book.find(X+'sheets')}


def _cell_reference(address):
    column, row = coordinate_from_string(address)
    return f'{column.upper()}{row}', row


def _write_text(parent, value):
    text = E.SubElement(parent, X + 't')
    if value != value.strip():
        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    text.text = value


def update_sheet(root,values,*,shared_text=None):
    reference_delta=0
    data=root.find(X+'sheetData')
    rows={int(r.get('r')):r for r in data}
    cells = {}
    for cell in data.iter(X + 'c'):
        original = cell.get('r')
        address, _ = _cell_reference(original)
        if address in cells:
            raise ValueError(f'The worksheet contains duplicate references to cell {address}.')
        if original != address:
            cell.set('r', address)
        cells[address] = cell
    for address,value in values.items():
        address, number = _cell_reference(address)
        if address not in cells:
            if value is None:continue
            if number not in rows:rows[number]=E.SubElement(data,X+'row',r=str(number))
            cells[address]=E.SubElement(rows[number],X+'c',r=address)
        cell=cells[address]
        formula=cell.find(X+'f')
        if formula is not None:
            raise ValueError('The requested value overlaps a calculation formula.')
        if cell.get('t')=='s':reference_delta-=1
        for tag in ('v','is'):
            for previous in cell.findall(X+tag):cell.remove(previous)
        cell.attrib.pop('t',None)
        if value is None:continue
        if type(value) in (int,float) and isfinite(value):
            E.SubElement(cell,X+'v').text=repr(value)
        elif isinstance(value,str) and shared_text and value in shared_text:
            cell.set('t','s')
            E.SubElement(cell,X+'v').text=str(shared_text[value])
            reference_delta+=1
        elif isinstance(value,str):
            cell.set('t','inlineStr')
            _write_text(E.SubElement(cell, X + 'is'), value)
        else:raise ValueError('The requested cell value has an unsupported type.')
    for row in rows.values():
        row[:]=sorted(row,key=lambda c:column_index_from_string(coordinate_from_string(c.get('r'))[0]))
        row.attrib.pop('spans',None)
    data[:]=[rows[number] for number in sorted(rows)]
    dimension=root.find(X+'dimension')
    if dimension is not None:
        last_col=max((column_index_from_string(coordinate_from_string(a)[0]) for a in cells), default=1)
        dimension.set('ref',f'A1:{get_column_letter(last_col)}{max(rows,default=1)}')
    return reference_delta


def write_values(source,destination,values,*,shared_strings=()):
    source,destination=Path(source),Path(destination)
    if source.resolve()==destination.resolve() or destination.exists():
        raise ValueError('Use a new destination for the workbook copy.')
    with ZipFile(source) as original:
        parts=sheet_parts(original)
        if values.keys()-parts.keys():raise ValueError('A requested worksheet is missing.')
        updates={}
        shared_part='xl/sharedStrings.xml'
        shared=None
        indexes={}
        delta=0
        if shared_strings and shared_part in original.namelist():
            shared=_read_xml_part(original, shared_part)
            entries=shared.findall(X+'si')
            for index,entry in enumerate(entries):
                if len(entry)==1 and entry[0].tag==X+'t' and entry[0].text in shared_strings:
                    indexes[entry[0].text]=index
            insert_at=len(entries)
            for value in sorted(set(shared_strings)-indexes.keys()):
                indexes[value]=insert_at
                entry=E.Element(X+'si')
                _write_text(entry, value)
                shared.insert(insert_at,entry)
                insert_at+=1
        for sheet,items in values.items():
            root=_read_xml_part(original, parts[sheet], worksheet=True)
            delta+=update_sheet(root,items,shared_text=indexes)
            updates[parts[sheet]]=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
        if shared is None and delta and shared_part in original.namelist():
            shared = _read_xml_part(original, shared_part)
        if shared is not None:
            if shared.get('count') is not None:
                shared.set('count',str(int(shared.get('count'))+delta))
            if shared.get('uniqueCount') is not None:
                shared.set('uniqueCount',str(len(shared.findall(X+'si'))))
            updates[shared_part]=E.tostring(shared,encoding='UTF-8',xml_declaration=True,standalone=True)
        destination.parent.mkdir(parents=True,exist_ok=True)
        with ZipFile(destination,'x',ZIP_DEFLATED,compresslevel=5) as output:
            for part in original.namelist():output.writestr(part,updates.get(part,original.read(part)))
