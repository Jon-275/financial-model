"""Install visible range contracts and the on-demand run control."""
from lxml import etree as E

X='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'


def install_names(workbook,ranges):
    names=workbook.find(X+'definedNames')
    if names is None:
        names = E.Element(X+'definedNames')
        preceding = {X+tag for tag in ('fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection',
                                       'bookViews', 'sheets', 'functionGroups', 'externalReferences')}
        position = 1 + max((i for i, child in enumerate(workbook) if child.tag in preceding), default=-1)
        workbook.insert(position, names)
    sheets={s.get('name'):str(i) for i,s in enumerate(workbook.find(X+'sheets'))}
    for sheet,items in ranges.items():
        scope=sheets[sheet]
        for key,address in items.items():
            for old in list(names):
                if old.get('name')==key and old.get('localSheetId')==scope:names.remove(old)
            E.SubElement(names,X+'definedName',name=key,localSheetId=scope).text=(
                address if '!' in address else f"'{sheet}'!{address}")


def _button(drawing, identifier, label, macro, description, offset, width, fill, ink):
    anchor=E.SubElement(drawing,D+'oneCellAnchor')
    start=E.SubElement(anchor,D+'from')
    for tag,value in (('col',10),('colOff',offset*9525),('row',19),('rowOff',0)):
        E.SubElement(start,D+tag).text=str(value)
    E.SubElement(anchor,D+'ext',cx=str(width*9525),cy=str(34*9525))
    shape=E.SubElement(anchor,D+'sp',macro='[0]!'+macro)
    nonvisual=E.SubElement(shape,D+'nvSpPr')
    E.SubElement(nonvisual,D+'cNvPr',id=str(identifier),name=label,descr=description)
    E.SubElement(nonvisual,D+'cNvSpPr')
    properties=E.SubElement(shape,D+'spPr')
    geometry=E.SubElement(properties,A+'prstGeom',prst='roundRect')
    E.SubElement(geometry,A+'avLst')
    E.SubElement(E.SubElement(properties,A+'solidFill'),A+'srgbClr',val=fill)
    E.SubElement(E.SubElement(properties,A+'ln',w='9525'),A+'noFill')
    text=E.SubElement(shape,D+'txBody')
    E.SubElement(text,A+'bodyPr',anchor='ctr',wrap='none',lIns='12700',rIns='12700')
    E.SubElement(text,A+'lstStyle')
    paragraph=E.SubElement(text,A+'p')
    E.SubElement(paragraph,A+'pPr',algn='ctr')
    run=E.SubElement(paragraph,A+'r')
    font=E.SubElement(run,A+'rPr',lang='en-US',sz='1000',b='1')
    E.SubElement(E.SubElement(font,A+'solidFill'),A+'srgbClr',val=ink)
    E.SubElement(font,A+'latin',typeface='Arial')
    E.SubElement(run,A+'t').text=label
    E.SubElement(paragraph,A+'endParaRPr',lang='en-US',sz='1100')
    E.SubElement(anchor,D+'clientData',fPrintsWithSheet='0')


def install_buttons(drawing):
    """Place accessible Run and Cancel controls together outside the Summary print area."""
    for anchor in list(drawing):
        if any(item.get('name') in ('Run market risk', 'Cancel') for item in anchor.iter(D+'cNvPr')):
            drawing.remove(anchor)
    identifier=1+max((int(item.get('id','0')) for item in drawing.iter(D+'cNvPr')),default=0)
    _button(drawing, identifier, 'Run market risk', 'RunMarketRisk',
            'Run the selected market paths and stress tests, then update the report.',
            0, 126, '173B4B', 'FFFFFF')
    _button(drawing, identifier+1, 'Cancel', 'CancelMarketRisk',
            'Stop the current run and keep the last completed batch.',
            132, 65, 'E3ECEC', '173B4B')


def install(workbook,changes,ranges):
    if not ranges:return
    install_names(workbook,ranges)
    part='xl/drawings/drawing1.xml'
    drawing=E.fromstring(changes[part])
    install_buttons(drawing)
    changes[part]=E.tostring(drawing,encoding='UTF-8',xml_declaration=True,standalone=True)
