"""Remove inherited document metadata from the release workbook package."""

from lxml import etree
from workbook_privacy import DISCARDED_PARTS


MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_WORKBOOK_METADATA = (
    '{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}revisionPtr',
    MAIN + 'fileSharing',
)
_EXTENDED = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
_REVISION_NAMESPACES = {
    'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    'xr2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
}
_SHARED_STRINGS = 'xl/sharedStrings.xml'


def _pack(root):
    return etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def make_sheets_visible(workbook):
    """Use Excel's default visible state for every worksheet."""
    for sheet in workbook.find(MAIN + 'sheets'):
        sheet.attrib.pop('state', None)
    for view in workbook.findall(f'{MAIN}bookViews/{MAIN}workbookView'):
        view.attrib.pop('showSheetTabs', None)


def clear_revision_ids(root):
    """Discard coauthoring IDs inherited from Excel saves."""
    nodes = root.xpath('descendant-or-self::*[@xr:uid or @xr2:uid]', namespaces=_REVISION_NAMESPACES)
    for node in nodes:
        for namespace in _REVISION_NAMESPACES.values():
            node.attrib.pop(f'{{{namespace}}}uid', None)


def clean_metadata(archive, workbook, app, content_types, workbook_rels, changes):
    """Drop the known nonfunctional metadata and its package references."""
    clear_revision_ids(workbook)
    for tag in _WORKBOOK_METADATA:
        for element in workbook.findall(tag):
            workbook.remove(element)
    for element in list(app):
        if element.tag not in (_EXTENDED + 'HeadingPairs', _EXTENDED + 'TitlesOfParts'):
            app.remove(element)

    core = etree.fromstring(archive.read('docProps/core.xml'))
    for element in list(core):
        core.remove(element)
    changes['docProps/core.xml'] = _pack(core)

    package_rels = etree.fromstring(archive.read('_rels/.rels'))
    for manifest in (content_types, package_rels, workbook_rels):
        for entry in list(manifest):
            target = (entry.get('PartName') or entry.get('Target') or '').lstrip('/')
            if manifest is workbook_rels and target == 'calcChain.xml':
                target = 'xl/calcChain.xml'
            if target in DISCARDED_PARTS:
                manifest.remove(entry)
    changes['_rels/.rels'] = _pack(package_rels)
    return DISCARDED_PARTS


def compact_shared_strings(archive, worksheet_parts, changes):
    """Drop strings left behind by edits and retired sheets, preserving used entries."""
    if _SHARED_STRINGS not in archive.namelist():
        return
    strings = etree.fromstring(archive.read(_SHARED_STRINGS))
    worksheets = []
    used = set()
    count = 0
    for part in worksheet_parts:
        data = changes[part] if part in changes else archive.read(part)
        if b't="s"' not in data and b"t='s'" not in data:
            continue
        root = etree.fromstring(data)
        values = root.findall(f'.//{MAIN}c[@t="s"]/{MAIN}v')
        if values:
            worksheets.append((part, root, values))
            used.update(int(value.text) for value in values)
            count += len(values)

    indices = {old: new for new, old in enumerate(sorted(used))}
    if len(indices) != len(strings):
        strings[:] = [strings[index] for index in indices]
        for part, root, values in worksheets:
            for value in values:
                value.text = str(indices[int(value.text)])
            changes[part] = _pack(root)
    strings.set('count', str(count))
    strings.set('uniqueCount', str(len(indices)))
    changes[_SHARED_STRINGS] = _pack(strings)
