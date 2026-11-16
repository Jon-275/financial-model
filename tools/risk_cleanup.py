"""Move editable risk controls and remove the retired worksheet package parts."""

from pathlib import PurePosixPath
from posixpath import normpath

from lxml import etree as ET


LEGACY_SHEETS = ('Sequence_Risk', 'Monte_Carlo')
RISK_SHEET = 'Risk_Settings'
MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = '{' + REL_NS + '}'


def configure(model, baseline):
    """Carry the old workbook's report inputs into one short worksheet."""
    sequence, monte_carlo = (baseline[name] for name in LEGACY_SHEETS)
    put = lambda address, value, style: model.put(RISK_SHEET, address, value, style)
    put('A1', 'Risk settings', 'title')
    put('A2', 'Save and close your workbook after editing. Follow the market risk guide to refresh the results on Summary.', 'note')
    put('A4', 'Setting', 'header')
    put('B4', 'Value', 'header')
    put('C4', 'What it changes', 'header')
    settings = (
        ('Years of weak returns', 'Apply the stock returns below during these first retirement years.', 'input'),
        ('Balance that triggers cuts', 'Compare the actual portfolio with this share of the planned balance.', 'percent_input'),
        ('Spending cut when triggered', 'Cut this percentage of flexible spending when the balance falls below the trigger.', 'percent_input'),
        ('Share of spending you can cut', 'Enter the share of retirement living costs that you could reduce.', 'percent_input'),
        ('Investment return convention', 'Use 1 for compound growth or 2 for a simple annual average.', 'input'),
    )
    for row, (label, description, style) in enumerate(settings, 5):
        source = monte_carlo['B8'] if row == 9 else sequence[f'B{row}']
        put(f'A{row}', label, 'label')
        put(f'B{row}', source.value, style)
        put(f'C{row}', description, 'note')
    put('A12', 'Early retirement stock returns', 'band')
    put('A13', 'Test', 'header')
    put('B13', 'Annual return', 'header')
    put('C13', 'Enter returns in order from lowest to highest.', 'note')
    for offset, source in enumerate(sequence['G13:S13'][0], 14):
        put(f'A{offset}', offset - 13, 'whole')
        put(f'B{offset}', source.value, 'percent_input')


def _target(source, target):
    if target.startswith('/'):
        return target[1:]
    return normpath(str(PurePosixPath(source).parent / target))


def _linked_drawings(archive, available, part):
    """Find a sheet's drawing and chart parts, leaving unrelated assets alone."""
    file = PurePosixPath(part)
    relationships = str(file.parent / '_rels' / f'{file.name}.rels')
    linked = {part}
    if relationships not in available:
        return linked
    linked.add(relationships)
    for relationship in ET.fromstring(archive.read(relationships)):
        if relationship.get('Type') in (REL_NS + '/drawing', REL_NS + '/chart'):
            destination = _target(part, relationship.get('Target'))
            linked.update(_linked_drawings(archive, available, destination))
    return linked


def retire_legacy(archive, workbook, relationships, content_types):
    """Remove legacy sheet relationships, their charts and related workbook names."""
    available = set(archive.namelist())
    by_id = {relationship.get('Id'): relationship for relationship in relationships}
    sheets = workbook.find(MAIN + 'sheets')
    discard, retain = set(), set()
    for sheet in list(sheets):
        relationship = by_id[sheet.get(REL + 'id')]
        part = _target('xl/workbook.xml', relationship.get('Target'))
        linked = _linked_drawings(archive, available, part)
        if sheet.get('name') in LEGACY_SHEETS:
            discard.update(linked)
            sheets.remove(sheet)
            relationships.remove(relationship)
        else:
            retain.update(linked)
    discard.difference_update(retain)
    for override in list(content_types):
        if override.get('PartName', '').lstrip('/') in discard:
            content_types.remove(override)
    names = workbook.find(MAIN + 'definedNames')
    if names is not None:
        for name in list(names):
            formula = name.text or ''
            if any(f'{sheet}!' in formula or f"'{sheet}'!" in formula for sheet in LEGACY_SHEETS):
                names.remove(name)
    return discard
