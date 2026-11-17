"""Inspect protection and editability on every worksheet in a built release."""

from pathlib import Path
import sys
from zipfile import ZipFile

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, range_boundaries

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from workbook_protection import X, _entry_cells
from xlsx_values import sheet_parts


def check(archive):
    styles = ET.fromstring(archive.read('xl/styles.xml')).find(X + 'cellXfs')
    unlocked = set()
    for i, xf in enumerate(styles):
        options = xf.find(X + 'protection')
        if options is not None:
            if options.get('hidden') in ('1','true'):
                raise AssertionError(f'Style {i} hides formulas')
            if options.get('locked') in ('0','false'):
                unlocked.add(str(i))
    formula_count = edit_count = 0
    for name, path in sheet_parts(archive).items():
        root = ET.fromstring(archive.read(path))
        flags = root.find(X + 'sheetProtection')
        if flags is None or flags.get('sheet') not in ('1','true'):
            raise AssertionError(f'{name} is not protected')
        if flags.get('autoFilter', '1') not in ('0','false'):
            raise AssertionError(f'Filtering is disabled on {name}')
        if any(flags.get(key) in ('1','true') for key in ('selectLockedCells','selectUnlockedCells')):
            raise AssertionError(f'Selection is restricted on {name}')
        if any(key in flags.attrib for key in ('password','hashValue','saltValue','algorithmName','spinCount')):
            raise AssertionError(f'{name} requires a password to remove protection')
        expected = _entry_cells(name, root)
        for merge in root.iter(X + 'mergeCell'):
            ref = merge.get('ref')
            if ref.split(':')[0] in expected:
                left, top, right, bottom = range_boundaries(ref)
                expected.update(f'{get_column_letter(c)}{r}' for c in range(left, right + 1)
                                for r in range(top, bottom + 1))
        actual = {cell.get('r') for cell in root.iter(X + 'c') if cell.get('s','0') in unlocked}
        if expected != actual:
            raise AssertionError(f'{name} editable cells differ at {sorted(expected ^ actual)}')
        for formula in root.iter(X + 'f'):
            if formula.getparent().get('r') in actual:
                raise AssertionError(f"Editable formula at {name}!{formula.getparent().get('r')}")
            formula_count += 1
        if name == 'Inputs' and any(row.get('hidden') in ('1','true') for row in root.iter(X + 'row')):
            raise AssertionError('An Inputs row is hidden when the workbook opens')
        edit_count += len(actual)
    return formula_count, edit_count


if __name__ == '__main__':
    with ZipFile(sys.argv[1]) as archive:
        formulas, entries = check(archive)
    print(f'PASS {formulas:,} protected formula cells and {entries:,} editable cells.')
