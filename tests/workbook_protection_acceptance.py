"""Inspect protection and editability on every worksheet in a built release."""

from pathlib import Path
import re
import sys
from zipfile import ZipFile

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, range_boundaries

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
import investment_results
from start_here import LINKS as START_HERE_LINKS
from workbook_protection import X, _entry_cells
from xlsx_values import sheet_parts


# Check exact cell references so ordinary Inputs rows such as B170 do not match.
_RETIRED_REFERENCE = re.compile(
    r"(?<![\w.])(?:'?(?:Trial_Inputs|Inputs)'?)!\$?B\$?(?:" +
    '|'.join(str(row) for row in sorted(investment_results.RETIRED_ROWS)) + r")(?!\d)", re.I)
_ACCOUNT_BASE_RATES = tuple(investment_results.reference(row)
                            for row in (17, 18, 19, 20, 21, 22, 134, 135))


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
        if name == investment_results.SHEET:
            formulas = {formula.getparent().get('r') for formula in root.iter(X + 'f')}
            if not set(investment_results.CELLS.values()) <= formulas:
                raise AssertionError('Selected account returns and taxable yields must be formulas on Asset_Allocation.')
        if name == 'Start Here':
            links = {link.get('ref'): link.get('location') for link in root.iter(X + 'hyperlink')}
            if links != START_HERE_LINKS or next(root.iter(X + 'dataValidation'), None) is not None:
                raise AssertionError('Start Here must link directly to Inputs, Summary and Appendix without entry fields.')
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
            text = formula.text or ''
            if 'Inputs' in text and _RETIRED_REFERENCE.search(text):
                raise AssertionError(f"{name}!{formula.getparent().get('r')} still uses a retired input row")
            if name.startswith('Trial_') and 'Asset_Allocation!' in text and any(ref in text for ref in _ACCOUNT_BASE_RATES):
                raise AssertionError(f"{name}!{formula.getparent().get('r')} must use the active market path")
            formula_count += 1
        if name in ('Inputs', 'Trial_Inputs'):
            required = investment_results.RETIRED_ROWS - {134, 135}
            retired_rows = set()
            for row in root.iter(X + 'row'):
                number = int(row.get('r'))
                hidden = row.get('hidden') in ('1', 'true')
                retired = number in investment_results.RETIRED_ROWS
                if retired:
                    retired_rows.add(number)
                if retired and (not hidden or len(row) or float(row.get('ht', '1')) != 0):
                    raise AssertionError(f'Retired {name} row {number} must be empty and have zero height')
                if hidden and not retired:
                    raise AssertionError(f'{name} row {number} is hidden when the workbook opens')
            if not required <= retired_rows:
                raise AssertionError(f'{name} is missing zero-height rows in its retired sections')
        edit_count += len(actual)
    return formula_count, edit_count


if __name__ == '__main__':
    with ZipFile(sys.argv[1]) as archive:
        formulas, entries = check(archive)
    print(f'PASS {formulas:,} protected formula cells and {entries:,} editable cells.')
