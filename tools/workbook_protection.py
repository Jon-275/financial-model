"""Protect calculations from accidental edits while keeping formulas visible in Excel."""

from copy import deepcopy

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, range_boundaries

from plan_report import DOLLAR_VIEW_CELL
from risk_engine.scenario import (
    ASSET_ROWS, CORRELATION_ROWS, ONE_TIME_EXPENSE_ROWS, RISK_CONTROL_ROWS,
    RISK_RETURN_BASIS_ROW, RISK_STRESS_ROWS, SCENARIO_INPUT_ROWS, TAX_TABLE_LAST_ROW,
    asset_profile_start,
)

X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def _grid(columns, rows):
    return {f'{column}{row}' for row in rows for column in columns}


def _entry_cells(name, root):
    if name == 'Inputs':
        return ({'B4', 'B130', 'B131', 'B132'} | _grid('DEF', SCENARIO_INPUT_ROWS)
                | _grid('ABC', ONE_TIME_EXPENSE_ROWS))
    if name == 'Summary':
        return {DOLLAR_VIEW_CELL, 'D90'}
    if name == 'Asset_Allocation':
        weights = (row for profile in range(1, 4)
                   for row in range(asset_profile_start(profile), asset_profile_start(profile) + 5))
        return (_grid('BCDE', ASSET_ROWS) | _grid('BCDEF', CORRELATION_ROWS)
                | _grid('BCDE', weights) | {'A24', 'A33', 'A42'})
    if name == 'Income_Targeting':
        return {'B5', 'B21', 'B36'}
    if name == 'Risk_Settings':
        return _grid('B', (*RISK_CONTROL_ROWS, RISK_RETURN_BASIS_ROW, *RISK_STRESS_ROWS, *range(29, 37)))
    if name == 'Risk_Trials':
        return {'B3'}
    if name == 'Tax_Tables':
        return {cell.get('r') for row in root.find(X + 'sheetData') if int(row.get('r')) <= TAX_TABLE_LAST_ROW
                for cell in row if cell.get('t', 'n') == 'n' and cell.find(X + 'v') is not None
                and cell.find(X + 'f') is None}
    return set()


class WorksheetProtection:
    def __init__(self, styles):
        self.xfs = styles.find(X + 'cellXfs')
        if self.xfs is None or not len(self.xfs):
            raise ValueError('The workbook needs a cell format before its input cells can be unlocked.')
        self.variants = {}
        # Defaults are locked and visible. Enforce that contract if the template's styles change.
        for collection in (styles.find(X + 'cellStyleXfs'), self.xfs):
            if collection is None:
                continue
            for style in collection.findall(X + 'xf'):
                protection = style.find(X + 'protection')
                if protection is not None:
                    protection.set('locked', '1')
                    protection.set('hidden', '0')

    def _unlocked_style(self, source):
        if source not in self.variants:
            style = deepcopy(self.xfs[int(source)])
            style.set('applyProtection', '1')
            protection = style.find(X + 'protection')
            if protection is None:
                protection = ET.Element(X + 'protection')
                extension = style.find(X + 'extLst')
                if extension is None:
                    style.append(protection)
                else:
                    extension.addprevious(protection)
            protection.set('locked', '0')
            protection.set('hidden', '0')
            self.variants[source] = str(len(self.xfs))
            self.xfs.append(style)
            self.xfs.set('count', str(len(self.xfs)))
        return self.variants[source]

    def apply(self, name, root):
        data = root.find(X + 'sheetData')
        entries = _entry_cells(name, root)
        if entries:
            # Excel treats a merged entry as one field. Unlock every component, including blank cells.
            for merge in root.iter(X + 'mergeCell'):
                reference = merge.get('ref')
                if reference.split(':')[0] in entries:
                    left, top, right, bottom = range_boundaries(reference)
                    entries.update(_grid([get_column_letter(col) for col in range(left, right + 1)],
                                         range(top, bottom + 1)))
            rows = {int(row.get('r')): row for row in data}
            for address in sorted(entries):
                column, row_number, _, _ = range_boundaries(address)
                row = rows.get(row_number)
                if row is None:
                    raise ValueError(f'Missing input row {name}!{row_number}')
                cell = row.find(f'{X}c[@r="{address}"]')
                if cell is None:
                    cell = ET.Element(X + 'c', r=address)
                    following = next((other for other in row if range_boundaries(other.get('r'))[0] > column), None)
                    if following is None:
                        row.append(cell)
                    else:
                        following.addprevious(cell)
                if cell.find(X + 'f') is not None:
                    raise ValueError(f'Formula overlaps an editable field at {name}!{address}')
                cell.set('s', self._unlocked_style(cell.get('s', '0')))

        previous = root.find(X + 'sheetProtection')
        if previous is not None:
            root.remove(previous)
        protection = ET.Element(X + 'sheetProtection', sheet='1', objects='1', scenarios='1',
                                selectLockedCells='0', selectUnlockedCells='0', autoFilter='0',
                                formatRows='0', formatColumns='0')
        # sheetCalcPr, when present, precedes sheetProtection in the worksheet schema.
        anchor = root.find(X + 'sheetCalcPr')
        (data if anchor is None else anchor).addnext(protection)
