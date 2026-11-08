"""Read a Phase 2 workbook without editing it or trusting uncached calculations."""

from __future__ import annotations

import hashlib
import json
import math
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from types import MappingProxyType
from typing import Mapping
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException


Cell = int | float | str | None
Cells = Mapping[str, Cell]
Matrix = tuple[tuple[float, ...], ...]
_OPENING_YEAR = 2026
_FIRST_ROW = 5
_LAST_ROW = 130
_INPUT_ROWS = (*range(7, 17), *range(25, 63), 67, 68, 69,
               *range(86, 99), *range(101, 125), 126, 127, 128)
_CACHED_INPUT_ROWS = (*range(17, 25), *range(63, 67), 70, 71, 99, 125, 129)
_MODEL_RANGES = {'Projection': 84, 'Cash_Flow': 29, 'Healthcare': 32,
                 'Roth_Access': 42, 'Checks': 13}
_MODEL_TEXT = {'Projection': frozenset(('E', 'F')), 'Cash_Flow': frozenset(('L',)),
               'Healthcare': frozenset(('D', 'E', 'T', 'Z')), 'Roth_Access': frozenset()}


class ScenarioError(ValueError):
    """Workbook validation failed. Messages name locations, never input values."""


@dataclass(frozen=True, slots=True)
class Allocation:
    asset_names: tuple[str, ...]
    means: tuple[float, ...]
    volatilities: tuple[float, ...]
    ordinary_yields: tuple[float, ...]
    qualified_yields: tuple[float, ...]
    correlation: Matrix
    # Account rows are taxable, traditional, Roth and HSA, in that order.
    working: Matrix
    retirement: Matrix


@dataclass(frozen=True, slots=True)
class YearSnapshot:
    row: int
    year: int
    projection: Cells
    cash: Cells
    health: Cells
    roth: Cells


@dataclass(frozen=True, slots=True)
class Scenario:
    selected: int
    inputs: Mapping[int, Cell]
    tax: Cells
    allocation: Allocation
    sequence: Mapping[int, float]
    stress_returns: tuple[float, ...]
    return_basis: int
    opening_projection: Cells
    opening_cash: Cells
    opening_health: Cells
    opening_roth: Cells
    years: tuple[YearSnapshot, ...]
    digest: str
    workbook_sha256: str

    @property
    def tax_tables(self) -> Cells:
        return self.tax


def _invalid(location: str, reason: str) -> None:
    raise ScenarioError(f'{location}: {reason}.')


def _number(value: Cell, location: str) -> float:
    if type(value) not in (int, float) or not math.isfinite(value):
        _invalid(location, 'expected a finite number')
    return float(value)


def _integer(value: Cell, location: str, low: int, high: int) -> int:
    number = _number(value, location)
    if not number.is_integer() or not low <= number <= high:
        _invalid(location, 'integer outside the supported range')
    return int(number)


def _matches(actual: Cell, expected: Cell, location: str) -> None:
    if type(expected) in (int, float):
        if math.isclose(_number(actual, location), expected, rel_tol=2e-12, abs_tol=1e-7):
            return
    elif actual == expected:
        return
    _invalid(location, 'saved calculation differs from the active assumptions; recalculate the workbook')


def _cell(cell, location: str) -> Cell:
    if cell.data_type in ('e', 'f'):
        _invalid(location, 'formula error or expected a literal value')
    value = cell.value
    if value is not None and type(value) not in (int, float, str):
        _invalid(location, 'unsupported cell type')
    if type(value) is float and not math.isfinite(value):
        _invalid(location, 'expected a finite number')
    return value


def _grid(sheet, min_row: int, max_row: int, max_column: int) -> dict[int, dict[str, Cell]]:
    columns = [get_column_letter(i) for i in range(1, max_column + 1)]
    return {
        row: {col: _cell(cell, f'{sheet.title}!{col}{row}') for col, cell in zip(columns, cells)}
        for row, cells in enumerate(
            sheet.iter_rows(min_row=min_row, max_row=max_row, max_col=max_column), min_row
        )
    }


def _one_offs(raw) -> dict[int, float]:
    entries: dict[int, float] = {}
    for row, (_, year_cell, amount_cell) in enumerate(
        raw['Inputs'].iter_rows(min_row=75, max_row=82, max_col=3), 75
    ):
        year = _cell(year_cell, f'Inputs!B{row}')
        amount = _cell(amount_cell, f'Inputs!C{row}')
        if year in (None, '') and amount in (None, '', 0):
            continue
        year = _integer(year, f'Inputs!B{row}', 1900, 9999)
        amount = _number(amount, f'Inputs!C{row}') if amount not in (None, '') else 0.0
        if amount < 0:
            _invalid(f'Inputs!C{row}', 'expected a nonnegative expense')
        entries[year] = entries.get(year, 0.0) + amount
    return entries


def _inputs(raw, cached) -> tuple[int, dict[int, Cell], dict[int, float]]:
    source = list(raw['Inputs'].iter_rows(min_row=4, max_row=129, max_col=6))
    selected = _integer(_cell(source[0][1], 'Inputs!B4'), 'Inputs!B4', 1, 3)
    data = _grid(cached['Inputs'], 4, 129, 2)
    selected_col = get_column_letter(selected + 3)
    detailed_cell = source[101 - 4][selected + 2]
    detailed = _integer(_cell(detailed_cell, f'Inputs!{selected_col}101'), 'Inputs!B101', 0, 1)
    inputs = {4: selected}
    for row in _INPUT_ROWS:
        location = f'Inputs!{selected_col}{row}'
        value = _cell(source[row - 4][selected + 2], location)
        if row == 7:
            if not isinstance(value, str) or not value.strip():
                _invalid(location, 'expected a scenario name')
        elif value is None and not detailed and 102 <= row <= 124:
            # INDEX returns zero for an unused blank legacy healthcare input.
            value = 0
        else:
            _number(value, location)
        _matches(data[row]['B'], value, f'Inputs!B{row}')
        inputs[row] = value

    for row in _CACHED_INPUT_ROWS:
        value = data[row]['B']
        if row in (71, 99, 125, 129):
            if value != 'OK':
                _invalid(f'Inputs!B{row}', 'scenario validation has not passed')
        else:
            _number(value, f'Inputs!B{row}')
        inputs[row] = value

    _validate_assumptions(inputs)
    return selected, inputs, _one_offs(raw)


def _validate_assumptions(inputs: Mapping[int, Cell]) -> None:
    def integer(row, low, high):
        return _integer(inputs[row], f'Inputs!B{row}', low, high)

    start = integer(8, 0, 125)
    end = integer(53, start, 125)
    _matches(inputs[70], end - start + 1, 'Inputs!B70')
    for row in (9, 10, 45, 50, 67, 87, 88):
        integer(row, 0, 125)
    for row in (54, 58):
        if inputs[row] != 999:
            integer(row, 0, 125)
    for row in (11, 28, 48, 97):
        integer(row, 1, 2)
    for row in (15, 16):
        integer(row, 1, 3)
    for row in (29, 92, 93, 94):
        integer(row, 0, _OPENING_YEAR)
    for row in (95, 96):
        integer(row, 0, 126)
    for row in (68, 98):
        integer(row, 0, 1)
    integer(39, 0, 9999)
    integer(40, 0, 12)
    integer(60, 0, 125)

    for row in (12, *range(25, 28), *range(30, 39), 41, 42, 43, 44,
                46, 47, 49, 51, 52, 55, 56, 57, 59, 62, 127, 128):
        if inputs[row] < 0:
            _invalid(f'Inputs!B{row}', 'expected a nonnegative amount')
    for row in (61, 69, 86, 89, 90, 91, 126):
        if not 0 <= inputs[row] <= 1:
            _invalid(f'Inputs!B{row}', 'expected a fraction from zero to one')
    if inputs[14] <= -1 or inputs[13] < -1:
        _invalid('Inputs!B13:B14', 'wage growth or inflation is outside the supported range')
    if inputs[87] > inputs[88]:
        _invalid('Inputs!B87:B88', 'outside income ends before it starts')
    if inputs[28] == 1 and inputs[30] > inputs[27]:
        _invalid('Inputs!B27:B30', 'accessible Roth IRA basis exceeds its opening balance')
    if inputs[25] == 0 and inputs[51] > 0:
        _invalid('Inputs!B25/B51', 'taxable basis requires an opening taxable account')


def _allocation(raw, inputs) -> Allocation:
    sheet = raw['Asset_Allocation']
    rows = _grid(sheet, 6, 10, 5) | _grid(sheet, 15, 20, 6)
    for number in {int(inputs[15]), int(inputs[16])}:
        start = 26 + 9 * (number - 1)
        rows.update(_grid(sheet, start, start + 4, 5))

    def numbers(col, indexes):
        return tuple(_number(rows[row][col], f'Asset_Allocation!{col}{row}') for row in indexes)

    asset_rows = range(6, 11)
    names = tuple(rows[row]['A'] for row in asset_rows)
    if any(not isinstance(name, str) or not name.strip() for name in names) or len(set(names)) != 5:
        _invalid('Asset_Allocation!A6:A10', 'expected five distinct asset labels')
    means, vols, ordinary, qualified = (numbers(col, asset_rows) for col in 'BCDE')
    if min(means) <= -1 or min(*vols, *ordinary, *qualified) < 0:
        _invalid('Asset_Allocation!B6:E10', 'return, volatility or yield is outside the supported range')
    corr = tuple(tuple(numbers(col, (row,))[0] for col in 'BCDEF') for row in range(16, 21))
    if any(rows[row]['A'] != name or rows[15][col] != name
           for row, col, name in zip(range(16, 21), 'BCDEF', names)):
        _invalid('Asset_Allocation!A15:F20', 'correlation labels differ from the asset order')
    for i, values in enumerate(corr):
        if not math.isclose(values[i], 1, abs_tol=1e-10) or any(
            abs(value) > 1 or not math.isclose(value, corr[j][i], abs_tol=1e-10)
            for j, value in enumerate(values)
        ):
            _invalid('Asset_Allocation!B16:F20', 'correlations must be symmetric with unit diagonals')

    def profile(number):
        start = 26 + 9 * (int(number) - 1)
        profile_rows = range(start, start + 5)
        if tuple(rows[row]['A'] for row in profile_rows) != names:
            _invalid('Asset_Allocation', 'profile labels differ from the asset order')
        weights = tuple(numbers(col, profile_rows) for col in 'BCDE')
        for col, account in zip('BCDE', weights):
            if min(account) < 0 or not math.isclose(sum(account), 1, abs_tol=1e-9):
                _invalid(f'Asset_Allocation!{col}{start}:{col}{start + 4}',
                         'allocation weights must be nonnegative and total 100%')
        return weights

    working = profile(inputs[15])
    retirement = profile(inputs[16])
    for weights, return_rows, qualified_row, ordinary_row in (
        (working, (17, 19, 21), 63, 65), (retirement, (18, 20, 22), 64, 66)
    ):
        for row, weight in zip(return_rows, weights):
            _matches(inputs[row], math.fsum(x * y for x, y in zip(weight, means)), f'Inputs!B{row}')
        for row, yields in ((qualified_row, qualified), (ordinary_row, ordinary)):
            _matches(inputs[row], math.fsum(x * y for x, y in zip(weights[0], yields)), f'Inputs!B{row}')
    return Allocation(names, means, vols, ordinary, qualified, corr, working, retirement)


def _tax_tables(raw, cached) -> tuple[dict[str, Cell], dict[str, Cell]]:
    rows = _grid(cached['Tax_Tables'], 1, 252, 8)
    cached_values = {f'{col}{row}': value for row, cells in rows.items() for col, value in cells.items()
                     if value not in (None, '')}
    literal_numbers = {}
    for row, cells in enumerate(raw['Tax_Tables'].iter_rows(min_row=1, max_row=252, max_col=8), 1):
        for col, cell in enumerate(cells, 1):
            loc = f'{get_column_letter(col)}{row}'
            if cell.data_type == 'f':
                _number(cached_values.get(loc), f'Tax_Tables!{loc}')
            elif type(cell.value) in (int, float):
                literal_numbers[loc] = _cell(cell, f'Tax_Tables!{loc}')

    for start, length in ((6, 7), (17, 3), (30, 9), (144, 7), (154, 3), (160, 9)):
        prior_floor = 0.0
        for row in range(start, start + length):
            floor = _number(cached_values.get(f'B{row}'), f'Tax_Tables!B{row}')
            _matches(cached_values.get(f'D{row}'), floor - prior_floor, f'Tax_Tables!D{row}')
            prior_floor = floor
    for row in range(89, 95):
        amounts = [_number(cached_values.get(f'{col}{row}'), f'Tax_Tables!{col}{row}') for col in 'CD']
        _matches(cached_values.get(f'E{row}'), sum(amounts) * 12, f'Tax_Tables!E{row}')
    return cached_values, literal_numbers


def _validate_year(row, p, cash, health, roth, checks, inputs, one_offs, real_wage_growth) -> None:
    step = row - _FIRST_ROW
    year = _OPENING_YEAR + step
    age, spouse = inputs[8] + step, inputs[9] + step
    index = (1 + inputs[14]) ** step
    working = age < inputs[10]
    single = step > 0 and age > inputs[54]
    filing = 'Single' if single else 'MFJ'
    phase = 'Start' if step == 0 else ('Working' if working else
                                       'Bridge' if age < 60 else
                                       'Pre-Medicare' if age < 65 else 'Medicare')
    projection_expected = dict(A=step, B=year, C=age, D=spouse, E=phase, F=filing, G=index, H=index)
    if step == 0:
        projection_expected.update(BC=inputs[25], BD=inputs[51], BE=inputs[26], BF=inputs[27],
                                   BG=inputs[31], BW=inputs[127], BY=inputs[128])
        split = inputs[90]
        ira = inputs[27] if inputs[28] == 1 else 0
        plan = inputs[27] - ira
        basis = inputs[30] if inputs[28] == 1 else 0
        for col, expected in dict(D=ira * split, E=ira * (1 - split), F=plan * split,
                                  G=plan * (1 - split), H=basis * split, I=basis * (1 - split),
                                  N=inputs[29], O=inputs[92]).items():
            _matches(roth[col], expected, f'Roth_Access!{col}{row}')
    else:
        wage_index = index if real_wage_growth else 1
        wages = inputs[12] * wage_index * (1 + inputs[13]) ** step if working else 0
        primary = wages if single or inputs[11] == 1 else wages * inputs[126]
        deferral = min(wages, inputs[32] * index) if working else 0
        owner = int(inputs[121]) if inputs[101] == 1 else 1
        owner_age = age if owner == 1 else spouse
        medicare_age = inputs[104] if owner == 1 else inputs[105]
        hsa_eligible = inputs[101] == 0 or ((owner == 1 or not single)
                                            and (medicare_age == 126 or owner_age < medicare_age))
        hsa = min(wages - deferral, inputs[34] * index) if working and hsa_eligible else 0
        social_security = (inputs[55] if single else inputs[44]) * index if age >= inputs[45] else 0
        other = inputs[52] * index if inputs[87] <= age <= inputs[88] else 0
        exempt = inputs[46] * index
        fixed = inputs[38] if year < inputs[39] else 0
        if year == inputs[39]:
            fixed = inputs[38] * inputs[40] / 12
        spending = max(0, inputs[36 if working else 37] - inputs[38]) * index + fixed
        spending *= inputs[56] if single else 1
        care = inputs[59] * index if inputs[58] <= age < inputs[58] + inputs[60] else 0
        projection_expected.update(I=wages, J=deferral, K=hsa, L=inputs[33] * index if working else 0,
                                   N=social_security, O=other, AG=exempt, AW=spending,
                                   AX=one_offs.get(year, 0) * index, AY=care, BT=primary, BU=wages - primary)
        _matches(cash['AB'], wages - deferral - hsa + social_security + other * inputs[86]
                 + exempt * inputs[89], f'Cash_Flow!AB{row}')

    for col, expected in projection_expected.items():
        _matches(p[col], expected, f'Projection!{col}{row}')
    for sheet, cells in (('Cash_Flow', cash), ('Healthcare', health), ('Roth_Access', roth), ('Checks', checks)):
        _matches(cells['A'], year, f'{sheet}!A{row}')
        _matches(cells['B'], age, f'{sheet}!B{row}')
    for sheet, cells in (('Healthcare', health), ('Roth_Access', roth)):
        _matches(cells['C'], spouse, f'{sheet}!C{row}')
    if checks['J'] != 'OK':
        _invalid(f'Checks!J{row}', 'annual workbook reconciliation has not passed')
    if health['T'] not in ('Opening balances', 'Legacy budget', 'OK'):
        _invalid(f'Healthcare!T{row}', 'healthcare assumptions need review')


def _model_rows(cached, inputs, one_offs, *, real_wage_growth) -> tuple[YearSnapshot, ...]:
    ranges = {sheet: _grid(cached[sheet], _FIRST_ROW, _LAST_ROW, columns)
              for sheet, columns in _MODEL_RANGES.items()}
    active_last = int(inputs[70]) + _FIRST_ROW - 1
    snapshots = []
    for row in range(_FIRST_ROW, _LAST_ROW + 1):
        p, cash, health, roth, checks = (rows[row] for rows in ranges.values())
        if row > active_last:
            for sheet, values in zip(_MODEL_RANGES, (p, cash, health, roth, checks)):
                if any(value not in (None, '') for value in values.values()):
                    _invalid(f'{sheet}!{row}', 'saved results extend beyond the active horizon')
            continue
        for (sheet, text_columns), cells in zip(_MODEL_TEXT.items(), (p, cash, health, roth)):
            for col, value in cells.items():
                location = f'{sheet}!{col}{row}'
                if sheet == 'Healthcare' and col == 'Z' and value in (None, ''):
                    continue
                if value in (None, ''):
                    _invalid(location, 'missing saved calculation; recalculate the workbook')
                if col in text_columns:
                    if not isinstance(value, str):
                        _invalid(location, 'expected saved text')
                else:
                    _number(value, location)
        try:
            _validate_year(row, p, cash, health, roth, checks, inputs, one_offs, real_wage_growth)
        except OverflowError:
            _invalid(f'Projection!{row}', 'assumptions exceed the supported numeric range')
        snapshots.append(YearSnapshot(row, int(p['B']), *(MappingProxyType(x) for x in (p, cash, health, roth))))
    return tuple(snapshots)


def _risk_inputs(raw) -> tuple[dict[int, float], tuple[float, ...], int]:
    current_layout = 'Risk_Settings' in raw.sheetnames
    sequence_sheet = raw['Risk_Settings' if current_layout else 'Sequence_Risk']
    name = sequence_sheet.title
    settings = _grid(sequence_sheet, 5, 8, 2)
    sequence = {row: _number(settings[row]['B'], f'{name}!B{row}') for row in range(5, 9)}
    sequence[5] = _integer(sequence[5], f'{name}!B5', 0, 125)
    for row in range(6, 9):
        if not 0 <= sequence[row] <= 1:
            _invalid(f'{name}!B{row}', 'expected a fraction from zero to one')
    if current_layout:
        returns = _grid(sequence_sheet, 14, 26, 2)
        stress_returns = tuple(_number(returns[row]['B'], f'{name}!B{row}') for row in range(14, 27))
        stress_range = f'{name}!B14:B26'
        basis_sheet, basis_row = sequence_sheet, 9
    else:
        headings = _grid(sequence_sheet, 13, 13, 19)[13]
        columns = [get_column_letter(i) for i in range(7, 20)]
        stress_returns = tuple(_number(headings[col], f'{name}!{col}13') for col in columns)
        stress_range = f'{name}!G13:S13'
        basis_sheet, basis_row = raw['Monte_Carlo'], 8
    if min(stress_returns) < -1 or sorted(set(stress_returns)) != list(stress_returns):
        _invalid(stress_range, 'expected strictly increasing returns of at least minus 100%')
    basis = _grid(basis_sheet, basis_row, basis_row, 2)[basis_row]['B']
    return sequence, stress_returns, _integer(basis, f'{basis_sheet.title}!B{basis_row}', 1, 2)


def _digest(inputs, one_offs, tax, allocation, sequence, stress_returns, return_basis, *, real_wage_growth) -> str:
    # Numeric cells are canonicalized across engines that distinguish 1 and 1.0.
    payload = {
        'schema': 'risk-scenario-v1',
        'inputs': {row: float(inputs[row]) for row in (4, *_INPUT_ROWS) if row != 7},
        'one_offs': one_offs,
        'tax': {coordinate: float(value) for coordinate, value in tax.items()},
        'allocation': [allocation.means, allocation.volatilities, allocation.ordinary_yields,
                       allocation.qualified_yields, allocation.correlation, allocation.working,
                       allocation.retirement],
        'sequence': [sequence, stress_returns],
        'return_basis': return_basis,
    }
    if not real_wage_growth:
        payload['wage_growth_includes_inflation'] = True
    canonical = json.dumps(payload, sort_keys=True, separators=(',', ':'), allow_nan=False)
    return hashlib.sha256(canonical.encode('utf-8')).hexdigest()


def load_scenario(path: str | Path) -> Scenario:
    """Validate active inputs, cached model rows and provenance from one file snapshot."""
    try:
        workbook = Path(path).read_bytes()
        with BytesIO(workbook) as source:
            raw = load_workbook(source, read_only=True, keep_links=False)
            try:
                with BytesIO(workbook) as calculated_source:
                    cached = load_workbook(calculated_source, read_only=True, data_only=True, keep_links=False)
                    try:
                        return _read(raw, cached, hashlib.sha256(workbook).hexdigest())
                    finally:
                        cached.close()
            finally:
                raw.close()
    except ScenarioError:
        raise
    except (BadZipFile, InvalidFileException, OSError, KeyError, ParseError, ValueError, OverflowError):
        raise ScenarioError('Could not read the required Phase 2 workbook sheets.') from None


def _read(raw, cached, workbook_sha256: str) -> Scenario:
    # Stable model headers guard fixed-coordinate reads against moved or replaced sheets.
    for sheet, cell, expected in (
        ('Projection', 'CF4', 'Taxable basis added'), ('Roth_Access', 'AP4', 'Roth net transfer'),
        ('Healthcare', 'AF4', 'Healthcare paid from cash'), ('Cash_Flow', 'AC4', 'ACA premium credit'),
        ('Checks', 'M4', 'Taxable basis reconciliation'), ('Inputs', 'A129', 'Tax inputs'),
        ('Tax_Tables', 'A230', 'PUBLISHED TAX-YEAR AMOUNTS'),
    ):
        if raw[sheet][cell].value != expected:
            _invalid(f'{sheet}!{cell}', 'workbook layout is incompatible with this risk engine')
    selected, inputs, one_offs = _inputs(raw, cached)
    wage_label = raw['Inputs']['A13'].value
    if wage_label not in ('Real wage growth', 'Annual wage growth'):
        _invalid('Inputs!A13', 'unknown wage growth convention')
    real_wage_growth = wage_label == 'Real wage growth'
    allocation = _allocation(raw, inputs)
    sequence, stress_returns, return_basis = _risk_inputs(raw)
    tax, tax_source = _tax_tables(raw, cached)
    if not real_wage_growth and not 0 <= _number(tax_source.get('B251'), 'Tax_Tables!B251') <= 1:
        _invalid('Tax_Tables!B251', 'expected a fraction from zero to one')
    opening, *years = _model_rows(cached, inputs, one_offs, real_wage_growth=real_wage_growth)
    return Scenario(selected, MappingProxyType(inputs), MappingProxyType(tax), allocation,
                    MappingProxyType(sequence), stress_returns, return_basis,
                    opening.projection, opening.cash, opening.health, opening.roth, tuple(years),
                    _digest(inputs, one_offs, tax_source, allocation, sequence, stress_returns, return_basis,
                            real_wage_growth=real_wage_growth),
                    workbook_sha256)
