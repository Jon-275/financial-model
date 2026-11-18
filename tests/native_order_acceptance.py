"""Verify the VBA calculation sequence against the published worksheet formulas."""
import argparse
from collections import defaultdict
from dataclasses import dataclass
from graphlib import CycleError, TopologicalSorter
from pathlib import Path
import re
import sys
from zipfile import ZipFile

from lxml import etree
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import column_index_from_string as number, get_column_letter as letter, range_boundaries

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT / 'tools'))

from funding_search import ROWS
import native_capture
import native_funding_paths
from native_capture import ANNUAL_CAPACITY
from xlsx_values import X, sheet_parts


MARKET, SOLVER = 'Market_Paths', 'Trial_Funding_Calc'
SHEETS = {'marketYear': MARKET, 'projectionYear': 'Trial_Projection', 'cashYear': 'Trial_Cash_Flow',
          'rothYear': 'Trial_Roth_Access', 'healthYear': 'Trial_Healthcare', 'checksYear': 'Trial_Checks'}
STABLE = {'Inputs', 'Risk_Settings', 'Asset_Allocation', 'Tax_Tables', 'Projection'}
REFERENCE = re.compile(r"(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?")
ADDRESS = re.compile(r'([A-Z]+)(\d+)')
FIRST, LAST = 5, 4 + ANNUAL_CAPACITY


def require(ok, message):
    if not ok:
        raise ValueError(message)


@dataclass(frozen=True)
class Reference:
    sheet: str
    left: int
    top: int
    right: int
    bottom: int
    absolute: bool

    def inside(self, regions):
        for sheet, (left, top, right, bottom) in regions:
            if (sheet == self.sheet and left <= self.left <= self.right <= right
                    and top <= self.top <= self.bottom <= bottom):
                return True
        return False

    def is_cell(self, sheet, col, row):
        return (self.sheet == sheet and self.left == self.right == col
                and self.top == self.bottom == row)


def references(formula, owner):
    """Include both the declared index range and any end point after INDEX(...):."""
    for token in Tokenizer('=' + formula).items:
        if token.type == 'FUNC' and token.subtype == 'OPEN':
            require(token.value.upper() not in ('INDIRECT(', 'OFFSET('), f'Unsupported dynamic address in {owner}')
        if token.type != 'OPERAND' or token.subtype != 'RANGE':
            continue
        m = REFERENCE.fullmatch(token.value.removeprefix(':'))
        require(m is not None, f'Unclassified reference {token.value} in {owner}')
        qualified, bare, ac, col, ar, row, bc, end_col, br, end_row = m.groups()
        yield Reference(qualified or bare or owner, number(col), int(row), number(end_col or col),
                        int(end_row or row), bool(ac and ar and (not end_col or bc and br)))


def cells(archive, path):
    for _, row in etree.iterparse(archive.open(path), events=('end',), tag=X+'row'):
        for cell in row.iterchildren(X+'c'):
            f = cell.find(X+'f')
            if f is not None:
                col, index = ADDRESS.fullmatch(cell.get('r')).groups()
                yield number(col), int(index), f.text, f.get('ref')
        row.clear()
        while row.getprevious() is not None:
            del row.getparent()[0]


def solver_references(col, row, formula, shared, first_row, width):
    """Relative references in a shared formula can move at either edge of its range."""
    formulas = [formula]
    if shared:
        left, top, right, bottom = range_boundaries(shared)
        require((left, top) == (col, row) and right <= width
                and first_row <= top <= bottom < first_row + ROWS,
                f'Shared solver formula crosses a year or column boundary at {shared}')
        translator = Translator('=' + formula, origin=f'{letter(col)}{row}')
        for x, y in {(left, bottom), (right, top), (right, bottom)} - {(col, row)}:
            formulas.append(translator.translate_formula(f'{letter(x)}{y}')[1:])
    for expression in formulas:
        for ref in references(expression, SOLVER):
            if ref.sheet == SOLVER and ref.left != number('BE'):
                require(ref.right <= width and first_row <= ref.top <= ref.bottom < first_row + ROWS,
                        f'Solver reads outside its year at {letter(col)}{row}')
            yield ref


def read_stages(source):
    body = source.split('Private Sub RecalculateTrialYears(', 1)[1].split('End Sub', 1)[0]
    setup, annual = body.split('If yearRow = 5 Then', 1)
    opening, later = annual.split('\n        Else\n', 1)
    later = later.split('\n        End If\n', 1)[0]
    bindings = dict(re.findall(r'Set (\w+) = book\.Worksheets\("([^"]+)"\)',setup))
    widths = {}
    for alias, sheet in SHEETS.items():
        row = re.escape('(yearRow + 1)' if sheet == MARKET else 'yearRow')
        pattern = rf'Set {alias} = (\w+)\.Range\("A" & {row} & ":([A-Z]+)" & {row}\)'
        found = re.search(pattern, setup)
        require(found and bindings.get(found[1]) == sheet, f'Missing or incorrect annual range for {sheet}')
        widths[sheet] = number(found[2])
    fixed = re.search(r'solver\.Range\("BE2:BE(\d+)"\)\.Calculate', setup)
    links = re.search(r'firstSourceRow = (\d+) \+ \(yearRow - (\d+)\) \* (\d+)', later)
    require(fixed and links, 'Funding source ranges are missing')
    prefix = int(fixed[1])
    next_source, base_row, stride = map(int, links.groups())
    require(f'firstSourceRow + {stride - 1}' in later, 'Funding source width has changed')
    require(f'firstSolverRow = 2 + (yearRow - 6) * {ROWS}' in later and
            f'firstSolverRow + {ROWS - 1}' in later, 'Funding solver row count differs from its generator')
    last_col = re.search(r'firstSolverRow & ":([A-Z]+)"', later)
    require(last_col, 'Funding solver range is missing')
    first_source = f'solver.Range("BE2:BE{prefix}").Calculate'
    later_source = f'solver.Range("BE" & firstSourceRow & ":BE" & (firstSourceRow + {stride - 1})).Calculate'
    solver = f'solver.Range("A" & firstSolverRow & ":{last_col[1]}" & (firstSolverRow + {ROWS - 1})).Calculate'

    def parse(part):
        steps = []
        for text in part.splitlines():
            line = text.strip()
            require(not line.startswith(('If ', 'Else', 'End If')), f'Unclassified annual VBA branch: {line}')
            if line.endswith('Year.Calculate'):
                key = line.removesuffix('.Calculate')
                require(key in SHEETS, f'Unknown annual VBA range {key}')
                steps.append(SHEETS[key])
            elif line in (first_source, later_source):
                steps.append('sources')
            elif line == solver:
                steps.append('solve')
            else:
                require('.Calculate' not in line, f'Unclassified annual calculation: {line}')
        return steps

    # Year 6 uses the opening source block. Each subsequent year has its own
    # source block. Ensure that each branch actually calculates its inputs.
    before, selector, remainder = later.partition('            If yearRow = 6 Then\n')
    initial, otherwise, remainder = remainder.partition('\n            Else\n')
    following, end_if, after = remainder.partition('\n            End If\n')
    require(selector and otherwise and end_if, 'The separate first-year and later funding source branches are missing')
    require(initial.strip() == first_source and following.rstrip().endswith(later_source),
            'Both funding source branches must calculate their own input ranges')
    orders = (parse(opening), parse(before) + parse(initial) + parse(after),
              parse(before) + parse(following) + parse(after))
    return orders, widths, (prefix, next_source, base_row, stride), number(last_col[1])


def verify_funding_cleanup(source, linked_horizon, solver_width):
    """Keep every cleanup range within the generated withdrawal-search table."""
    rows = re.search(r'Private Const FUNDING_ROWS_PER_YEAR As Long = (\d+)', source)
    require(rows and int(rows[1]) == ROWS, 'Inactive funding row count differs from the solver generator')
    require(linked_horizon is not None and linked_horizon.is_cell('Trial_Inputs', 2, 70),
            'The inactive funding horizon no longer points to Trial_Inputs!B70')
    body = source.split('Private Sub UpdateInactiveFunding(', 1)[-1].split('End Sub', 1)[0]
    required = (
        'If years < mLastFundingYears And mLastFundingYears > 1 Then',
        'firstRow = 2',
        'If years > 1 Then firstRow = firstRow + (years - 1) * FUNDING_ROWS_PER_YEAR',
        'lastRow = 1 + (mLastFundingYears - 1) * FUNDING_ROWS_PER_YEAR',
        '.Range("BE2").Calculate',
        f'Set inactive = .Range("A" & firstRow & ":{letter(solver_width)}" & lastRow)',
        'inactive.Calculate',
        'Application.WorksheetFunction.CountBlank(inactive) <> inactive.CountLarge',
        'mLastFundingYears = years',
    )
    for operation in required:
        require(operation in body, f'Inactive funding calculation is missing {operation}')
        body = body.split(operation, 1)[1]


def prep_regions(vba):
    """Read the actual constant-range calls that prepare controls, draws and inflation."""
    body = vba.split('Private Sub RecalculateTrialYears(',1)[1].split('    For yearRow = 5',1)[0]
    bindings = dict(re.findall(r'Set (\w+) = book\.Worksheets\("([^"]+)"\)',body))
    regions = []
    for variable,address in re.findall(r'(\w+)\.Range\("([A-Z0-9:]+)"\)\.Calculate',body):
        regions.append((bindings[variable],range_boundaries(address)))
    return regions


def verify_preparation(archive, parts, regions):
    """Check the controls and the first market observations before the opening balance."""
    formulas = {}
    for sheet in (MARKET, 'Random_Numbers', 'Trial_Inputs'):
        formulas[sheet] = {}
        for col, row, formula, shared in cells(archive, parts[sheet]):
            require(formula is not None and not shared, f'Unsupported shared control formula on {sheet}')
            formulas[sheet][col, row] = formula

    # Market correlation helpers sit outside the annual grid. Trace their formulas to static inputs.
    cached = {}
    def stable(ref, seen):
        if ref.sheet in STABLE or ref.is_cell('Risk_Trials', 2, 3):
            return True
        if ref.sheet != MARKET or ref.left <= number('AI') or ref.top < 4:
            return False
        for col in range(ref.left, ref.right + 1):
            for row in range(ref.top, ref.bottom + 1):
                cell = (col, row)
                if cell in seen:
                    return False
                if cell not in cached:
                    formula = formulas[MARKET].get(cell)
                    cached[cell] = not formula or all(stable(child, seen | {cell})
                                                       for child in references(formula, MARKET))
                if not cached[cell]:
                    return False
        return True

    ready = []
    for sheet, box in regions:
        if sheet == SOLVER:
            continue  # Rechecked after all first-year funding inputs have been calculated.
        require(sheet in formulas, f'Unexpected preparation stage {sheet}')
        ready.append((sheet, box))
        left, top, right, bottom = box
        stage = TopologicalSorter()
        for (col, row), formula in formulas[sheet].items():
            if left <= col <= right and top <= row <= bottom:
                local = set()
                for ref in references(formula, sheet):
                    require(ref.inside(ready) or stable(ref, set()),
                            f'Unprepared {ref.sheet}!{letter(ref.left)}{ref.top} at {sheet}!{letter(col)}{row}')
                    if ref.sheet == sheet:
                        for c in range(max(left, ref.left), min(right, ref.right) + 1):
                            for r in range(max(top, ref.top), min(bottom, ref.bottom) + 1):
                                if (c, r) in formulas[sheet]:
                                    local.add((c, r))
                stage.add((col, row), *local)
        try:
            stage.prepare()
        except CycleError as error:
            raise ValueError(f'Circular preparation dependency on {sheet}: {error}') from error
    require({sheet for sheet, _ in ready} == {MARKET, 'Random_Numbers', 'Trial_Inputs'},
            'Missing financial preparation stage')
    return ready, stable, formulas['Trial_Inputs']


def walk_order(graph, order):
    """Excel resolves same-range edges within each explicitly calculated range."""
    pending = dict(graph)
    for stage in order:
        while True:
            ready = [node for node, deps in pending.items() if node[0] == stage and not any(d in pending for d in deps)]
            if not ready:
                break
            for node in ready:
                del pending[node]
    require(not pending, f'Calculation sequence leaves unresolved cells: {list(pending)[:3]}')


def annual_reference(ref, fiscal_row, fields):
    row = fiscal_row + int(ref.sheet == MARKET)
    require(ref.bottom <= row, f'Forward financial dependency from row {fiscal_row} to {ref.sheet}!{letter(ref.right)}{ref.bottom}')
    if not ref.top <= row <= ref.bottom:
        return set()
    return {(ref.sheet, col) for col in range(ref.left, ref.right + 1)
            if (ref.sheet, col) in fields}


def public_input_link(col, row, input_formulas):
    address = f'{letter(col)}{row}'
    require((col, row) in input_formulas, f'Missing Trial_Inputs!{address} financial assumption')
    if address not in ('B14', 'B53', 'B70'):
        links = list(references(input_formulas[col, row], 'Trial_Inputs'))
        require(links and all(ref.is_cell('Inputs', col, row) for ref in links),
                f'Trial assumption {address} does not read its public input')


def audit(workbook, vba):
    source = vba.read_text()
    orders, widths, (prefix, next_source, base_row, stride), solver_width = read_stages(source)
    years = {y: {} for y in range(FIRST, LAST+1)}
    source_rows, solver_sources, solver_direct = {}, defaultdict(set), defaultdict(set)
    terminals = {}
    captures = ((native_capture.SHEET, native_capture.FIRST - 1, native_capture.SIGNATURE_COLUMN),
                (native_funding_paths.DATA_SHEET, native_funding_paths.CURRENT_ROW, native_funding_paths.SIGNATURE_COLUMN))
    with ZipFile(workbook) as archive:
        parts = sheet_parts(archive)
        prepared, static_market, input_formulas = verify_preparation(archive,parts,prep_regions(source))
        for sheet, trial_row, last_column in captures:
            output = {}
            for col, row, formula, _ in cells(archive, parts[sheet]):
                if row > trial_row:
                    break
                if row == trial_row and 2 <= col <= last_column:
                    require(formula, f'Missing current result formula on {sheet}!{letter(col)}{row}')
                    output[col] = tuple(references(formula, sheet))
            require(set(output) == set(range(2, last_column + 1)), f'Incomplete current capture row on {sheet}')
            terminals[sheet] = output

        for sheet, width in widths.items():
            for col, row, formula, shared in cells(archive, parts[sheet]):
                y = row - int(sheet == MARKET)
                if y not in years or (sheet == MARKET and col > number('AI')):
                    continue
                require(col <= width and formula and not shared, f'Annual VBA range or formula representation changed: {sheet}!{letter(col)}{row}')
                years[y][sheet, col] = tuple(references(formula, sheet))
        for col, row, formula, shared in cells(archive, parts[SOLVER]):
            if row < 2:
                continue
            if col == number('BE'):
                links = list(references(formula, SOLVER))
                require(len(links) == 1 and links[0].sheet != SOLVER, f'Unexpected funding source BE{row}')
                source_rows[row] = links[0]
                continue
            require(col <= solver_width, f'Funding formula outside the evaluated solver range: {letter(col)}{row}')
            if formula is None:
                continue
            y = 6 + (row - 2) // ROWS
            start = 2 + (y - 6) * ROWS
            for ref in solver_references(col, row, formula, shared, start, solver_width):
                if ref.sheet != SOLVER:
                    solver_direct[y].add(ref)
                elif ref.left == ref.right == number('BE'):
                    require(ref.absolute, f'Relative funding source in {letter(col)}{row}')
                    solver_sources[y].update(range(ref.top, ref.bottom+1))
                else:
                    require(ref.right <= solver_width, f'Unclassified solver range at {letter(col)}{row}')

    verify_funding_cleanup(source, source_rows.get(2), solver_width)
    checked_inputs = set()
    checked_financial_ranges = set()
    def check_public_input(col, row):
        if (col, row) in checked_inputs:
            return
        public_input_link(col, row, input_formulas)
        checked_inputs.add((col, row))

    def current_deps(ref, y):
        if ref.sheet == SOLVER:
            start = 2 + (y - 6) * ROWS
            require(start <= ref.top <= ref.bottom < start+ROWS and ref.right <= solver_width,
                    f'Annual ledger reads a different solver year at model row {y}')
            return {('solve', 0)}
        if ref.sheet in STABLE:
            return set()
        if ref.sheet in {'Random_Numbers', 'Trial_Inputs'}:
            if ref.sheet == 'Trial_Inputs':
                for row in range(ref.top, ref.bottom + 1):
                    for col in range(ref.left, ref.right + 1):
                        check_public_input(col, row)
            require(ref.inside(prepared), f'Financial input was not prepared: {ref.sheet}!{letter(ref.left)}{ref.top}')
            return set()
        if ref.sheet == 'Risk_Trials':
            require(ref.is_cell('Risk_Trials', 2, 3), 'Stored results cannot be an annual financial precedent')
            return set()
        require(ref.sheet in widths, f'Unclassified source worksheet {ref.sheet}')
        if ref.sheet == MARKET and (ref.top < 6 or ref.left > number('AI')):
            require(ref.inside(prepared) or static_market(ref, set()), f'Unprepared market helper {letter(ref.left)}{ref.top}')
            return set()
        if ref not in checked_financial_ranges:
            for row in range(ref.top, ref.bottom + 1):
                fiscal_row = row - int(ref.sheet == MARKET)
                for col in range(ref.left, ref.right + 1):
                    require((ref.sheet, col) in years.get(fiscal_row, {}),
                            f'Missing annual financial formula at {ref.sheet}!{letter(col)}{row}')
            checked_financial_ranges.add(ref)
        return annual_reference(ref, y, years[y])

    for y, fields in years.items():
        graph = {node: set().union(*(current_deps(ref,y) for ref in links)) for node,links in fields.items()}
        if y > FIRST:
            start = next_source + (y-base_row)*stride
            linked = solver_sources[y]
            for addr in linked:
                require(addr in source_rows, f'Absent solver source BE{addr}')
                if y > 6:
                    require(start <= addr < start+stride or
                            (addr <= prefix and source_rows[addr].sheet in STABLE | {'Trial_Inputs'}),
                            f'Solver source BE{addr} is not refreshed in model row {y}')
                else:
                    require(addr <= prefix, f'Unrefreshed first-year funding source BE{addr}')
            graph['sources',0] = set().union(*(current_deps(source_rows[i],y) for i in linked))
            graph['solve',0] = {('sources',0)} | set().union(*(current_deps(ref,y) for ref in solver_direct[y]))
        walk_order(graph, orders[min(y - FIRST, 2)])
    for sheet, trial_row, last_column in captures:
        graph = {}
        for col, links in terminals[sheet].items():
            current = set()
            for ref in links:
                if ref.is_cell(MARKET, number('BB'), 3):
                    continue  # Assumption signature, checked again before publication.
                if ref.sheet == sheet and ref.top == ref.bottom == trial_row:
                    require(2 <= ref.left <= ref.right <= last_column, f'Capture source outside the terminal row on {sheet}')
                    current.update((sheet, i) for i in range(ref.left, ref.right + 1))
                elif (sheet == native_funding_paths.DATA_SHEET and ref.sheet == native_capture.SHEET
                      and ref.top == ref.bottom == native_capture.FIRST - 1):
                    require(all(i in terminals[ref.sheet] for i in range(ref.left, ref.right + 1)), 'Missing ranked capture precedent')
                else:
                    current_deps(ref, LAST)
            graph[sheet, col] = current
        walk_order(graph, [sheet])
    print(f'Verified workbook dependencies and VBA order across {len(years)} positions, '
          f'{sum(map(len, years.values())):,} annual formulas and {LAST-FIRST} funding groups.')


if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook',type=Path)
    parser.add_argument('--vba',type=Path,default=PROJECT / 'tools/vba/RunMarketRisk.bas')
    args=parser.parse_args()
    audit(args.workbook, args.vba)
