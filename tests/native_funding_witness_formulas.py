"""Generate test-only worksheet formulas that can refute a funded selection."""

from argparse import ArgumentParser
from collections import defaultdict
from collections.abc import Iterable
from functools import cache
import json
from types import SimpleNamespace

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import get_column_letter

import funding_search
import healthcare
from native_capture import ANNUAL_CAPACITY
import tax_policy
from native_risk import _rename


FIRST_YEAR = 6
LAST_YEAR = 4 + ANNUAL_CAPACITY
DRAW_STEP = .01
CONVERGENCE = .005


def _unguarded(formula: str, condition: str) -> str:
    """Extract an emitted expression after verifying its expected guard."""
    prefix, suffix = f'=IF({condition},', ',"")'
    if not formula.startswith(prefix) or not formula.endswith(suffix):
        raise ValueError('The funding emitter changed its candidate quote guard.')
    return formula[len(prefix):-len(suffix)]


@cache
def _quote_template() -> dict[str, str]:
    """Record one emitted candidate and its own withdrawal breakpoints."""
    row = FIRST_YEAR
    recorded = {}

    def record(_sheet: str, address: str, formula: str, _style: str | None = None) -> None:
        """Collect the solver expressions without building a workbook."""
        recorded[address] = formula

    def direct_reference(_sheet: str, formula: str) -> str:
        """Keep direct annual references in place of the solver's cached aliases."""
        return formula

    model = SimpleNamespace(put=record, shared=defaultdict(dict),
                            source_cache=SimpleNamespace(localize=direct_reference))
    context = {
        'p': lambda col: f'Projection!${col}{row}',
        'prev': lambda col: f'Projection!${col}{row - 1}',
        'cash': lambda col: f'Cash_Flow!${"AB" if col == "C" else col}{row}',
        'health': lambda col: f'Healthcare!${col}{row}',
        'active': 'TRUE()',
        'fed_loss': f'Projection!$BW{row - 1}',
        'ca_loss': f'Projection!$BY{row - 1}',
    }
    _, columns = funding_search.emit(model, 'Disposable_Quote', row, context, None, tax_policy, healthcare)
    control = row + 2 * funding_search.GRID
    candidate_guard = f'AND(ISNUMBER(D{row}),ROW()<>{control})'
    candidates = {funding_search.detail_column(column): _unguarded(
        recorded[funding_search.detail_address(f'{column}{row}')], candidate_guard
    ) for name, column in columns.items() if name != 'FundingStatus'}

    scratch = funding_search.detail_column(get_column_letter(27 + len(tax_policy.EXTRA_HEADERS)))
    aliases = {}
    for offset, destination in enumerate(('AM', 'AN', 'AO'), 1):
        origin = f'{scratch}{row + offset}'
        candidates[destination] = _unguarded(recorded[origin], 'TRUE()')
        aliases[f'${scratch}${row + offset}'] = f'${destination}{row}'

    template = {}
    for column, expression in candidates.items():
        tokens = Tokenizer(f'=IF(ISNUMBER($D{row}),{expression},"")')
        for token in tokens.items:
            if token.type == 'OPERAND' and token.subtype == 'RANGE' and token.value in aliases:
                token.value = aliases[token.value]
        template[column] = _rename(tokens.render())
    return template


def witness_formulas(rows: Iterable[int] = range(FIRST_YEAR, LAST_YEAR + 1)) -> dict[str, str]:
    """Assign one lower quote per annual row on a separate disposable worksheet.

    Supply A with the claimed funding status, B with the gross requested draw,
    and C with the independently established allowed conversion branch (0/1).
    The caller refreshes direct annual inputs before computing the quote. AP
    reports a counterexample or leaves the claim open for independent checks.
    """
    output = {}
    for row in rows:
        if type(row) is not int or not FIRST_YEAR <= row <= LAST_YEAR:
            raise ValueError(f'Funding quote rows must be {FIRST_YEAR} through {LAST_YEAR}.')
        for column, formula in _quote_template().items():
            output[f'{column}{row}'] = Translator(formula, origin=f'{column}{FIRST_YEAR}').translate_formula(
                f'{column}{row}')
        active = f'AND(ISNUMBER(Trial_Inputs!$B$70),{row}-4<=Trial_Inputs!$B$70)'
        amount, mode, status = f'$B{row}', f'$C{row}', f'$A{row}'
        cap = f'Trial_Cash_Flow!$K{row}'
        lower = f'{amount}-{DRAW_STEP}'
        guard = (f'AND({active},EXACT({status},"Funded"),ISNUMBER({mode}),OR({mode}=0,{mode}=1),'
                 f'ISNUMBER({amount}),ISNUMBER({cap}),{amount}<={cap},{amount}>={DRAW_STEP},'
                 f'{amount}-({lower})>{CONVERGENCE})')
        output[f'D{row}'] = f'=IFERROR(IF({guard},{lower},""),"")'
        output[f'G{row}'] = f'=IF(ISNUMBER($D{row}),{mode},"")'
        feasible = f'AND(ISNUMBER($D{row}),ISNUMBER($Z{row}),$Z{row}>{CONVERGENCE})'
        output[f'AP{row}'] = (f'=IFERROR(IF({active},IF(EXACT({status},"Unfunded"),"grid-required",'
                             f'IF({feasible},"earlier-feasible","inconclusive")),""),"inconclusive")')
    return output


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('rows', type=int, nargs='*', help='Projection row numbers. Defaults to all annual rows.')
    options = parser.parse_args()
    print(json.dumps(witness_formulas(options.rows) if options.rows else witness_formulas(), indent=2))
