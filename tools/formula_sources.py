"""Link solver source values once so trial rows can share local formulas."""

from collections import defaultdict

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter, range_boundaries


REFERENCE_FUNCTIONS = {
    'AREAS', 'CELL', 'COLUMN', 'COLUMNS', 'COUNTA', 'COUNTBLANK',
    'FORMULATEXT', 'INDIRECT', 'ISBLANK', 'ISFORMULA', 'ISLOGICAL',
    'ISNUMBER', 'ISREF', 'ISTEXT', 'OFFSET', 'ROW', 'ROWS',
}
VALUE_COLUMN = 57


class LocalSources:
    """Keep array shapes and absolute links stable within each solver sheet."""

    def __init__(self, put):
        self.put = put
        self.aliases = defaultdict(dict)
        self.next_row = defaultdict(lambda: 2)

    def _alias(self, sheet, reference):
        source, address = reference.rsplit('!', 1)
        source = source.strip("'").replace("''", "'")
        first_col, first_row, last_col, last_row = range_boundaries(address)
        key = source, first_col, first_row, last_col, last_row
        if key in self.aliases[sheet]:
            return self.aliases[sheet][key]
        row = self.next_row[sheet]
        height = last_row - first_row + 1
        width = last_col - first_col + 1
        if row == 2:
            self.put(sheet, 'BD1', 'Calculation source', 'header')
            self.put(sheet, 'BE1', 'Linked value', 'header')
        quoted = "'" + source.replace("'", "''") + "'"
        for offset in range(height):
            original_row = first_row + offset
            first = f'{get_column_letter(first_col)}{original_row}'
            self.put(sheet, f'BD{row + offset}', f'{source}!{first}', 'report_caption')
            for column in range(width):
                original = f'${get_column_letter(first_col + column)}${original_row}'
                local = f'{get_column_letter(VALUE_COLUMN + column)}{row + offset}'
                self.put(sheet, local, f'={quoted}!{original}')
        start = f'$BE${row}'
        end = f'${get_column_letter(VALUE_COLUMN + width - 1)}${row + height - 1}'
        alias = start if height == width == 1 else f'{start}:{end}'
        self.aliases[sheet][key] = alias
        self.next_row[sheet] += height
        return alias

    def localize(self, sheet, formula):
        tokens = Tokenizer(formula)
        functions = []
        for token in tokens.items:
            if token.type == 'FUNC':
                if token.subtype == 'OPEN':
                    functions.append(token.value[:-1].upper())
                else:
                    functions.pop()
            elif token.type == 'OPERAND' and token.subtype == 'RANGE' and '!' in token.value:
                if not any(name in REFERENCE_FUNCTIONS for name in functions):
                    token.value = self._alias(sheet, token.value)
        return tokens.render()
