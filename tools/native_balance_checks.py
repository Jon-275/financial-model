"""Recalculate annual account closings from the financial formulas the trial uses."""

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter


SHEET = 'Trial_Checks'
FIRST_COLUMN = 14  # N, after the existing tax and funding checks.
PROJECTION = 'Trial_Projection'
ROTH = 'Trial_Roth_Access'
# Source column, closing-balance label, and the same-year opening cell to expand,
# if the financial formula reaches last year's Roth close through a local link.
CLOSINGS = (
    (PROJECTION, 'BC', 'Brokerage', None),
    (PROJECTION, 'BE', 'Traditional', None),
    (PROJECTION, 'BG', 'HSA', None),
    (ROTH, 'AI', 'Primary Roth IRA', 'D'),
    (ROTH, 'AJ', 'Spouse Roth IRA', 'E'),
    (ROTH, 'AK', 'Primary workplace Roth', 'F'),
    (ROTH, 'AL', 'Spouse workplace Roth', 'G'),
    (PROJECTION, 'BD', 'Taxable cost basis', None),
    (ROTH, 'AM', 'Combined Roth accounts', None),
    (PROJECTION, 'BF', 'Roth on projection', None),
    (PROJECTION, 'BH', 'Total portfolio', None),
    (PROJECTION, 'BN', "Portfolio in today's dollars", None),
)


def _on_another_sheet(formula, sheet, opening=None, opening_formula=None):
    """Keep the financial expression and resolve its local references on its source sheet."""
    tokens = Tokenizer(formula)
    found_opening = False
    for token in tokens.items:
        if token.type != 'OPERAND' or token.subtype != 'RANGE':
            continue
        owner, separator, address = token.value.rpartition('!')
        local = not separator or owner.strip("'") == sheet
        if opening is not None and local and address.replace('$', '') == opening:
            token.value = '(' + opening_formula + ')'
            found_opening = True
        elif not separator:
            token.value = sheet + '!' + address
    if opening is not None and not found_opening:
        raise ValueError(f'{sheet} closing formula no longer uses the expected opening cell {opening}')
    return tokens.render()[1:]


def configure(model):
    model.put(SHEET, 'A2', 'Choose a trial on Risk_Trials. Each closing difference in columns N–Y should be zero.', 'report_caption')
    for index, (owner, field, label, opening) in enumerate(CLOSINGS, FIRST_COLUMN):
        column = get_column_letter(index)
        model.put(SHEET, column + '4', label + ' difference', 'header')
        for row in range(5, 131):
            address = field + str(row)
            source = model.cells[owner]
            opening_cell = opening + str(row) if opening else None
            opening_formula = _on_another_sheet(source[opening_cell], owner) if opening else None
            expected = _on_another_sheet(source[address], owner, opening_cell, opening_formula)
            actual = f'{owner}!{address}'
            model.put(SHEET, f'{column}{row}',
                      f'=IF({row}-4>Trial_Inputs!$B$70,"",IF(ISNUMBER({actual}),{actual}-({expected}),NA()))', 'cents')
    last = get_column_letter(FIRST_COLUMN + len(CLOSINGS) - 1)
    first = get_column_letter(FIRST_COLUMN)
    for row in range(5, 131):
        area = f'{first}{row}:{last}{row}'
        model.put(SHEET, f'J{row}',
                  f'=IF({row}-4>Trial_Inputs!$B$70,"",IF(AND(Trial_Cash_Flow!AM{row}="OK",'
                  f'COUNT({area})={len(CLOSINGS)},MIN({area})=0,MAX({area})=0),"OK","Review"))', 'report_caption')
