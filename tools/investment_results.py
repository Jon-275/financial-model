"""Selected portfolio assumptions and the cells that own them in the workbook."""

import re


SHEET = 'Asset_Allocation'
TITLE = 'Investment assumptions used by the plan'
MEASURES = (
    ('Taxable brokerage return', 17, 18), ('Traditional account return', 19, 20),
    ('Roth account return', 21, 22), ('Health savings account return', 134, 135),
    ('Return using opening balances', 23, 24), ('Taxable qualified dividend yield', 63, 64),
    ('Taxable interest yield', 65, 66),
)
# Older workbooks and the Python scenario interface identify these measures by Inputs row.
CELLS = {key: f'{column}{row}' for row, (_, working, retired) in enumerate(MEASURES, 99)
         for key, column in ((working, 'B'), (retired, 'C'))}
LABELS = {key: f'{label}, {phase}' for label, working, retired in MEASURES
          for key, phase in ((working, 'working years'), (retired, 'retirement'))}
RETIRED_ROWS = frozenset(CELLS)
LEGACY_REFERENCE = re.compile(r"(?<![\w.])(?:'Inputs'|Inputs)!\$?B\$?(\d+)(?!\d)", re.IGNORECASE)


def reference(row):
    cell = CELLS[row]
    return f'{SHEET}!${cell[0]}${cell[1:]}'


def from_template(formula):
    """Translate the audited template's direct Inputs references into their current owner."""
    def replace(match):
        row = int(match[1])
        return reference(row) if row in CELLS else match[0]
    return LEGACY_REFERENCE.sub(replace, formula)


def configure(put):
    put(SHEET, 'A97', TITLE, 'detail_band')
    for column, text in zip('ABC', ('Account or measure', 'Working years', 'Retirement')):
        put(SHEET, f'{column}98', text, 'report_table_head')
    for index, (label, _, _) in enumerate(MEASURES, 99):
        put(SHEET, f'A{index}', label, 'report_caption')
    for column, selection in (('B', 15), ('C', 16)):
        profile = f'Inputs!$B${selection}'
        for offset in range(4):
            gross = f'INDEX($B$54:$B$65,({profile}-1)*4+{offset + 1})'
            put(SHEET, f'{column}{99 + offset}', f'=(1+{gross})*(1-Inputs!$B$131)-1', 'percent_active')
        weighted = '+'.join(f'Inputs!$B${balance}*{column}{99 + i}'
                            for i, balance in enumerate((25, 26, 27, 31)))
        put(SHEET, f'{column}103', f'=IFERROR(({weighted})/SUM(Inputs!$B$25:$B$27,Inputs!$B$31),{column}99)', 'percent_active')
        for row, source in ((104, 'D'), (105, 'C')):
            put(SHEET, f'{column}{row}', f'=INDEX(${source}$54:${source}$65,({profile}-1)*4+1)', 'percent_active')
    put(SHEET, 'A106', 'Account returns include the annual investment fee. The portfolio return uses opening account balances.', 'report_caption')
    put(SHEET, 'A107', 'Interest and qualified dividends are part of the taxable return. The model uses these yields to calculate tax.', 'report_caption')
