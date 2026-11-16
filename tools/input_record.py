"""Lay out the printable record of the assumptions used by the active scenario."""

from openpyxl.formula.tokenizer import Tokenizer

FIRST_ROW = 97
SECOND_ROW = 129
LAST_ROW = 157
PAGE_BREAKS = (FIRST_ROW - 1, SECOND_ROW - 1)
FIRST_HEADERS = FIRST_ROW + 7
SECOND_HEADERS = SECOND_ROW + 7
EACH_PAGE = (
    (FIRST_HEADERS, 'Timeline and income', 'Savings and investments', (
        ('Current age', 8, 'row', 'Taxable account balance', 25, 'money'),
        ('Spouse’s current age', 9, 'row', 'Traditional account balance', 26, 'money'),
        ('Retirement age', 10, 'row', 'Roth balance', 27, 'money'),
        ('Last age shown', 53, 'row', 'Health savings balance', 31, 'money'),
        ('Number of wage earners', 11, 'row', 'Annual 401(k) contributions', 32, 'money'),
        ('Annual household pay', 12, 'money', 'Annual employer contributions', 33, 'money'),
        ('Annual pay increase¹', 13, 'percent', 'Annual health savings contribution', 34, 'money'),
        ('Annual inflation', 14, 'percent', 'Annual Roth IRA savings goal', 35, 'money'),
        ('Annual Social Security', 44, 'money', 'Working-years investment profile', 15, 'row'),
        ('Social Security claim age', 45, 'row', 'Retirement investment profile', 16, 'row'),
        ('Other taxable income', 52, 'money', 'Annual investment fee²', 131, 'percent'),
        ('Tax-exempt interest', 46, 'money', 'Savings benchmark withdrawal rate', 41, 'percent'),
    )),
    (SECOND_HEADERS, 'Spending and healthcare', 'Account and tax choices', (
        ('Annual spending now', 36, 'money', 'Opening Roth type', 'IF(B28=1,"Roth IRA","Workplace Roth")', 'body'),
        ('Annual retirement spending', 37, 'money', 'Roth IRA funds available today', 'IF(B28=1,B30,"")', 'money'),
        ('Annual fixed mortgage payment', 38, 'money', 'Withdrawal order', 'IF(B48=1,"Taxable first","Tax target")', 'body'),
        ('Final mortgage payment year', 'IF(B39>0,B39,"")', 'row', 'Annual Roth conversion', 47, 'money'),
        ('Mortgage months in final year', 'IF(B39>0,B40,"")', 'row', 'Last age for conversions', 'IF(B47>0,B50,"")', 'row'),
        ('Healthcare approach', 'IF(B101=1,"By person","Budget")', 'body', 'Traditional account owner', 'IF(B97=1,"You","Spouse")', 'body'),
        ('Annual healthcare before Medicare', 'IF(B101=0,B42,"")', 'money', 'Taxable account cost basis', 51, 'money'),
        ('Annual healthcare with Medicare', 'IF(B101=0,B43,"")', 'money', 'Cash share of other income', 86, 'percent'),
        ('Your Medicare start age', 'IF(B101=1,IF(B104=126,"None",B104),"")', 'row', 'Cash share of tax-exempt interest', 89, 'percent'),
        ('Spouse’s Medicare start age', 'IF(B101=1,IF(B105=126,"None",B105),"")', 'row', 'Traditional balance tax estimate', 61, 'percent'),
        ('Long-term care start age', 'IF(B58=999,"",B58)', 'row', 'Survivor’s Social Security share', 'IF(B54=999,"",INDEX(D55:F55,1,B4))', 'percent'),
        ('Annual long-term care cost', 'IF(B58=999,"",B59)', 'money', 'Age when survivor assumptions start', 'IF(B54=999,"",B54)', 'row'),
    )),
)


def _on_inputs(expression):
    formula = Tokenizer('=' + expression)
    for token in formula.items:
        if token.type == 'OPERAND' and token.subtype == 'RANGE':
            token.value = 'Inputs!' + token.value
    return formula.render()


def configure(put):
    def display(address, value, style='report_body'):
        put('Summary', address, value, style)

    valid = 'IF(ISNUMBER(B4),AND(B4>=1,B4<=3,B4=INT(B4)),FALSE())'
    for start, kicker, title in (
        (FIRST_ROW, '4 / Input decisions', 'Input decisions'),
        (SECOND_ROW, '5 / Spending decisions', 'Spending and account decisions'),
    ):
        display(f'B{start}', _on_inputs(f'IF({valid},"{kicker}   •   "&B7,"{kicker}")'), 'report_kicker')
        display(f'B{start + 1}', title, 'report_title')
        display(f'B{start + 3}', _on_inputs(
            f'IF({valid},"Opening year "&B130&"   •   "&B7,"Choose a scenario on Inputs.")'
        ), 'report_caption')
    display(f'B{FIRST_ROW + 5}', 'These are your input amounts. Dollars use the opening year’s prices.', 'report_caption')
    display(f'B{SECOND_ROW + 5}', 'Input amounts use opening-year prices. See Inputs for detailed healthcare settings.', 'report_caption')
    for row, left, right, entries in EACH_PAGE:
        display(f'B{row}', left, 'report_section')
        display(f'F{row}', right, 'report_section_gutter')
        for index, pair in enumerate(entries, row + 1):
            for label_col, value_col, (label, source, style) in (
                ('B', 'E', pair[:3]), ('F', 'I', pair[3:]),
            ):
                if isinstance(source, int):
                    source = (f'B{source}' if source >= 130 else
                              f'IF(ISBLANK(INDEX(D{source}:F{source},1,B4)),"",B{source})')
                display(f'{label_col}{index}', label, 'report_caption_gutter' if label_col == 'F' else 'report_caption')
                display(f'{value_col}{index}', _on_inputs(f'IF({valid},{source},"")'), 'report_' + style)
    display('B119', '¹ Includes inflation.  ² The fee applies to all scenarios.', 'report_caption')
    display('B150', 'One-time expenses for all scenarios', 'report_section')
    for index, source in enumerate(range(75, 83)):
        row = 152 + index % 4
        col = 'B' if index < 4 else 'F'
        no_costs = 'IF(COUNT(B75:B82)=0,"None entered","")' if index == 0 else '""'
        display(f'{col}{row}', _on_inputs(
            f'IF(N(B{source})>0,"Year "&TEXT(B{source},"0")&"    "&TEXT(C{source},"$#,##0"),{no_costs})'
        ))


def merges():
    ranges = [f'B{row}:I{row}' for row in (97, 98, 100, 102, 119, 129, 130, 132, 134, 150)]
    for row, _, _, entries in EACH_PAGE:
        ranges += [f'B{row}:E{row}', f'F{row}:I{row}']
        for i in range(row + 1, row + len(entries) + 1):
            ranges += [f'B{i}:D{i}', f'F{i}:H{i}']
    for row in range(152, 156):
        ranges += [f'B{row}:E{row}', f'F{row}:I{row}']
    return ranges
