"""Lay out the printable record of the assumptions used by the active scenario."""

# The report lives below the entry grid so expanding or collapsing input groups
# cannot change what prints.
FIRST_ROW = 136
PAGE_BREAK = 162
LAST_ROW = 194
PRINT_AREA = f'Inputs!$L${FIRST_ROW}:$S${LAST_ROW}'
FIRST_HEADERS = 143
SECOND_HEADERS = 170
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
        ('Annual spending now', 36, 'money', 'Opening Roth type', 'IF(B28=1,"IRA","Workplace")', 'body'),
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


def configure(put):
    def display(address, value, style='report_body'):
        put('Inputs', address, value, style)

    valid = 'IF(ISNUMBER(B4),AND(B4>=1,B4<=3,B4=INT(B4)),FALSE())'
    context = f'=IF({valid},"Scenario: "&B7&"   •   Opening year: "&B130&"   •   Prepared "&TEXT(TODAY(),"mmm d, yyyy"),"Choose a scenario before printing.")'
    for start, kicker, title in (
        (136, 'Input record / 1', 'Input decisions'),
        (163, 'Input record / 2', 'Spending and other decisions'),
    ):
        caption = kicker if start == FIRST_ROW else f'=IF({valid},"{kicker}   •   "&B7,"{kicker}")'
        display(f'L{start}', caption, 'report_kicker')
        display(f'L{start + 1}', title, 'report_title')
        if start == FIRST_ROW:
            display(f'L{start + 3}', context, 'report_caption')
    display('L141', 'Figures reflect the selected scenario. Dollar amounts use opening-year prices.', 'report_caption')
    display('L168', 'Review the full coverage and account settings on Inputs when you use those options.', 'report_caption')
    for row, left, right, entries in EACH_PAGE:
        display(f'L{row}', left, 'report_section')
        display(f'P{row}', right, 'report_section_gutter')
        for index, pair in enumerate(entries, row + 1):
            for label_col, value_col, (label, source, style) in (
                ('L', 'O', pair[:3]), ('P', 'S', pair[3:]),
            ):
                if isinstance(source, int):
                    source = (f'B{source}' if source >= 130 else
                              f'IF(ISBLANK(INDEX(D{source}:F{source},1,B4)),"",B{source})')
                display(f'{label_col}{index}', label, 'report_caption_gutter' if label_col == 'P' else 'report_caption')
                display(f'{value_col}{index}', f'=IF({valid},{source},"")', 'report_' + style)
    display('L158', '¹ Includes inflation.  ² The fee applies to all scenarios.', 'report_caption')
    display('L184', 'One-time expenses for all scenarios', 'report_section')
    for index, source in enumerate(range(75, 83)):
        row = 186 + index % 4
        col = 'L' if index < 4 else 'P'
        no_costs = 'IF(COUNT(B75:B82)=0,"None entered","")' if index == 0 else '""'
        display(f'{col}{row}', f'=IF(N(B{source})>0,"Year "&TEXT(B{source},"0")&"    "&TEXT(C{source},"$#,##0"),{no_costs})')
    display('L196', 'Return to input fields', 'report_link')


def merges():
    ranges = [f'L{row}:S{row}' for row in (136, 137, 139, 141, 158, 163, 164, 166, 168, 184, 196)]
    for row, _, _, entries in EACH_PAGE:
        ranges += [f'L{row}:O{row}', f'P{row}:S{row}']
        for i in range(row + 1, row + len(entries) + 1):
            ranges += [f'L{i}:N{i}', f'P{i}:R{i}']
    for row in range(186, 190):
        ranges += [f'L{row}:O{row}', f'P{row}:S{row}']
    return ranges
