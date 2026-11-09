"""Build the Summary headlines and the two chart schedules."""

SHEET = 'Report_Data'
LAST_PORTFOLIO_ROW = 127
LAST_CASH_ROW = 14


def configure(put):
    labels = {
        'A1': ('RETIREMENT PLAN SUMMARY', 'title'),
        'A2': ('=IF(Inputs!$B$70=0,"Check Inputs","Scenario "&Inputs!$B$7&"  |  retire at age "&TEXT(Inputs!$B$10,"0")&"  |  end at age "&TEXT(Inputs!$B$53,"0"))', None),
        'A3': ('Can the plan cover its costs?', 'header'),
        'B3': ('=IF(OR(B25="YES",B25="NO - cash shortfall"),IF(B25="YES","YES","NO"),"CHECK PLAN")', 'header'),
        'C3': ('=IF(B25="YES","Pays all modeled spending and tax.",IF(B25="NO - cash shortfall","See the first year below and open Cash_Flow.",IF(B25="Check funding calculation","Open Checks to review the calculation.",B25)))', 'note'),
        'A4': ('First year with unpaid costs', 'label'),
        'B4': ('=IF(B26="","",B26)', 'whole'),
        'C4': ('Look up this year on Cash_Flow.', 'note'),
        'A5': ('Ending portfolio in today’s dollars', 'label'),
        'B5': ('=IF(OR(B25="YES",B25="NO - cash shortfall"),B29,"")', 'summary_money'),
        'C5': ('Balance at the selected end age, before estimated tax on traditional savings.', 'note'),
        'A7': ('SAVINGS TARGET', 'band'),
        'A8': ('Savings target in today’s dollars', None),
        'C8': ('Based on spending after the mortgage ends, healthcare and the withdrawal rate you entered.', 'note'),
        'A12': ('First-year savings rate', None),
        'C12': ('Share of household cash income and employer contributions saved in year one.', 'note'),
        'A15': ('Portfolio at retirement (future dollars)', None),
        'A17': ('Savings target at retirement (future dollars)', None),
        'A18': ('Gap to savings target (future dollars)', None),
        'C18': ('A negative number is below the target. See the funding answer above for actual planned costs.', 'note'),
        'A19': ('Savings available at retirement', None),
        'C19': ('Brokerage and Roth money the plan can use in the first retirement year before age 60.', 'note'),
        'A22': ('YEAR-BY-YEAR PLAN DETAILS', 'band'),
        'A23': ('First year the savings target is met', None),
        'C23': ('First calendar year the portfolio reaches the savings target.', 'note'),
        'A24': ('Lowest retirement balance (future dollars)', None),
        'A29': ('Portfolio at end age (today’s dollars)', None),
        'A32': ('After-tax net worth at end (future dollars)', None),
        'A49': ('SCENARIO COMPARISON', 'band'),
        'A50': ('Inputs that differ across scenarios', 'label'),
        'B50': ('=COUNTIF(Inputs!$H$8:$H$69,"VARIES")+COUNTIF(Inputs!$H$86:$H$98,"VARIES")+COUNTIF(Inputs!$H$101:$H$124,"VARIES")+COUNTIF(Inputs!$H$126:$H$128,"VARIES")', 'whole'),
        'C50': ('Change one assumption at a time to see what drives the result.', 'note'),
    }
    for cell, (value, style) in labels.items():
        put('Summary', cell, value, style)

    put('Income_Targeting', 'A3',
        '=IF(AND($B$6>0,$B$10=0),"This year has no traditional balance to convert.","")', 'note')

    for cell, label in {
        'A1': 'Age', 'B1': 'Portfolio balance', 'C1': 'Retirement starts',
        'E1': 'Age', 'F1': 'Income used', 'G1': 'Withdrawals used',
        'H1': 'Health savings used', 'I1': 'Unpaid',
        'J1': 'Projection position', 'K1': 'Chart ranges',
    }.items():
        put(SHEET, cell, label, 'header')
    for cell, formula in {
        'K2': '=IF(OR(Summary!$B$25="YES",Summary!$B$25="NO - cash shortfall"),Inputs!$B$70,0)',
        'K3': '=MAX(2,K5-2)',
        'K4': '=IF(OR(K2<2,K5>K2),0,MIN(K2,K5+10)-K3+1)',
        'K5': '=IF(K2=0,2,MAX(2,Inputs!$B$10-Inputs!$B$8+1))',
    }.items():
        put(SHEET, cell, formula, 'whole')

    for row in range(2, LAST_PORTFOLIO_ROW + 1):
        pos, src = row - 1, row + 3
        put(SHEET, f'A{row}', f'=IF({pos}<=$K$2,Projection!C{src},"")', 'whole')
        put(SHEET, f'B{row}', f'=IF({pos}<=$K$2,Projection!BN{src},"")', 'money')
        put(SHEET, f'C{row}', f'=IF(AND({pos}<=$K$2,{pos}=$K$5),B{row},"")', 'money')

    for row in range(2, LAST_CASH_ROW + 1):
        pos = f'$J{row}'
        put(SHEET, f'J{row}', f'=IF({row - 1}<=$K$4,$K$3+{row - 2},"")', 'whole')
        cash = lambda c: f'INDEX(Cash_Flow!${c}$5:${c}$130,{pos})'
        cpi = f'INDEX(Projection!$G$5:$G$130,{pos})'
        paid = f'({cash("N")}+{cash("P")})'
        values = {
            'E': f'INDEX(Projection!$C$5:$C$130,{pos})',
            'F': f'MIN({cash("C")},{paid})/{cpi}',
            'G': f'MAX(0,{paid}-{cash("C")})/{cpi}',
            'H': f'{cash("E")}/{cpi}',
            'I': f'{cash("V")}/{cpi}',
        }
        for col, formula in values.items():
            style = 'whole' if col == 'E' else 'money'
            put(SHEET, f'{col}{row}', f'=IF({pos}="","",{formula})', style)
