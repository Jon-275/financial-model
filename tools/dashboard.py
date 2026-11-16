"""Build the detailed measures and the data that feed the retirement plan."""

from plan_report import add_transition_values

SHEET = 'Report_Data'
LAST_PORTFOLIO_ROW = 127
LAST_CASH_ROW = 14


def configure(put):
    end_row = 'Inputs!$B$70'
    end_roth = lambda col: f'INDEX(Roth_Access!${col}$5:${col}$130,{end_row})'
    restricted = f'IF({end_roth("R")}=0,{end_roth("AK")},0)+IF({end_roth("S")}=0,{end_roth("AL")},0)'
    end_cpi = f'INDEX(Projection!$G$5:$G$130,{end_row})'
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
        'A6': ('Workplace Roth unavailable at end', 'label'),
        'B6': (f'=IF(B5="","",({restricted})/{end_cpi})', 'summary_money'),
        'C6': ('=IF(N(B6)>0,"Included above. The plan cannot draw it at the end age.","")', 'note'),
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
        'A1': 'Age', 'B1': 'Portfolio balance', 'C1': 'Retirement starts', 'D1': 'First shortfall',
        'E1': 'Age', 'F1': 'Income used', 'G1': 'Withdrawals used',
        'H1': 'Health savings used', 'I1': 'Unpaid',
        'J1': 'Projection position', 'K1': 'Report positions',
        'U1': 'Year', 'V1': 'Savings used', 'W1': 'Other spending',
        'X1': 'Healthcare', 'Y1': 'Tax due',
    }.items():
        put(SHEET, cell, label, 'header')
    for cell, formula in {
        'K2': '=IF(OR(N25="YES",N25="NO - cash shortfall"),Inputs!$B$70,0)',
        'K3': '=MAX(2,K5-2)',
        'K4': '=IF(OR(K2<2,K5>K2),0,MIN(K2,K5+10)-K3+1)',
        'K5': '=IF(K2=0,2,MAX(2,Inputs!$B$10-Inputs!$B$8+1))',
        'K6': '=IF(ISNUMBER(N26),MATCH(N26,Projection!$B$5:$B$130,0),0)',
    }.items():
        put(SHEET, cell, formula, 'whole')

    for row in range(2, LAST_PORTFOLIO_ROW + 1):
        pos, src = row - 1, row + 3
        put(SHEET, f'A{row}', f'=IF({pos}<=$K$2,Projection!C{src},"")', 'whole')
        put(SHEET, f'B{row}', f'=IF({pos}<=$K$2,Projection!BN{src},"")', 'money')
        put(SHEET, f'C{row}', f'=IF(AND({pos}<=$K$2,{pos}=$K$5),B{row},"")', 'money')

        put(SHEET, f'D{row}', f'=IF(AND({pos}<=$K$2,{pos}=$K$6),B{row},"")', 'money')

    # Two-point guide lines survive Excel's PDF export and also mark a
    # shortfall when the portfolio is already depleted.
    for age, balance, position, label in (
        ('AA', 'AB', '$K$5', 'Retirement'), ('AC', 'AD', '$K$6', 'First shortfall'),
    ):
        valid = f'AND({position}>0,{position}<=$K$2)'
        put(SHEET, f'{age}1', f'{label} age', 'header')
        put(SHEET, f'{balance}1', f'{label} guide', 'header')
        for row, height in ((2, '0'), (3, 'MAX(1,MAX($B$2:$B$127))')):
            put(SHEET, f'{age}{row}', f'=IF({valid},INDEX($A$2:$A$127,{position}),"")', 'whole')
            put(SHEET, f'{balance}{row}', f'=IF({valid},{height},"")', 'money')

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
        add_transition_values(put, row, cash, cpi)
