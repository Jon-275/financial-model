"""Build the detailed measures and the data that feed the retirement plan."""

from plan_report import METRICS_REF, display_dollars
from risk_presentation import CURRENT, annual_dollars, result

SHEET = 'Report_Data'
LAST_PORTFOLIO_ROW = 127


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
        'A8': ('Spending benchmark before income (today’s dollars)', None),
        'C8': ('Living costs after the mortgage ends plus the pre-Medicare healthcare budget, divided by the withdrawal rate. The annual funding result includes income and taxes.', 'note'),
        'A12': ('First-year savings rate', None),
        'C12': ('Share of household cash income and employer contributions saved in year one.', 'note'),
        'A15': ('Portfolio at retirement (future dollars)', None),
        'A17': ('Spending benchmark at retirement (future dollars)', None),
        'A18': ('Savings compared with benchmark (future dollars)', None),
        'C18': ('Difference from the spending benchmark before income. The funding result shows whether the timed income and accessible savings pay the actual modeled costs.', 'note'),
        'A19': ('Savings available at retirement', None),
        'C19': ('Brokerage and Roth money the plan can use in the first retirement year before age 60.', 'note'),
        'A22': ('YEAR-BY-YEAR PLAN DETAILS', 'band'),
        'A23': ('First year savings reach the benchmark', None),
        'C23': ('First calendar year the portfolio reaches the spending benchmark before income.', 'note'),
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
        'E1': 'Year', 'F1': 'Income used', 'G1': 'Withdrawals used',
        'H1': 'Health savings used', 'I1': 'Unpaid', 'J1': 'Projection position',
        'K1': 'Savings used', 'L1': 'Other spending', 'M1': 'Healthcare', 'N1': 'Tax due',
        'P1': 'Report controls', 'Q1': 'Position or count',
    }.items():
        put(SHEET, cell, label, 'header')
    for row, label in enumerate(('Modeled years', 'First cash flow position', 'Cash flow rows',
                                 'Retirement position', 'First shortfall position'), 2):
        put(SHEET, f'P{row}', label, 'report_caption')
    for cell, formula in {
        'Q2': f'=IF(OR({METRICS_REF}B25="YES",{METRICS_REF}B25="NO - cash shortfall"),Inputs!$B$70,0)',
        'Q3': '=MAX(2,Q5-2)',
        'Q4': '=IF(OR(Q2<2,Q5>Q2),0,MIN(Q2,Q5+10)-Q3+1)',
        'Q5': '=IF(Q2=0,2,MAX(2,Inputs!$B$10-Inputs!$B$8+1))',
        'Q6': f'=IF({METRICS_REF}B25="NO - cash shortfall",MATCH({METRICS_REF}B26,Projection!$B$5:$B$130,0),0)',
    }.items():
        put(SHEET, cell, formula, 'whole')

    for row in range(2, LAST_PORTFOLIO_ROW + 1):
        pos, src = row - 1, row + 3
        cpi, healthcare = f'Projection!G{src}', f'Projection!BA{src}'
        cash = lambda column: f'Cash_Flow!{column}{src}'
        paid = f'({cash("N")}+{cash("P")})'
        values = {
            'A': f'Projection!C{src}', 'B': f'Projection!BN{src}',
            'C': f'IF({pos}=$Q$5,B{row},"")', 'D': f'IF({pos}=$Q$6,B{row},"")',
            'E': cash('A'), 'F': f'MIN({cash("C")},{paid})/{cpi}',
            'G': f'MAX(0,{paid}-{cash("C")})/{cpi}', 'H': f'{cash("E")}/{cpi}',
            'I': f'{cash("V")}/{cpi}', 'J': str(pos), 'K': f'G{row}+H{row}',
            'L': f'({cash("D")}-{healthcare})/{cpi}', 'M': f'{healthcare}/{cpi}',
            'N': f'({cash("F")}+{cash("M")})/{cpi}',
            'AE': display_dollars(f"$B{row}", cpi),
        }
        for column, expression in values.items():
            style = 'whole' if column in ('A', 'E', 'J') else 'money'
            put(SHEET, f'{column}{row}', f'=IF({pos}<=$Q$2,{expression},"")', style)

    risk_current = CURRENT
    for column, heading in zip(('AE', 'AF', 'AG', 'AH'),
                               ('Base projection', '10th percentile', 'Width of middle 80%', 'Simulated median')):
        put(SHEET, f'{column}1', heading, 'header')
    for row in range(2, LAST_PORTFOLIO_ROW + 1):
        source = row + 44
        current_year = f'AND({risk_current},{row - 1}<=$Q$2)'
        for column, expression in {
            'AF': annual_dollars('D','BI',source-41),
            'AG': f'({annual_dollars("F","BK",source-41)})-({annual_dollars("D","BI",source-41)})',
            'AH': annual_dollars('E','BJ',source-41),
        }.items():
            put(SHEET, f'{column}{row}', f'=IF({current_year},{expression},"")', 'money')

    # Two-point guide lines survive Excel's PDF export and also mark a
    # shortfall when the portfolio is already depleted.
    for age, balance, position, label in (
        ('AA', 'AB', '$Q$5', 'Retirement'), ('AC', 'AD', '$Q$6', 'First shortfall'),
    ):
        valid = f'AND({position}>0,{position}<=$Q$2)'
        put(SHEET, f'{age}1', f'{label} age', 'header')
        put(SHEET, f'{balance}1', f'{label} guide', 'header')
        for row, height in ((2, '0'), (3, 'MAX(1,MAX($B$2:$B$127))')):
            put(SHEET, f'{age}{row}', f'=IF({valid},INDEX($A$2:$A$127,{position}),"")', 'whole')
            put(SHEET, f'{balance}{row}', f'=IF({valid},{height},"")', 'money')
