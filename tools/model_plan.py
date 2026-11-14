"""Compile the annual funding and Roth worksheet formulas from the horizon baseline."""
from collections import defaultdict
from pathlib import Path
import json, sys
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter as L, coordinate_from_string, column_index_from_string
import dashboard
import funding_search
import risk_cleanup
import roth_contributions
import workbook_layout
import healthcare
import published_tax
import tax_policy
import targeting

END = 130
FUNDING_ROWS = STEPS = funding_search.ROWS
MODULES = ['Cash_Flow', 'Roth_Access', 'Funding_Calc', 'Checks','Healthcare','Targeting_Calc','Risk_Settings',dashboard.SHEET]
cells = defaultdict(dict)
styles = defaultdict(dict)
shared = defaultdict(dict)

def put(sheet, address, value, style=None):
    cells[sheet][address] = value
    if style: styles[sheet][address] = style

def row(sheet, r, values, style=None):
    for i,value in enumerate(values,1):
        put(sheet,f'{L(i)}{r}',value,style)

def outside(r): return f'{r}-4>Inputs!$B$70'
def active(r,x): return f'=IF({outside(r)},"",{x})'
def pj(col,r): return f'Projection!${col}${r}'
def cf(col,r): return f'Cash_Flow!${col}${r}'
def hc(col,r): return f'Healthcare!${col}${r}'
def rt(col,r): return f'Roth_Access!${col}${r}'
def fd(col,r): return f'Funding_Calc!${col}${r}'
def group(r): return 2+(r-6)*STEPS
def primary_share(r): return f'IF({pj("F",r)}="Single",1,Inputs!$B$91)'

def original(ws, addr):
    value=ws[addr].value
    if not isinstance(value,str) or not value.startswith('='): raise ValueError((ws.title,addr,value))
    return value[1:]

def inputs(wb):
    rows = [
        ('Cash share of other taxable income',1,'0 to 1. The full amount is taxable during its active years.'),
        ('Other taxable income starts at age',0,'Age of the primary person. 0 includes all years from today.'),
        ('Other taxable income ends at age',125,'Last year included. Use the same age for a one-year receipt.'),
        ('Cash share of tax-exempt interest',1,'External cash only. Do not also include this interest in brokerage returns.'),
        ('Primary share of opening Roth',0.5,'Allocate the opening balance and accessible IRA basis by owner. Confirm the split.'),
        ('Primary share of Roth IRA savings',0.5,'Remaining savings go to the spouse. Use 100% for the primary person only.'),
        ('Spouse first Roth IRA tax year',0,'0 means unknown. Confirm before relying on access to earnings.'),
        ('Primary workplace Roth first tax year',0,'0 excludes this balance from withdrawals until the plan is confirmed.'),
        ('Spouse workplace Roth first tax year',0,'0 excludes this balance from withdrawals until the plan is confirmed.'),
        ('Primary first workplace distribution age',126,'126 disables withdrawals. Enter the age the plan allows distributions.'),
        ('Spouse first workplace distribution age',126,'126 disables withdrawals. Enter the age the plan allows distributions.'),
        ('Traditional account owner',1,'1 = primary, 2 = spouse. One owner for withdrawals, conversions and RMDs.'),
        ('Include eligible Roth IRA savings?',1,'1 applies the direct Roth IRA limit. 0 leaves new Roth IRA savings out.'),
    ]
    put('Inputs','A85','CASH AND ACCOUNT ACCESS','band')
    for r,(label,value,note) in enumerate(rows,86):
        put('Inputs',f'A{r}',label,'label')
        put('Inputs',f'B{r}',f'=INDEX($D{r}:$F{r},1,$B$4)','percent_active' if r in (86,89,90,91) else 'input_active')
        for col in 'DEF': put('Inputs',f'{col}{r}',value,'percent_input' if r in (86,89,90,91) else 'input')
        put('Inputs',f'G{r}',note,'note')
        put('Inputs',f'H{r}',f'=IF(AND(D{r}=E{r},E{r}=F{r}),"","VARIES")','note')
    put('Inputs','A99','Cash and account input check','label')
    numeric='COUNT(INDEX(D86:F98,0,B4))=13'
    ages=','.join(f'B{r}=INT(B{r}),B{r}>=0,B{r}<={limit}' for r,limit in ((29,"B130"),(87,125),(88,125),(92,"B130"),(93,"B130"),(94,"B130"),(95,126),(96,126)))
    checks=','.join((numeric,ages,'B87<=B88','MIN(B86,B89:B91)>=0','MAX(B86,B89:B91)<=1',
                     'OR(B97=1,B97=2)','OR(B98=0,B98=1)','OR(B28=1,B28=2)',
                     'MIN(B25:B27,B30:B38,B42:B44,B46:B47,B52,B59)>=0',
                     'ISNUMBER(CHOOSE(B4,D55,E55,F55))','CHOOSE(B4,D55,E55,F55)>=0','CHOOSE(B4,D55,E55,F55)<=1',
                     'OR(B28=2,B30<=B27)','B13>=-1','B14>-1','MIN(B17:B22)>-1'))
    put('Inputs','B99',f'=IFERROR(IF(AND({checks}),"OK","Check cash and Roth inputs"),"Check cash and Roth inputs")','note')
    put('Inputs','A83',"Use today's dollars. Shared by all scenarios.", 'note')
    put('Inputs','A84','One-time expense check','label')
    put('Inputs','B84','=IF(COUNTA(B83:C83)=0,"OK","Move one-time expenses to Inputs rows 75–82")','note')
    put('Inputs','A130','Opening balance year','label')
    put('Inputs','B130',2026,'input')
    put('Inputs','G130','Use the year your opening ages and balances apply to. The first full forecast year is the following year.','note')
    put('Inputs','A131','Annual investment fee','label')
    put('Inputs','B131',0,'percent_input')
    put('Inputs','G131','Applied to all accounts. Enter 0 if your investment returns already include fees.','note')
    put('Inputs','A132','Annual healthcare inflation','label')
    put('Inputs','B132','','percent_input')
    put('Inputs','G132','Leave blank to use general inflation. Applies to medical costs you enter.','note')
    put('Inputs','A133','Year, fee and healthcare check','label')
    put('Inputs','B133','=IFERROR(IF(AND(ISNUMBER(B130),B130=INT(B130),B130>=2026,B130<=9999-125,ISNUMBER(B131),B131>=0,B131<1,OR(B132="",AND(ISNUMBER(B132),B132>-1))),"OK","Check opening year, fees and healthcare inflation"),"Check opening year, fees and healthcare inflation")','note')
    put('Inputs','B70','=IF(AND(B71="OK",B84="OK",B99="OK",B125="OK",B129="OK",B133="OK"),B53-B8+1,0)')
    for r in range(17,23):
        rate = original(wb['Inputs'], f'B{r}')
        put('Inputs',f'B{r}',f'=(1+({rate}))*(1-$B$131)-1')
    put('Inputs','G14','Returns and pay growth are nominal. Raising inflation alone reduces purchasing power.','note')
    for col in 'DEF':
        styles['Inputs'][f'{col}55'] = 'percent_input'
    put('Inputs','A55','Survivor’s share of combined Social Security','label')
    put('Inputs','B55','=B44*CHOOSE(B4,D55,E55,F55)')
    put('Inputs','C55','today’s $','label')
    put('Inputs','G55','Enter the survivor’s share in the scenario column. Column B shows the benefit in today’s dollars.','note')
    for addr, val in {'A13':'Annual wage growth','G13':'Enter the full annual increase, including inflation.',
                      'A27':'Opening Roth balance','A28':'Opening Roth type: 1 = IRA, 2 = workplace',
                      'A29':'Primary first Roth IRA tax year','A30':'Opening Roth IRA basis accessible today',
                      'A35':'Annual Roth IRA savings target (combined)','A50':'Last conversion age (traditional owner)',
                      'G27':'The opening balance is assigned by type and owner share. New IRA savings and conversions are separate.',
                      'G28':'Workplace Roth access requires verified plan access and a qualified distribution. No early workplace draw is assumed.',
                      'G29':'Enter the primary person’s first Roth IRA tax year, or 0 if unknown. See row 92 for the spouse.',
                      'G30':'Enter only IRA principal available today. Workplace Roth balances do not use this field.',
                      'G35':'After-tax savings target in today’s dollars. Funded from remaining cash after taxes and spending.',
                      'G52':'Enter taxable income separately from salary. Rows 86–88 control the cash share and active ages.'}.items():
        put('Inputs',addr,val)

CASH_HEADERS = ['Year','Age','Cash income','Required spending','HSA paid','Opening tax payable',
                'Roth target','Taxable available','Traditional available','Roth available','Gross draw capacity',
                'Funding status','Current tax','Tax paid','Ending tax payable','Spending paid','Unpaid spending',
                'Roth contribution','Unfunded Roth target','Taxable deposit','Cash residual','Essential shortfall',
                'Net cash need before tax','Eligible Roth savings cap','Cash before payments','Draw search width',
                'Eligible assets before draws','Cash income before premium credit','ACA premium credit']
ROTH_HEADERS = ['Year','Primary age','Spouse age','Primary IRA opening','Spouse IRA opening',
                'Primary workplace opening','Spouse workplace opening','Primary eligible opening basis','Spouse eligible opening basis',
                'Primary conversion basis','Spouse conversion basis','Primary seasoned basis','Spouse seasoned basis',
                'Primary IRA first year','Spouse IRA first year','Primary IRA accessible','Spouse IRA accessible',
                'Primary workplace accessible','Spouse workplace accessible','Roth accessible','Primary IRA draw',
                'Spouse IRA draw','Primary workplace draw','Spouse workplace draw','Primary savings','Spouse savings',
                'Primary conversion','Spouse conversion','Primary basis remaining','Spouse basis remaining',
                'Primary conversion basis remaining','Spouse conversion basis remaining','Primary seasoned basis remaining',
                'Spouse seasoned basis remaining','Primary IRA closing','Spouse IRA closing','Primary workplace closing',
                'Spouse workplace closing','Roth closing','Primary savings capacity','Spouse savings capacity','Roth net transfer']
FUND_HEADERS = ['Requested draw','Lower draw','Upper draw','Allow conversion','Taxable draw','Traditional draw',
                'Roth draw','Conversion','Realized gain','Ordinary before Social Security','Provisional income',
                'Taxable Social Security','Ordinary income','Adjusted gross income','Ordinary taxable','Federal ordinary',
                'Federal gains','Investment income tax','AMT','California taxable','California tax','Current tax','Funding margin']

def cash_ledger():
    s='Cash_Flow'
    put(s,'A1','Annual cash flow','title'); put(s,'A2','Tax is funded in the year it is calculated. Unpaid tax and unmet spending are shown separately.','note')
    row(s,4,CASH_HEADERS,'header')
    for r in range(5,END+1):
        prev=r-1; f=group(r)+STEPS-1; p=lambda c:pj(c,r); b=lambda c:pj(c,prev)
        if r==5:
            values=[p('B'),p('C'),*['0']*9,'"Opening balances"',*['0']*17]
        else:
            cap=f'MAX(0,{b("BE")})'
            age=f'IF(Inputs!$B$97=1,{p("C")},{p("D")})'
            capacity=f'IF(OR({p("E")}="Working",{age}<60,{p("R")}>0),{p("T")},{cap})'
            share=primary_share(r)
            savings_cap=f'MIN(MAX(0,{p("I")}-{p("J")}-{p("K")}),IF({share}=0,G{r},{rt("AN",r)}/{share}),IF({share}=1,G{r},{rt("AO",r)}/(1-{share})))'
            base_cash=f'{p("I")}-{p("J")}-{p("K")}+{p("N")}+{p("O")}*Inputs!$B$86+{p("AG")}*Inputs!$B$89'
            values=[p('B'),p('C'),f'AB{r}+AC{r}',
                    f'SUM({p("AW")},{p("AX")},{p("BA")})',p('P'),f'O{prev}',
                    f'IF({p("E")}="Working",Inputs!$B$35*{p("G")},0)',b('BC'),capacity,
                    f'IF({p("E")}="Working",0,{rt("T",r)})',f'SUM(H{r}:J{r})',
                    fd('Z',f),
                    p('AV'),f'MIN(F{r}+M{r},Y{r})',f'MAX(0,F{r}+M{r}-N{r})',
                    f'MIN(D{r}-E{r},MAX(0,Y{r}-N{r}))',f'MAX(0,D{r}-E{r}-P{r})',
                    f'MIN(G{r},X{r},MAX(0,Y{r}-N{r}-P{r}))',f'G{r}-R{r}',
                    f'MAX(0,Y{r}-N{r}-P{r}-R{r})',f'Y{r}-N{r}-P{r}-R{r}-T{r}',
                    f'O{r}+Q{r}',f'D{r}-E{r}+F{r}-C{r}',
                    f'IF(AND({p("E")}="Working",Inputs!$B$98=1),{savings_cap},0)',
                    f'C{r}+{p("V")}+{p("W")}+{p("Y")}',f'{fd("C",f)}-{fd("B",f)}',
                    f'H{r}+I{r}+J{r}',base_cash,fd('Y',f)]
        assert len(values)==len(CASH_HEADERS),(r,len(values))
        row(s,r,[active(r,x) for x in values],'money')

def roth():
    s='Roth_Access'; put(s,'A1','Roth balances and access','title')
    put(s,'A2','IRA basis remains separate from market value. Workplace Roth withdrawals require verified plan access and full qualification.','note')
    row(s,4,ROTH_HEADERS,'header')
    for r in range(5,END+1):
        prev=r-1; p=lambda c:pj(c,r)
        rv='Inputs!$B$27'; split='Inputs!$B$90'; typ='Inputs!$B$28'; pri='Inputs!$B$29'; spouse='Inputs!$B$92'
        initial=lambda kind,share:f'IF({typ}={kind},{rv}*{share},0)'
        if r==5:
            opens=[initial(1,split),initial(1,f'(1-{split})'),initial(2,split),initial(2,f'(1-{split})'),
                   f'IF({typ}=1,Inputs!$B$30*{split},0)',f'IF({typ}=1,Inputs!$B$30*(1-{split}),0)',
                   '0','0','0','0',pri,spouse]
        else:
            seasoned1=f'AA{r-5}' if r>=11 else '0'; seasoned2=f'AB{r-5}' if r>=11 else '0'
            opens=[f'AI{prev}',f'AJ{prev}',f'AK{prev}',f'AL{prev}',f'AC{prev}',f'AD{prev}',f'AE{prev}',f'AF{prev}',
                   f'IF(B{r}>=60,J{r},MIN(J{r},AG{prev}+{seasoned1}))',
                   f'IF(C{r}>=60,K{r},MIN(K{r},AH{prev}+{seasoned2}))',
                   f'IF(N{prev}>0,N{prev},IF(Y{prev}+AA{prev}>0,A{prev},0))',
                   f'IF(O{prev}>0,O{prev},IF(Z{prev}+AB{prev}>0,A{prev},0))']
        acc=[f'IF(AND(B{r}>=60,N{r}>0,A{r}>=N{r}+5),D{r},MIN(D{r},H{r}+L{r}))',
             f'IF(AND(C{r}>=60,O{r}>0,A{r}>=O{r}+5),E{r},MIN(E{r},I{r}+M{r}))',
             f'IF(AND(B{r}>=MAX(60,Inputs!$B$95),Inputs!$B$95<126,Inputs!$B$93>0,A{r}>=Inputs!$B$93+5),F{r},0)',
             f'IF(AND(C{r}>=MAX(60,Inputs!$B$96),Inputs!$B$96<126,Inputs!$B$94>0,A{r}>=Inputs!$B$94+5),G{r},0)',f'SUM(P{r}:S{r})']
        y=p('Y'); m=p('M'); cv=p('AA'); rate=f'IF({p("E")}="Working",Inputs!$B$21,Inputs!$B$22)'
        alloc=primary_share(r)
        moves=['0']*8 if r==5 else [f'MIN({y},P{r})',f'MIN(MAX(0,{y}-U{r}),Q{r})',
            f'MIN(MAX(0,{y}-SUM(U{r}:V{r})),R{r})',f'MIN(MAX(0,{y}-SUM(U{r}:W{r})),S{r})',
            f'{m}*{alloc}',
            f'{m}-Y{r}',f'IF(Inputs!$B$97=1,{cv},0)',f'IF(Inputs!$B$97=2,{cv},0)']
        basis=[f'MAX(0,H{r}-U{r})+Y{r}',f'MAX(0,I{r}-V{r})+Z{r}',
               f'MAX(0,J{r}-MAX(0,U{r}-H{r}))+AA{r}',f'MAX(0,K{r}-MAX(0,V{r}-I{r}))+AB{r}',
               f'MAX(0,L{r}-MAX(0,U{r}-H{r}))',f'MAX(0,M{r}-MAX(0,V{r}-I{r}))']
        close=['D5','E5','F5','G5'] if r==5 else [f'MAX(0,D{r}+Y{r}+AA{r}-U{r})*(1+{rate})',
            f'MAX(0,E{r}+Z{r}+AB{r}-V{r})*(1+{rate})',f'MAX(0,F{r}-W{r})*(1+{rate})',f'MAX(0,G{r}-X{r})*(1+{rate})']
        single = f'{p("F")}="Single"'
        magi = f'({p("AK")}-{p("AA")})'
        capacity = ['0', '0'] if r == 5 else [
            roth_contributions.formula_capacity(p('B'), p('H'), f'B{r}', single, magi),
            f'IF({single},0,{roth_contributions.formula_capacity(p("B"), p("H"), f"C{r}", single, magi)})',
        ]
        vals=[p('B'),p('C'),p('D'),*opens,*acc,*moves,*basis,*close,f'SUM(AI{r}:AL{r})',*capacity,
              f'SUM(Y{r}:AB{r})-SUM(U{r}:X{r})']
        assert len(vals)==len(ROTH_HEADERS),(r,len(vals))
        row(s,r,[active(r,x) for x in vals],'money')

TAX_MAP={'AE':'I','AD':'J','AH':'K','AI':'L','AJ':'M','AK':'N','AM':'O','AN':'P','AO':'Q','AP':'R','AQ':'S','AS':'T','AT':'U','AV':'V'}
CANDIDATE={**TAX_MAP,'V':'E','W':'F','Y':'G','AA':'H'}

def funding_headers(sheet):
    funding_search.headers(sys.modules[__name__],sheet,tax_policy)


def emit_funding_group(sheet, first, context, requested_conversion):
    return funding_search.emit(sys.modules[__name__],sheet,first,context,requested_conversion,tax_policy,healthcare)


def funding(wb):
    del wb
    s="Funding_Calc"; funding_headers(s)
    for r in range(6,END+1):
        p=lambda c:pj(c,r); prior=lambda c:pj(c,r-1)
        context={"p":p,"prev":prior,"cash":lambda c:cf("AB" if c=="C" else c,r),
                 "health":lambda c:hc(c,r),"active":f"NOT({outside(r)})",
                 "fed_loss":prior("BW"),"ca_loss":prior("BY")}
        end,_=emit_funding_group(s,group(r),context,None)
        assert end==group(r)+STEPS-1


def projection(wb):
    s='Projection'
    for a,label in {'M4':'Roth IRA savings funded','Q4':'Cash need before current tax','X4':'Accessible Roth at year start',
                    'Z4':'Essential funding shortfall','AV4':'Current-year tax','BB4':'Taxable savings deposited'}.items(): put(s,a,label)
    for r in range(5,END+1):
        p=r-1; f=group(r)+STEPS-1
        put(s,f'H{r}',active(r,f'(1+Inputs!$B$14)^(B{r}-2026)'))
        if r==5:
            put(s,'B5',active(5,'Inputs!$B$130'))
            put(s,'X5',active(r,rt('T',r))); put(s,'BF5',active(r,rt('AM',r)))
            continue
        values={**{c:fd(fc,f) for c,fc in CANDIDATE.items()},
            'I':f'IF(E{r}="Working",Inputs!$B$12*(1+Inputs!$B$13)^A{r},0)',
            'J':f'IF(E{r}="Working",MIN(MAX(0,I{r}),Inputs!$B$32*G{r}),0)',
            'K':f'IF(E{r}="Working",MIN(MAX(0,I{r}-J{r}),Inputs!$B$34*G{r}),0)',
            'M':cf('R',r),'O':f'IF(AND(C{r}>=Inputs!$B$87,C{r}<=Inputs!$B$88),Inputs!$B$52*G{r},0)',
            'Q':f'MAX(0,{cf("W",r)})','X':rt('T',r),'Z':cf('V',r),'BB':cf('T',r),'BF':rt('AM',r),
            'U':f'MAX(0,Inputs!$B$49*G{r}+AL{r}-(T{r}+O{r}+AC{r}+0.85*N{r}))',
            'BJ':f'BH{r}-BI{r}-{cf("O",r)}',
            'BO':f'SUM(J{r}:M{r})+BB{r}-SUM(V{r}:W{r})-Y{r}-P{r}',
            'BR':f'Z{r}',
        }
        age=f'IF(Inputs!$B$97=1,C{r},D{r})'
        start_age='MAX(Inputs!$B$10,Inputs!$B$8+1)-Inputs!$B$8+IF(Inputs!$B$97=1,Inputs!$B$8,Inputs!$B$9)'
        start_pos='MAX(1,Inputs!$B$10-Inputs!$B$8)'
        values['R']=f'IF(AND(Inputs!$B$68=1,E{r}<>"Working",{start_age}<60,{age}<MAX(60,{start_age}+5)),MIN(BE{p},INDEX($BE$5:BE{p},{start_pos})*Inputs!$B$69),0)'
        values['S']=f'IF({age}>=MAX(72,Inputs!$B$67),BE{p}/IF({age}=72,27.4,IF({age}<=105,INDEX(Tax_Tables!$B$99:$B$131,{age}-72),INDEX(Tax_Tables!$H$99:$H$113,MIN({age},120)-105))),0)'
        for c,x in values.items(): put(s,f'{c}{r}',active(r,x))

def reporting():
    s='Summary'; h='Inputs!$B$70'; first='MAX(2,Inputs!$B$10-Inputs!$B$8+1)'
    eligible=f'AND({h}>={first},MAX(Inputs!$B$8+1,Inputs!$B$10)<60)'
    target=f'INDEX(Cash_Flow!$H$5:$H$130,{first})+INDEX(Roth_Access!$T$5:$T$130,{first})'
    for a,x in {'B5':'=COUNTIF(Inputs!$H$8:$H$69,"VARIES")+COUNTIF(Inputs!$H$86:$H$98,"VARIES")',
        'A12':'First-year savings rate','C12':'Actual savings, including employer contributions, divided by gross household cash receipts plus employer contributions.',
        'B12':'=IF(AND(Inputs!$B$70>=2,Projection!E6="Working",SUM(Cash_Flow!C6,Projection!J6:L6)>0),(SUM(Projection!J6:M6)+MAX(0,Projection!BB6-Projection!V6))/SUM(Cash_Flow!C6,Projection!J6:L6),"")',
        'A19':'Available bridge assets at retirement','B19':f'=IF({eligible},{target},"")',
        'C19':'Opening brokerage balance plus Roth funds available in the first bridge year. See Roth_Access.',
        'B25':'=IF(Inputs!$B$70=0,IF(Inputs!$B$71<>"OK",Inputs!$B$71,IF(Inputs!$B$84<>"OK",Inputs!$B$84,IF(Inputs!$B$99<>"OK",Inputs!$B$99,IF(Inputs!$B$125<>"OK","Check healthcare inputs",IF(Inputs!$B$129<>"OK",Inputs!$B$129,Inputs!$B$133))))),IF(COUNTIF(Healthcare!$T$6:$T$130,"Missing*")+COUNTIF(Healthcare!$T$6:$T$130,"Confirm*")>0,"Check healthcare inputs",IF(COUNTIF(Cash_Flow!$AM$5:$AM$130,"OK")<>Inputs!$B$70,"Check funding calculation",IF(MAX(Projection!$BR$5:$BR$130)<0.01,"YES","NO - cash shortfall"))))',
        'B26':'=IF(B25="YES","None",IF(B25="NO - cash shortfall",INDEX(Projection!$B$5:$B$130,MATCH(TRUE(),INDEX(IF(ISNUMBER(Projection!$BR$5:$BR$130),Projection!$BR$5:$BR$130>=0.01,FALSE()),0),0)),""))',
        'C25':'Checks required spending and tax in each working and retirement year.',
        'A33':'Opening Roth account type','B33':'=IF(Inputs!$B$70=0,"",IF(Inputs!$B$28=2,"Workplace Roth","Roth IRA"))',
        'C33':'New Roth IRA savings and conversions are tracked separately by owner.',
        'A34':'First-year Roth savings below target','B34':'=IF(Inputs!$B$70>=2,Cash_Flow!S6,"")',
        'C34':'An unfunded savings target is separate from a spending or tax shortfall.'}.items(): put(s,a,x)
    for a,x in {'A26':'Cash income pays spending and current-year tax first. Optional Roth IRA savings use remaining cash.',
                'A28':'Before age 60, the model uses brokerage and eligible Roth IRA principal. Workplace Roth requires verified plan access and full qualification.',
                'A29':'The model uses age 60 for age-59½ access because it projects whole years. Traditional access follows the selected owner.',
                'A40':'Withdrawal order sets which eligible accounts fund annual cash needs. See Cash_Flow for taxes paid and shortfalls.',
                'A79':'The opening Roth IRA or workplace balance is split by owner. New Roth IRA savings and conversions go into separate IRA balances.',
                'A84':'Roth_Access tracks each owner’s available principal, conversion aging and workplace qualification. Earnings may remain restricted.',
                'A118':'The funding calculation evaluates this year’s tax on actual draws. Unpaid tax carries forward and reduces after-tax net worth.',
                'A120':'Risk summaries require a funded base plan. They apply market paths to actual net portfolio cash flows and keep the base tax and access schedule.'}.items(): put('README',a,x)

def reconciliation():
    """Calculate annual reconciliation in the ledger and show its details on Checks."""
    sheet = 'Checks'
    put(sheet, 'A1', 'Model checks', 'title')
    put(sheet, 'A2', 'Cash and account balances reconcile to one cent. An unfunded year is a planning result.', 'note')
    diagnostics = {
        'C': 'Cash balance error', 'D': 'Tax balance error', 'E': 'Roth balance error',
        'F': 'Portfolio flow error', 'G': 'Roth draw over access', 'H': 'Negative account balance',
        'K': 'ACA credit reconciliation', 'L': 'HSA over eligible costs', 'M': 'Taxable basis reconciliation',
    }
    row(sheet, 4, ['Year', 'Age', *(diagnostics[col] for col in 'CDEFGH'),
                   'Funding calculation', 'Reconciliation', *(diagnostics[col] for col in 'KLM')], 'header')
    ledger_columns = {col: L(index) for index, col in enumerate(diagnostics, len(CASH_HEADERS) + 1)}
    for col, label in diagnostics.items():
        put('Cash_Flow', f'{ledger_columns[col]}4', label, 'header')
    put('Cash_Flow', 'AM4', 'Reconciliation', 'header')
    for r in range(5, END + 1):
        p = lambda col: pj(col, r)
        prior = lambda col: pj(col, r - 1)
        cash = lambda col: cf(col, r)
        roth = lambda col: rt(col, r)
        rate = f'IF({p("E")}="Working",Inputs!$B$21,Inputs!$B$22)'
        errors = {
            'C': f'{cash("C")}+SUM({p("V")},{p("W")},{p("Y")})-SUM({cash("N")},{cash("P")},{cash("R")},{cash("T")})',
            'D': f'SUM({cash("F")},{cash("M")})-SUM({cash("N")},{cash("O")})',
            'E': '0' if r == 5 else f'{p("BF")}-({prior("BF")}+{roth("AP")})*(1+{rate})',
            'F': '0' if r == 5 else f'{p("BO")}-SUM({p("J")},{p("K")},{p("L")},{p("M")},{p("BB")})+SUM({p("V")},{p("W")},{p("Y")},{p("P")})',
            'G': f'MAX(0,{p("Y")}-{roth("T")})',
            'H': f'MIN(0,{p("BC")},{p("BE")},{p("BF")},{p("BG")})',
            'K': f'{cash("AC")}-{hc("R", r)}',
            'L': f'MAX(0,{p("P")}-{hc("Q", r)})',
            'M': '0' if r == 5 else f'{p("BD")}-IF({p("BC")}>0,MAX(0,{prior("BD")}-{p("CE")})+{p("CF")},0)',
        }
        for col, formula in errors.items():
            ledger_col = ledger_columns[col]
            put('Cash_Flow', f'{ledger_col}{r}', active(r, formula), 'cents')
            put(sheet, f'{col}{r}', active(r, cash(ledger_col)), 'cents')
        error_range = f'AD{r}:AL{r}'
        status = f'IF(OR(MAX({error_range})>0.01,MIN({error_range})<-0.01,L{r}="Check funding convergence"),"Review","OK")'
        put('Cash_Flow', f'AM{r}', active(r, status))
        for col, formula in {'A': p('B'), 'B': p('C'), 'I': cash('L'), 'J': cash('AM')}.items():
            put(sheet, f'{col}{r}', active(r, formula), 'money')

def guidance():
    """Explain the first run and put reminders beside the relevant cells."""
    notes = {
        'README': {
            'A3': 'START HERE',
            'A4': '1. Pick a scenario on Inputs. Enter 1, 2 or 3 in B4. The matching columns are D, E and F.',
            'A5': '2. Check your ages, wages, balances and spending. Enter the age the plan should run through in row 53.',
            'A6': '3. Open Summary. Start with “Can the plan cover its costs?”',
            'A7': 'If it says no, find the first year it falls short. Use Cash_Flow to see which bills went unpaid.',
            'A8': 'Use today’s dollars unless the input note says otherwise. Enter your full annual pay raise, including inflation.',
            'A10': 'WHAT TO EDIT',
            'A11': 'Edit the blue-text cells on Inputs. Column B shows the values for the scenario you selected.',
            'A12': 'To change investment assumptions, go to Asset_Allocation.',
            'A13': 'The plan includes the end age you enter, up to age 125.',
            'A15': 'WHERE TO LOOK NEXT',
            'A16': 'Cash_Flow shows how the plan pays each year’s bills. Roth_Access shows what it can withdraw.',
            'A17': 'Healthcare shows costs and credits by year. Turn on detailed coverage in Inputs to see costs by person.',
            'A18': 'Income_Targeting compares Roth conversions for one retirement year.',
            'A19': 'Projection shows the full year-by-year plan.',
            'A20': 'Risk_Settings controls the separate report on market returns and retirement spending.',
            'A21': 'Checks flags calculation or input issues for review.',
            'A22': 'The separate risk report recalculates costs and withdrawals for each market path. Rerun it after changing inputs.',
            'A23': 'Tax_Tables lists the published tax amounts and sources.',
            'A26': 'Cash pays taxes first, then required spending. Roth IRA savings use what’s left.',
            'A30': 'The model uses a household healthcare budget by default. Turn on detailed coverage to enter costs for each person.',
            'A47': 'See Tax_Tables for the published California disability insurance rates.',
            'A53': 'Turn on detailed healthcare in Inputs to estimate Marketplace credits. Conversions, investment gains and interest can reduce the credit.',
            'A57': 'The model checks whether the alternative minimum tax adds to federal income tax. See Tax_Tables for the amounts by year.',
            'A89': 'PORTFOLIO RISK REPORT',
            'A90': 'Use the separate risk report to see how the plan holds up under different market returns.',
            'A91': 'Set the early-loss and spending options on Risk_Settings.',
            'A92': 'The report shows how often the plan pays every bill and when shortfalls first occur.',
            'A93': '',
            'A94': '',
            'A95': '',
            'A96': '',
            'A97': '',
            'A98': '',
            'A109': '',
            'A110': '',
            'A111': '',
            'A112': '',
            'A113': '',
            'A114': '',
            'A115': '',
            'A120': 'Rerun the separate risk report after changing assumptions or risk settings.',
            'A122': 'Tax_Tables shows the published rates and the years they apply. Later years use the model’s inflation assumptions where applicable.',
        },
        'Inputs': {
            'A1': 'PLAN INPUTS',
            'A2': 'Pick a scenario in B4. Edit the blue-text cells in its column. Column B shows the selected values.',
            'G4': '“Varies” marks inputs that differ across scenarios.',
            'G9': 'Enter your spouse’s age at the start of the year.',
            'G10': 'Wages stop in the year you reach this age.',
            'G11': 'Enter how many people earn wages.',
            'G12': 'Enter annual wages for both earners before taxes and deductions.',
            'G25': 'Enter the current value of investments in taxable accounts.',
            'G26': 'Enter the current value of traditional retirement accounts.',
            'G27': 'Enter the combined Roth balance today.',
            'G30': 'Enter Roth IRA contributions or conversions you can already withdraw.',
            'G32': 'Enter combined annual pre-tax 401(k) contributions.',
            'G33': 'Enter what your employer adds each year.',
            'G34': 'Enter your annual health savings account contribution.',
            'G35': 'Enter the annual Roth IRA savings goal for both people.',
            'G36': 'Enter what you spend now, including healthcare.',
            'G37': 'Enter retirement spending. Enter healthcare separately below.',
            'G38': 'Enter mortgage principal and interest already included in spending. Leave property tax and insurance in regular spending.',
            'G47': 'Enter the annual amount to move from traditional savings to Roth. Withdrawal order 2 uses the ceiling in row 49.',
            'G48': '1 uses taxable investments first. 2 uses traditional savings up to the ceiling in row 49, then taxable investments.',
            'G49': 'Enter the limit for ordinary income after the standard deduction. The model adjusts it for inflation.',
            'G51': 'Enter the adjusted cost of the investments you still own. Use your brokerage records.',
            'G52': 'Enter annual taxable income apart from wages. See rows 86–88.',
            'G53': 'Enter the last age to show in the plan, up to 125.',
            'G86': 'Enter 100% if you can spend all of the other taxable income in row 52.',
            'G87': 'Enter your age when the income starts. Use 0 to start in the first projected year.',
            'G88': 'Enter the last age the income is received. Use the same start and end age for a one-time payment.',
            'G98': 'Use 1 to apply the direct Roth IRA limit. Use 0 to leave new Roth IRA savings out.',
        },
        'Summary': {
            'A25': 'Can the plan cover its costs?',
            'C25': 'YES means the plan covers spending, healthcare and tax through the selected end age.',
            'A26': 'First year the plan falls short',
            'C26': 'Find this year in Cash_Flow to see what went unpaid.',
            'A27': 'Largest amount unpaid in one year',
            'C27': 'Includes unpaid required spending and taxes.',
            'A30': 'Estimated tax on traditional savings today',
            'C30': 'Uses the tax rate you enter on Inputs.',
            'C34': 'This is the gap in planned savings. Required costs are checked above.',
            'C42': 'Total moved from traditional savings to Roth, in the dollars of each year.',
        },
        'Income_Targeting': {
            'A2': 'Enter an age and a conversion amount to compare costs for one retirement year.',
        },
    }
    for sheet, cells_by_address in notes.items():
        for address, text in cells_by_address.items():
            put(sheet, address, text)


def main():
    src,out=map(Path,sys.argv[1:3]); wb=load_workbook(src,data_only=False)
    if any(s in wb.sheetnames for s in MODULES): raise SystemExit('Use the pre-phase-1 horizon workbook as the build baseline.')
    m=sys.modules[__name__]
    inputs(wb); cash_ledger(); roth(); funding(wb); projection(wb); reporting(); risk_cleanup.configure(m, wb); reconciliation()
    tax_policy.configure(m,wb)
    published_tax.configure(m)
    roth_contributions.configure(m)
    tax_policy.bind_projection(m,lambda r,k:fd(funding_search.columns(tax_policy)[k],group(r)+STEPS-1))
    healthcare.configure(m,wb)
    targeting.configure(m,wb)
    put('Asset_Allocation','A33','PROFILE 2 - Higher cash allocation')
    put('Asset_Allocation','A42','PROFILE 3 - Highest cash allocation')
    put('Asset_Allocation','A22','The separate risk report uses these correlations. Keep the diagonal at 1.00 and matching values on either side equal.')
    put('Asset_Allocation','A91','The separate risk report applies early losses to equity. Other assets keep their assumed returns.')
    put('Summary','B5','=COUNTIF(Inputs!$H$8:$H$69,"VARIES")+COUNTIF(Inputs!$H$86:$H$98,"VARIES")+COUNTIF(Inputs!$H$101:$H$124,"VARIES")+COUNTIF(Inputs!$H$126:$H$128,"VARIES")')
    put('Summary','A41','Marketplace MAGI as % of poverty guideline')
    rr='MAX(2,Inputs!$B$10-Inputs!$B$8+1)'
    guard=f'AND(Inputs!$B$70>={rr},Inputs!$B$101=1,INDEX(Healthcare!$O$5:$O$130,MIN(126,{rr}))>0)'
    put('Summary','B41',f'=IF({guard},INDEX(Healthcare!$AD$5:$AD$130,{rr}),"")')
    put('Summary','C41','Shown for the first retirement year when an eligible Marketplace quote is modeled.')
    put('Summary','C38','Federal and California income tax, payroll taxes, and any applicable AMT or investment income tax.')
    put('Summary','C40','Tax liability from the first year without modeled wages.')
    put('Summary','A44','HEALTHCARE AT RETIREMENT','band')
    for r, title, col in [(45,'Gross annual healthcare','P'),(46,'Marketplace premium credit','R'),(47,'Net annual healthcare','S')]:
        put('Summary',f'A{r}',title,'label')
        put('Summary',f'B{r}',f'=IF(Inputs!$B$70>={rr},INDEX(Healthcare!${col}$5:${col}$130,{rr}),"")','summary_money')
    put('Summary','C47','After Marketplace credit, before HSA reimbursement. See Healthcare for coverage by person.','note')
    put('README','A101','Set a retirement age and an amount on Income_Targeting. The comparison uses that year’s filing status, balances and healthcare.')
    put('README','A102','Marketplace estimates use the published 2026 and 2027 contribution bands. Later years hold the 2027 percentages constant.')
    put('README','A104','Compare tax and net healthcare for funded amounts. A request can be limited by balances, required distributions or funding.')
    put('README','A107','Each comparison recalculates withdrawals needed to pay tax and costs for the selected year. It does not run the request through later years.')
    put('README','A128','Payroll tax applies the Social Security cap per earner. Ordinary taxable account sales can realize gains or losses; remaining losses carry forward separately for federal and California tax.')
    put('Income_Targeting','C38','Counts sweep rows that fully fund the requested conversion.','note')
    put('Income_Targeting','C39','Counts requests limited by cash, balances or the withdrawal policy.','note')
    guidance()
    dashboard.configure(put)
    workbook_layout.write_scenarios(put)
    order = workbook_layout.sheet_order([s for s in wb.sheetnames if s not in risk_cleanup.LEGACY_SHEETS] + MODULES)
    out.write_text(json.dumps({'order':order,'cells':cells,'styles':styles,'shared':shared},separators=(',',':')))
    print('Authored cells by module:',{s:len(c) for s,c in cells.items()})

if __name__=='__main__': main()
