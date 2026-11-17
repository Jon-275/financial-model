"""Recalculate price-view examples with independent expected report amounts."""
import argparse
from math import isclose
from pathlib import Path
import subprocess
import sys
import tempfile

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
from openpyxl import Workbook,load_workbook
import plan_report
from risk_presentation import annual_dollars,case_dollars,early_dollars


def add_income_timing_cases(source):
    formulas = {}

    def collect(sheet, address, value, style=None):
        if sheet == 'Summary' and address in ('B58', 'B59'):
            formulas[address] = value

    plan_report.configure(collect)
    cases = (
        # Opening age, retirement age, claim age, end age, age at death, survivor note.
        (35, 60, 67, 95, 80, 'From age 81,'),
        (35, 60, 67, 80, 80, None),
        (70, 60, 67, 95, 55, 'From age 71,'),
        (35, 60, 99, 95, 999, None),
    )
    expected = {}
    for index, (opening, retire, claim, end, death, note) in enumerate(cases):
        name = f'income-timing-{index}.xlsx'
        book = Workbook()
        summary = book.active
        summary.title = 'Summary'
        for address, formula in formulas.items():
            summary[address] = formula
        inputs = book.create_sheet('Inputs')
        for row, value in {8: opening, 10: retire, 44: 40000, 45: claim,
                           53: end, 54: death, 56: 0.75}.items():
            inputs[f'B{row}'] = value
        data = book.create_sheet('Report_Data')
        data['Q2'] = end - opening + 1
        data['Q4'] = 13
        if claim > end:
            timing = 'Social Security starts after the selected end age.'
        elif claim <= max(retire, opening + 1):
            timing = 'Social Security is included from the first projected retirement year.'
        else:
            timing = f'Cash income and accessible savings must cover retirement costs until Social Security starts at age {claim}.'
        survivor = f'{note} the plan uses single-filer tax rules and 75% of the couple’s living costs.' if note else None
        book.save(source / name)
        expected[name] = (timing, survivor)
    return expected


def verify(executable):
    with tempfile.TemporaryDirectory(prefix='retirement-report-') as folder:
        root=Path(folder);source=root/'source';output=root/'calculated'
        source.mkdir();output.mkdir()
        expected={}
        for native in (False,True):
            for future in (False,True):
                name=f'{"native" if native else "legacy"}-{"future" if future else "today"}.xlsx'
                book=Workbook();book.active.title='Summary'
                for sheet in ('Inputs','Projection','Risk_Data','Check'):book.create_sheet(sheet)
                book['Summary']['K18']='Future dollars' if future else 'Today’s dollars'
                book['Inputs']['B8']=40
                data=book['Risk_Data']
                data['Z2']='retirement-risk-native-v1' if native else 'retirement-risk-workbook-v4'
                data['BB4']=43;data['AH5']=42;data['AG5']=400;data['AI5']=20
                data['AY5']=1.21;data['AZ5']=1.331;data['P5']=1.21;data['Q5']=1.1025
                for row in range(5,9):book['Projection'][f'G{row}']=1.02**(row-5)
                for column,value in zip(('D','E','F','BI','BJ','BK'),(100,200,300,105,230,350)):
                    data[f'{column}7']=value
                formulas=(case_dollars('Risk_Data!AG5',21,ending=True),
                          case_dollars('Risk_Data!AI5',21),
                          annual_dollars('D','BI',7),annual_dollars('E','BJ',7),annual_dollars('F','BK',7),
                          early_dollars('10','42',5),early_dollars('15','42',5,flexible=True))
                for row,formula in enumerate(formulas,1):book['Check'][f'A{row}']='='+formula
                book.save(source/name)
                expected[name]=([400,20,100,200,300,10,15] if not future else
                    [400*1.331,20*1.21,105,230,350,10*1.21,15*1.1025] if native else
                    [400*1.02**3,20*1.02**2,100*1.02**2,200*1.02**2,300*1.02**2,10*1.02**2,15*1.02**2])
        timing_cases = add_income_timing_cases(source)
        result=subprocess.run([executable,f'-env:UserInstallation={(root/"profile").as_uri()}',
                               '--headless','--convert-to','xlsx','--outdir',str(output),
                               *(str(path) for path in source.glob('*.xlsx'))],capture_output=True,timeout=120)
        if result.returncode:raise RuntimeError('Report formula recalculation failed.')
        for name,values in expected.items():
            book=load_workbook(output/name,read_only=True,data_only=True)
            actual=[book['Check'][f'A{row}'].value for row in range(1,len(values)+1)]
            book.close()
            for number,(value,target) in enumerate(zip(actual,values,strict=True),1):
                assert isinstance(value,(int,float)) and isclose(value,target,abs_tol=1e-9),(name,number,value,target)
        for name, pair in timing_cases.items():
            book = load_workbook(output / name, read_only=True, data_only=True)
            actual = tuple(book['Summary'][address].value or None for address in ('B58', 'B59'))
            book.close()
            assert actual == pair, (name, actual, pair)
        print('PASS: 28 report amounts in both dollar views and eight income and survivor timing checks.')


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--soffice',default='soffice')
    verify(parser.parse_args().soffice)
