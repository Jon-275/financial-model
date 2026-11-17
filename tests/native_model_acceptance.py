"""Compare a selected Excel trial with the independent annual calculation."""
import argparse
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
from native_risk import MODEL_SHEETS
from risk_engine.native_annual import NativeAnnualModel,price_path
from risk_engine.paths import Guardrail,RiskInputs
from risk_engine.native_paths import NativeReturnModel
from risk_engine.scenario import load_scenario
from model_acceptance import validate_model
from xlsx_acceptance import scan_errors
from xlsx_read import Workbook


class TrialScope:
    def __init__(self,book):
        self.book=book
        self.parts={**book.parts,**{source:book.parts[trial] for source,trial in MODEL_SHEETS.items()}}

    def __getitem__(self,name):return self.book[MODEL_SHEETS.get(name,name)]


def validate_trial(path,reference):
    scenario=load_scenario(reference)
    calendar=NativeAnnualModel(scenario)
    allocation=scenario.allocation
    returns=NativeReturnModel(RiskInputs(allocation.means,allocation.volatilities,allocation.correlation,
                                        allocation.working,allocation.retirement))
    with Workbook(path) as book:
        scan_errors(book)
        scope=TrialScope(book)
        ledger=validate_model(scope,include_reports=False)
        trial=book['Risk_Trials']['B3'].value
        settings=book['Risk_Settings']
        assert book['Market_Paths']['J3'].value=='OK'
        end_age=int(scope['Inputs']['B53'].value)
        years=end_age-calendar.opening_age
        rate=.04 if trial==2201 else .05 if trial in (2202,2205) else scenario.inputs[14]
        annual=price_path(calendar,[rate]*years,end_age=end_age,
                          medical_floor=18000 if trial==2206 else 25000 if trial==2207 else 0)
        baseline=price_path(calendar,[scenario.inputs[14]]*(int(book['Inputs']['B53'].value)-calendar.opening_age),
                            end_age=int(book['Inputs']['B53'].value))
        planned=baseline.planned_openings(baseline.run(baseline.baseline_returns))
        flexible=1001<=trial<=2000 or 2101<=trial<=2113
        kwargs=({'guardrail':Guardrail(*(scenario.sequence[i] for i in (6,7,8))),
                 'planned_openings':planned} if flexible else {})
        if 1<=trial<=2000:
            path=returns.market_path(annual.ages,annual.retirement_age,int(settings['B30'].value),
                                      (trial-1)%1000,float(settings['B31'].value),scenario.return_basis,annual.baseline_returns,
                                      inflation=scenario.inputs[14],vary_inflation=settings['B32'].value==1,
                                      inflation_sd=settings['B33'].value,persistence=settings['B34'].value,
                                      lower=settings['B35'].value,upper=settings['B36'].value)
            annual=price_path(calendar,path.inflation,end_age=end_age)
            market=path.returns
        elif 2001<=trial<=2013 or 2101<=trial<=2113:
            market=returns.sequence_path(annual.ages,max(annual.opening_age+1,annual.retirement_age),
                                         int(settings['B5'].value),scenario.stress_returns[(trial-1)%100],annual.baseline_returns)
        elif trial==2205:
            gross=tuple(1+max(-1,mean-.02) for mean in allocation.means)
            market=tuple(returns._apply(age>=annual.retirement_age,original,gross)
                         for age,original in zip(annual.ages,annual.baseline_returns,strict=True))
        else:
            market=annual.baseline_returns
        expected=annual.run(market,audit=True,**kwargs)
        maximum=0;count=0
        for offset,(result,period) in enumerate(zip(expected.years,annual.periods,strict=True),6):
            for column,value in (('G',period.cpi),('H',period.tax_index)):
                assert abs(scope['Projection'][f'{column}{offset}'].value-value)<1e-10,(trial,column,offset)
            for sheet,values in result.audit.items():
                for column,value in values.items():
                    actual=scope[sheet][f'{column}{offset}'].value
                    assert isinstance(actual,(int,float)),(trial,sheet,column,offset,actual)
                    difference=abs(actual-value)
                    assert difference<=.01,(trial,sheet,column,offset,actual,value,difference)
                    maximum=max(maximum,difference);count+=1
        print(f'PASS trial {trial}: {count} independent annual values; largest difference ${maximum:.8f}; '+
              f'{sum(row.real_cut>0 for row in expected.years)} years with cuts',flush=True)
        if trial>=2201:
            if trial in (2201,2202,2205):
                rate=.04 if trial==2201 else .05
                for row in range(6,5+int(scope['Inputs']['B70'].value)):
                    assert abs(scope['Projection'][f'G{row}'].value-(1+rate)**(row-5))<1e-10
            if trial in (2203,2204,2205):
                assert scope['Inputs']['B53'].value>= (100 if trial==2203 else 105)
            if trial in (2206,2207):
                minimum=18000 if trial==2206 else 25000
                for row in range(6,5+int(scope['Inputs']['B70'].value)):
                    if scope['Projection'][f'E{row}'].value=='Working':continue
                    price=book['Market_Paths'][f'I{row+1}'].value
                    survivor=scope['Inputs']['B57'].value if scope['Projection'][f'F{row}'].value=='Single' else 1
                    gross=scope['Healthcare'][f'P{row}'].value
                    ltc=scope['Healthcare'][f'K{row}'].value
                    assert gross-ltc+.01>=minimum*price*survivor
            print(f'PASS stress {trial}: {ledger}',flush=True)


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbooks',type=Path,nargs='+')
    parser.add_argument('--reference',type=Path,default=Path('workbook/retirement-model.xlsm'))
    args=parser.parse_args()
    for path in args.workbooks:validate_trial(path,args.reference)
