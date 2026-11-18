"""Independent annual calculation for the native workbook's price paths."""
from copy import copy
from dataclasses import replace
from math import isfinite

from .annual import AnnualModel
from roth_contributions import annual_limits


class NativeAnnualModel(AnnualModel):
    """Keep the annual funding policy and add an explicit medical cash reserve."""

    medical_floor=0

    def _advance(self,state,period,rates,cut,audit):
        if self.medical_floor and not period.working:
            i=self.inputs
            medical_price=period.sheet.health['_native_medical_index']
            ltc=period.sheet.health['K']*period.sheet.health['_native_medical_ratio']
            surcharge=self._surcharge(state,period)
            floor=self.medical_floor*medical_price*(i[57] if period.single else 1)
            extra=max(0,floor-(period.medical+surcharge-ltc))
            period=replace(period,medical=period.medical+extra)
        return super()._advance(state,period,rates,cut,audit)


def price_path(model,inflation_rates,*,end_age=None,medical_floor=0):
    """Reprice indexed cash flows before rolling the account and tax ledger.

    Pay growth remains the entered nominal rate, inclusive of inflation.
    Mortgage payments and statutory fixed thresholds remain nominal amounts.
    The annual CPI approximation holds Social Security benefits at their
    prior high after benefits begin. The workbook documents the CPI-W lag
    that this annual approximation omits.
    """
    if end_age is None:end_age=model.ages[-1]
    length=end_age-model.opening_age
    if length<1 or length>len(model.periods) or len(inflation_rates)!=length:
        raise ValueError('The independent calendar must cover the selected end age.')
    rates=tuple(inflation_rates)
    if any(not isfinite(r) or r<=-1 for r in rates):
        raise ValueError('Each inflation rate must be finite and greater than -100%.')
    if not isfinite(medical_floor) or medical_floor<0:
        raise ValueError('The medical reserve must be a nonnegative annual amount.')
    if medical_floor and not isinstance(model,NativeAnnualModel):
        raise TypeError('Use NativeAnnualModel for a medical reserve stress.')
    result=copy(model)
    result.inputs={**model.inputs}
    i=result.inputs
    base_inflation=model.inputs[14]
    i[14]=rates[0] if model.opening_year==2026 else base_inflation
    i[53]=end_age;i[70]=length+1
    result.medical_floor=medical_floor
    cpi=medical_index=1
    previous_tax_index=(1+base_inflation)**(model.opening_year-2026)
    benefit_peak=1 if model.opening_age>=i[45] else 0
    constant=len(set(rates))==1
    periods=[]
    for elapsed,(period,rate) in enumerate(zip(model.periods[:length],rates,strict=True),1):
        cpi=(1+rate)**elapsed if constant else cpi*(1+rate)
        med_rate=(i[132] if rate==base_inflation
                  else (1+i[132])*(1+rate)/(1+base_inflation)-1)
        medical_index=(1+med_rate)**elapsed if constant else medical_index*(1+med_rate)
        tax_index=((1+base_inflation)**(period.sheet.year-2026)
                   if constant and rate==base_inflation
                   else cpi*(1+base_inflation)**(model.opening_year-2026))
        tax=result.tables.for_year(period.sheet.year,i[14],tax_index,period.single,period.age,period.spouse_age)
        payroll,sdi=tax.payroll_taxes(period.sheet.projection['BT'],period.sheet.projection['BU'])
        deferral=min(max(0,period.wages),i[32]*cpi) if period.working else 0
        owner_age=period.age if i[121]==1 else period.spouse_age
        medicare_age=i[104] if i[121]==1 else i[105]
        may_contribute=(not result.detailed_health or
                        (i[121]==1 or not period.single) and
                        (medicare_age==126 or owner_age<medicare_age))
        hsa=min(max(0,period.wages-deferral),i[34]*cpi) if period.working and may_contribute else 0
        match=i[33]*cpi if period.working else 0
        if period.age>=i[45]:benefit_peak=max(benefit_peak,cpi)
        social=((i[55] if period.single else i[44])*benefit_peak if period.age>=i[45] else 0)
        other=period.other_income*cpi/period.cpi
        exempt=period.exempt*cpi/period.cpi
        cash=period.wages-deferral-hsa+social+other*i[86]+exempt*i[89]
        fixed=period.core_spending-period.flexible_base
        flexible=period.flexible_base*cpi/period.cpi
        med_ratio=medical_index/(1+i[132])**elapsed
        aca_actual,benchmark,fpl=period.aca
        if fpl and period.sheet.year>result.scenario.tax['A214']:
            fpl=result.scenario.tax['B215' if period.single else 'C215']*previous_tax_index
        previous_tax_index=tax_index
        limits=annual_limits(result.scenario.tax,period.sheet.year,tax_index,i[14],period.single)
        if limits is None:
            limit,catchup=(result.scenario.tax[cell]*cpi for cell in ('B61','B62'))
            phaseout=None
        else:limit,catchup,low,high=limits;phaseout=(low,high)
        health={**period.sheet.health,'_native_medical_index':medical_index,'_native_medical_ratio':med_ratio}
        periods.append(replace(period,sheet=replace(period.sheet,health=health),year_tax=tax,
            cpi=cpi,tax_index=tax_index,deferral=deferral,hsa_contribution=hsa,match=match,
            social_security=social,other_income=other,exempt=exempt,cash_income=cash,
            payroll=payroll,sdi=sdi,
            core_spending=fixed+flexible,flexible_base=flexible,one_off=period.one_off*cpi/period.cpi,
            medical=period.medical*med_ratio,qualified_medical=period.qualified_medical*med_ratio,
            aca=(aca_actual*med_ratio,benchmark*med_ratio,fpl),roth_target=i[35]*cpi if period.working else 0,
            savings_capacity=(limit+(catchup if period.age>=50 else 0),
                              0 if period.single else limit+(catchup if period.spouse_age>=50 else 0)),
            roth_phaseout=phaseout))
    result.periods=tuple(periods)
    result.ages=tuple(p.age for p in periods)
    result.baseline_returns=tuple(p.baseline for p in periods)
    return result
