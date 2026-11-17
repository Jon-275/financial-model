"""Independent market paths using the workbook's documented integer stream."""
from dataclasses import dataclass
from math import exp,fsum,isfinite
from native_random import normal_draws
from .paths import ReturnModel


@dataclass(frozen=True,slots=True)
class MarketPath:
    returns: tuple
    inflation: tuple


class NativeReturnModel(ReturnModel):
    """Use common Box–Muller shocks for five assets and optional inflation."""

    def market_path(self,ages,retirement_age,seed,trial,volatility_multiplier=1,
                    return_basis=1,baseline_returns=None,*,inflation=.025,
                    vary_inflation=False,inflation_sd=.015,persistence=.5,
                    lower=-.02,upper=.10):
        parameters=self._log_parameters(volatility_multiplier,return_basis)
        baseline=self.baseline_path(ages,retirement_age,baseline_returns)
        if any(not isfinite(value) for value in (inflation,inflation_sd,persistence,lower,upper)):
            raise ValueError('Inflation assumptions must be finite numbers.')
        if inflation<=-1 or inflation_sd<0 or not 0<=persistence<1 or lower<=-1 or lower>upper:
            raise ValueError('Review the inflation variation, persistence and bounds.')
        if vary_inflation and not lower<=inflation<=upper:
            raise ValueError('Expected inflation must be within its annual bounds.')
        rates=[];prices=[];last=inflation
        for age,original,normal in zip(ages,baseline,normal_draws(seed,trial+1,len(ages)),strict=True):
            correlated=tuple(fsum(coefficient*normal[j] for j,coefficient in enumerate(row))
                             for row in self._factor)
            gross=tuple(mean*exp(shift+scale*z) for mean,(scale,shift),z in
                        zip(self._asset_gross,parameters,correlated,strict=True))
            rates.append(self._apply(age>=retirement_age,original,gross))
            last=(min(upper,max(lower,inflation+persistence*(last-inflation)+
                               inflation_sd*(1-persistence*persistence)**.5*normal[5]))
                  if vary_inflation else inflation)
            prices.append(last)
        return MarketPath(tuple(rates),tuple(prices))

    def random_path(self,ages,retirement_age,seed,trial,volatility_multiplier=1,
                    return_basis=1,baseline_returns=None):
        return self.market_path(ages,retirement_age,seed,trial,volatility_multiplier,
                                return_basis,baseline_returns).returns
