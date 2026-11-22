# How the model calculates the plan

[Workbook guide](model-guide.md) · [Input reference](input-reference.md) · [Calculation sheets](workbook-schedules.md)

Use this reference to understand a calculation. For steps to use the workbook, follow the workbook guide.

## Withdrawals and taxes

Cash pays taxes first, including prior years’ unpaid tax. It then pays overdue spending and current bills. Remaining cash can fund Roth IRA contributions. Unpaid taxes and bills carry forward and reduce what the household has left after tax. The model assumes zero financing interest on unpaid spending.

On **Cash_Flow**, columns AR–AW separate the opening unpaid bills, repayments, current payments and newly unpaid costs. **Cumulative newly unpaid spending** counts each new bill once. The funding result records the first shortfall even if later income repays it.

Eligible Roth funds can cover a cash need during working years. The model applies each owner's account-access rules and updates the remaining contribution and conversion basis. It tests withdrawals against the available balance and skips a planned conversion when eligible savings cannot cover its costs.

Withdrawal order 1 uses taxable investments first. Order 2 uses traditional savings up to the ordinary taxable income limit in row 49, then taxable investments. With order 2, Roth conversions use any room left under that limit. The limit is entered after the standard deduction and rises with inflation. The model counts 85% of Social Security for this withdrawal limit and calculates actual tax from the benefit tax formula.

Set the age when required minimum distributions begin in row 67. The workbook uses the IRS Uniform Lifetime Table to calculate them. Required withdrawals above what the plan spends go into the taxable account. Use row 68 to model fixed early withdrawals under section 72(t). The payment uses the rate in row 69 and stays fixed for at least five years or until age 60, whichever is later. The workbook uses whole ages for the age-59½ rule and pauses Roth conversions during that period.

## When Roth savings become available

The workbook tracks regular Roth IRA contributions, conversion principal and earnings separately for each owner. Conversion principal generally becomes available in the model after five tax years or at age 60. Earnings require the five-year account period and age 60. The workbook uses age 60 for the age-59½ rule because it projects whole years.

Workplace Roth requires verified plan access and a qualified withdrawal. Use a separate estimate for rollovers, an age-55 separation exception or early nonqualified workplace Roth withdrawals. [IRS Roth account comparison](https://www.irs.gov/retirement-plans/ten-differences-between-a-roth-ira-and-a-designated-roth-account) · [IRS IRA distribution guidance](https://www.irs.gov/publications/p590b)

## How income affects Social Security taxes

The workbook uses other adjusted gross income, tax-exempt interest and half of Social Security benefits to estimate how much of the benefit is taxable. Up to 85% of the benefit can count as taxable income. Tax-exempt municipal interest also counts in the income tests for Marketplace credits and Medicare surcharges.

After the first death, the workbook uses single-filer rules from the following year. The survivor’s Social Security amount comes from the share of combined benefits entered in row 55. Spending and healthcare use the shares in rows 56–57.

## Tax rates and healthcare credits

**Tax_Tables** lists published rates and the years they apply. The model estimates later amounts where applicable. The tax calculation includes federal income tax, California income tax, payroll tax, the alternative minimum tax and the net investment income tax. The federal estimate includes the additional standard deduction at age 65. See the [annual tax review](annual-tax-review.md) for sources and updates. For higher-income Medicare enrollees, the workbook calculates the income surcharge using income from two years earlier. A Roth conversion can raise future Medicare costs.

The Marketplace estimate uses published 2026 and 2027 contribution percentages and keeps the 2027 percentages for later years. Income for the credit includes items such as traditional withdrawals, Roth conversions, investment gains and tax-exempt interest. A gain can raise that income even when its federal tax rate is zero.

For each coverage year, the Marketplace uses the poverty guideline available when enrollment opens. The workbook uses the published guidelines for 2026 and 2027 coverage. It estimates later guidelines using inflation through the previous year. A price increase during the coverage year first affects the following year’s estimate. [IRS eligibility rules](https://www.irs.gov/affordable-care-act/individuals-and-families/eligibility-for-the-premium-tax-credit)

The direct Roth IRA calculation uses published 2026 and 2027 amounts and estimates later amounts with inflation. [IRS 2026 retirement limits](https://www.irs.gov/pub/irs-drop/n-25-67.pdf) · [IRS 2027 retirement limits](https://www.irs.gov/pub/irs-drop/n-26-66.pdf)

The model treats qualified dividends and capital gains as ordinary income for California tax. It excludes Social Security from California taxable income, adds back HSA contributions, and applies the California exemption credit to tax. It includes the additional state tax on taxable income over $1 million and lists published disability insurance rates on **Tax_Tables**.

Payroll taxes use wages by earner. Confirm the wage split in row 126. The model tracks federal and California capital losses separately. Enter each unused loss from prior returns as a positive amount in rows 127–128.

## Investment mix and account values

Each account uses the investment profile selected for working years or retirement. The portfolio return weights accounts by their opening balances. Taxable returns include interest and qualified dividends; their yields also determine taxable income. The market trials use the volatility and correlations on **Asset_Allocation**. [Review investment assumptions](model-guide.md#asset_allocation)

For taxable investments, enter the cost basis from your brokerage records. When the model sells part of the account, it assigns that same share of the opening basis to the sale. New deposits and reinvested income on the remaining investments add to the closing basis. If it sells the entire account and deposits new money that year, the new deposit starts with its own cost basis.

The workbook applies investment returns after deposits and withdrawals. The estimated tax rate in row 61 is used for after-tax net worth.

## How the Monte Carlo simulation works

Monte Carlo means running the same plan through many different sequences of market returns. One **path** is a year-by-year market sequence for the whole plan. The workbook tests each random path twice: once with fixed spending and once with the spending-cut rule.

### From inputs to results

1. **Read the assumptions.** The workbook uses the selected scenario, each account's investment mix, asset returns, volatility and correlations. Risk_Settings sets the path count, seed, spending cuts and inflation options.
2. **Generate repeatable random draws.** The seed and trial number set the random-number sequence on **Random_Numbers**. The same version, inputs and settings reproduce the same sampled markets.
3. **Build annual asset returns.** **Market_Paths** turns the draws into yearly returns for US stocks, international stocks, bonds, cash and property. Correlation settings link their movements. Asset growth uses a lognormal distribution. Each account uses its working-year or retirement mix and the plan's fee assumptions.
4. **Build the price path.** If variable inflation is on, a separate draw changes inflation around the entered rate, with persistence and bounds. Each path's prices feed spending, healthcare and indexed tax calculations. Inflation varies independently of the asset draws.
5. **Calculate the household's cash flow.** The **Trial_** sheets calculate each year's income, account access, healthcare, withdrawals, taxes, bills and balances. These use the base plan's annual formulas. A trial can run short before its end age even if its ending portfolio is positive.
6. **Compare spending policies.** Fixed and flexible spending use the same markets and inflation. The flexible policy applies the spending-cut rule when its trigger is met. The workbook saves both outcomes; it can reuse a fixed result after proving that the flexible policy would change nothing.
7. **Check and summarize.** The runner checks the trial balances before saving a complete batch. **Risk_Data** calculates the statistics shown on **Summary** and **Risk Results**. Failed or incomplete runs do not become current report results.

### How the success rate is calculated

```text
Fixed-spending success rate = successful random fixed-spending paths / random path count
```

For example, 950 successful paths out of 1,000 gives **95%**. This is an illustration, not the example plan's result. Success means every bill and tax was paid on time through the end age. Even one cent unpaid counts as a shortfall, including a bill repaid later. Flexible spending has its own success rate. The 33 stress outcomes are separate from these rates.

The balance chart calculates percentiles at each age. Its shaded band can use different paths at different ages. The low, median and high funding examples instead follow three complete fixed-spending paths, ranked by ending balance in today's dollars. [Result definitions](risk-analysis.md#read-the-summary)

### Trace one path in the workbook

Start with a completed run whose results are current. In **Risk_Trials B3**, select trial **1**. Follow its draws on **Random_Numbers**, returns and prices on **Market_Paths**, then cash flow and balances on the **Trial_** sheets. Select **1001** to compare its flexible-spending partner under the same markets. Enter **0** to return to the base plan. [Trial inspection details](macro-runner.md#inspect-a-trial)

The simulation uses fixed return, volatility and correlation assumptions rather than historical crisis replays. It does not estimate the chance that its assumptions are correct. More paths reduce sampling variation; they do not remove these limits.

<details>
<summary>Random-number and return formulas</summary>

**Random_Numbers** uses the `xoshiro128**` generator, with each trial's starting state derived from the seed and trial number. Each projected year uses six independent normal draws, generated from uniform draws by the Box–Muller method. Five feed asset returns; the sixth feeds optional inflation variation.

**Market_Paths** applies a Cholesky factor of the correlation matrix to the five asset draws. It then calculates each asset return as:

```text
return = (1 + entered return) × exp(log shift + log scale × correlated draw) − 1
```

The scale uses annual volatility and the market-variation multiplier. Under **compound growth**, the entered gross return anchors the median and the log shift is zero. Under **simple annual average**, the log shift is minus half the log variance so the entered gross return anchors the mean. Account returns apply the sampled change in the account's asset mix to its base gross return.

Inspect the factor and return parameters on **Market_Paths** beside the annual table. The source defines the exact scale calculations: [workbook formulas](../tools/native_risk.py), [random draws](../tools/native_random.py), [independent comparison](../tools/risk_engine/native_paths.py) and [return parameters](../tools/risk_engine/paths.py).

</details>

## Market returns and inflation

The base plan and market trials use each account’s investment mix, including the HSA, and the fee on **Inputs**. Enter expected returns before inflation. The change in buying power is `(1 + return after fees) / (1 + inflation) - 1`. The annual raise on **Inputs** already includes inflation.

The **Investment return convention** on **Risk_Settings** selects compound growth or a simple annual average. The model lets returns on different assets move together. It uses a lognormal distribution and keeps expected returns, volatility and correlations fixed throughout the plan. Only the listed stress tests impose a particular sequence of losses. The model does not replay historical crises.

Choose **Vary annual inflation?** to let inflation change around your expected rate. You can set how much it varies, how long a change tends to last and its upper and lower limits. The model varies inflation independently of investment returns. A market loss does not automatically trigger higher inflation.

Each path updates living costs, medical prices, indexed tax parameters, Social Security and purchasing-power measures from its own price sequence. Fixed-dollar costs and tax thresholds that do not rise with inflation stay fixed. After Social Security starts, the model uses annual price changes and keeps benefits at their previous high if prices fall. Actual benefit adjustments use quarterly CPI-W figures.

In the future-dollar view, the balance band uses separately calculated nominal percentiles. A representative path’s amounts use that path’s prices at the relevant age. The base plan continues to use its expected inflation rate.


## Longer-retirement comparisons

The tests through age 100 and 105 keep spending fixed. On **Market_Paths**, the first extra year can still compare with the base plan's final balance. Later comparison cells are blank and the actual spending cut stays at zero.
