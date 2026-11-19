# Workbook calculations and decision tools

Enter household figures on **Inputs** and read the plan on **Summary**. Print **Summary** alone for the six-page report, including your input and spending decisions. Before **Appendix**, you can review **Planning Metrics**, investment assumptions, Roth conversions, market settings and tax tables. The divider links to the supporting schedules. [Model guide](model-guide.md)

## Planning Metrics

**Planning Metrics**, next to **Summary**, shows the savings target, accessible balances, funding, taxes and healthcare. **Annual Detail** is the first tab after **Appendix**. It shows the opening position and each year through the selected end age, up to 125. It follows the dollar view on Summary and can be printed separately.

## Asset_Allocation

Choose profiles for the working and retirement years on **Inputs** in rows 15 and 16. On **Asset_Allocation**, review the stock, bond, cash and real-asset weights for each taxable, traditional, Roth and HSA account. Edit the blue cells. Each account column must total 100%.

**Investment assumptions used by the plan** shows each account’s expected return after fees and before inflation. The portfolio return weights each account by its opening balance. The taxable return includes interest and qualified dividends. Their yields also determine taxable income. Volatility sets how much returns vary. Correlations set how closely they move together.

## Income_Targeting

Use **Income_Targeting** to test a Roth conversion in one retirement year. In **Retirement age to compare**, enter 0 for the first retirement year or choose a later age. Enter the **Conversion to test** in that year's dollars. Compare the result with a zero conversion. Look at funded costs, withdrawals, taxes and healthcare together.

**Open conversion detail** shows the results for a range of conversion amounts. The display limit controls that range. To see the effect on later balances and Medicare costs, enter your planned conversion on **Inputs** and review **Summary**.

## Risk_Settings and Risk Results

On **Risk_Settings**, B6 sets the trigger as a percentage of the base plan’s opening balance for that year. B8 is the share of living costs you could reduce. B7 sets how much to cut from that share. B10 shows the combined cut to non-mortgage living costs. Mortgage payments, healthcare and scheduled bills remain due. The rule is checked each year.

Choose **Run market risk** on **Summary** after changing the plan. In **Risk Results**, read the fixed and flexible funding rates alongside the number of paths that needed spending cuts and the number of years those cuts lasted. The median includes flexible paths with no cuts. The early-loss tests show how sustained stock losses at retirement affect each spending policy. [Market risk guide](risk-analysis.md)

## Tax_Tables

**Tax_Tables** identifies the years and values used for U.S. federal tax, California tax, payroll contributions, Medicare and Marketplace coverage. Federal and California calculations use separate rates, deductions and capital-loss balances. Review each section's year before updating a value. The federal estimate includes the additional standard deduction at age 65. The temporary enhanced senior deduction is listed for reference and excluded from the estimate. [Tax assumptions](model-guide.md#tax-rates-and-healthcare-credits) · [Annual tax data review](annual-tax-review.md)

## Cash_Flow and Projection

**Cash_Flow** shows each year's future dollars. **Open income sources** separates pay after pre-tax contributions, Social Security, spendable other taxable income and cash received from outside tax-exempt interest. The model adds the Marketplace premium credit to those sources. HSA payments and withdrawals from savings appear separately.

1. The HSA pays eligible healthcare costs up to its available balance.
2. Cash income and available withdrawals pay taxes due, including prior unpaid tax, then prior unpaid spending and current bills.
3. During working years, remaining cash can fund eligible Roth IRA savings. Any cash left goes into taxable savings.

To investigate a shortfall on Summary, find that year on **Cash_Flow**. Columns AR–AW show prior unpaid spending, repayments, current bills paid and newly unpaid amounts. An amount unpaid in its due year still counts as a shortfall after repayment. Outstanding bills carry forward without financing interest.

**Projection** applies annual returns after deposits and withdrawals. Column CE shows the taxable cost basis removed by a sale. Column CF adds new deposits and reinvested income on the investments still held. After a full sale, a new deposit starts with its own cost basis.

## Roth_Access

**Roth_Access** tracks IRA and workplace Roth balances separately for each owner. Select one account type for the combined opening Roth balance on **Inputs**. Use separate estimates if the household starts with both types. For a Roth IRA, the schedule distinguishes contributions, conversion principal and earnings. For a workplace Roth, it uses the owner's first contribution year and the age the employer plan permits withdrawals. The model uses age 60 for the age-59½ rule because it projects whole years.

One owner's pre-tax traditional IRA and 401(k) savings share a balance and a set of withdrawal rules. Inspect traditional withdrawals on **Projection** and **Cash_Flow**. [Retirement account inputs](model-guide.md#enter-retirement-account-details)

## Healthcare

**Healthcare** uses either the household budgets for the years before and after Medicare or the costs entered for each person. The detailed schedule uses coverage, premiums, out-of-pocket costs and Medicare enrollment age. It calculates any Marketplace premium credit using income after withdrawals or conversion. It also accounts for long-term care and estimates Medicare income surcharges using income from two years earlier. Check this sheet if a coverage or income entry needs review. [Healthcare inputs](model-guide.md#enter-healthcare-costs)

## Checks and Report_Data

**Checks** reconciles income, withdrawals, payments, deposits, tax due and account balances. It also checks Roth access, taxable cost basis, the Marketplace credit and HSA use. The totals can agree even if bills went unpaid. A **Review** status flags a difference or a funding calculation to inspect.

**Report_Data** supplies the annual values used on **Summary**, including chart values in the selected dollar view. The figures run from the opening balance through the selected end age.

## Funding_Calc and Targeting_Calc

**Funding_Calc** tests withdrawals against spending, taxes and any Marketplace credit. It checks credit thresholds when they apply. If the plan cannot cover a requested conversion, it also tests withdrawals without the conversion. **Cash_Flow** uses the selected draw.

**Targeting_Calc** repeats the funding calculation for each proposed conversion in the year selected on **Income_Targeting**. Each test starts with the same balances and bills. **Conversion Detail** shows the outcomes by amount.

In either calculation table, filter **Calculation step** to **Selected draw** for the final withdrawal, or exclude **Skipped** to see only the steps used. Keep the rows in calculation order. Columns BD onward link to the calculation inputs.

## Funding_Paths and the trial schedules

**Funding_Paths** follows the low, median and high fixed-spending examples from income and bills through withdrawals, ending balances and unpaid costs. It uses Summary's dollar view. In **Funding_Path_Data**, match the trial number and year to see the saved values in that year's dollars.

On **Risk_Trials**, enter a trial number in B3 to inspect its results. **Random_Numbers** shows how the seed produces repeatable draws. **Market_Paths** shows that trial's returns, inflation and spending cuts. The **Trial_** schedules apply the same account, healthcare, tax and cash-flow formulas as the base plan. **Risk_Data** uses all completed trials to calculate report statistics and checks the run against the current assumptions.

On **Trial_Checks**, columns N–Y compare closing balances with the amounts calculated from opening balances, transactions and returns. All twelve differences should be zero in an active year. If column J says **Review**, inspect those differences and the cash-flow status in **Trial_Cash_Flow** column AM for the same year. [How the runner checks results](macro-runner.md)
