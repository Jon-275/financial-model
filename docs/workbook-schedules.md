# Workbook calculations and decision tools

Use **Summary** to read and print the plan. Use **Annual Detail** for the full timeline in the same dollar view. The tabs before **Appendix** help you review decisions. The appendix links to the annual schedules, reference amounts and calculation tables.

## Cash_Flow and Projection

**Cash_Flow** shows future dollars for each year. Select **Open income sources** to see pay after pre-tax 401(k) and HSA contributions, Social Security, the spendable share of other income, and the cash share of outside tax-exempt interest. Add the Marketplace premium credit to those four sources to get **Cash income**. Summary counts the share of that income used to pay bills and taxes. Withdrawals and HSA payments are shown under savings used.

1. The HSA pays eligible healthcare costs up to the available HSA balance.
2. Cash income and available withdrawals pay current taxes and any tax carried in from the prior year. The remaining cash pays the year’s bills.
3. Any cash left can fund eligible Roth IRA savings during working years. Remaining cash goes to the taxable account. Required retirement distributions can also leave cash to reinvest.
4. The model records unpaid bills and carries unpaid tax into the next year. **Projection** applies investment returns to the balances after those transactions.

## Roth_Access

Choose one account type for the combined opening Roth balance. It can be a **Roth IRA** or a **workplace Roth**, such as a Roth 401(k). Use separate estimates if your household starts with both. The schedule tracks the two types separately as the plan adds Roth IRA contributions or conversions. For Roth IRAs, it tracks contributions, conversion amounts that become available over time, and the value of the account. For workplace Roth it uses each owner’s first contribution year and the age the plan permits withdrawals. The model uses age 60 for age 59½ because it runs in whole years. [IRS Roth account comparison](https://www.irs.gov/retirement-plans/ten-differences-between-a-roth-ira-and-a-designated-roth-account)

The traditional balance on **Inputs** holds one owner’s pre-tax IRA and 401(k) money. The model uses one set of access and required-distribution assumptions for that balance. Plan-specific withdrawal rules, after-tax IRA basis, and separate owners need their own estimates. Use **Projection** and **Cash_Flow** to inspect traditional withdrawals. [IRS required distribution rules](https://www.irs.gov/retirement-plans/retirement-plan-and-ira-required-minimum-distributions-faqs)

## Healthcare

The household budget option uses the amounts entered before and after Medicare and allocates them by person. Detailed coverage uses each person’s premiums, out-of-pocket costs, Medicare start age and coverage type. It adds long-term care and estimates the applicable Medicare income surcharge using income from two years earlier. The Marketplace credit uses eligible premiums, the benchmark quote, household size and income after the modeled withdrawals or conversion. The HSA pays eligible expenses up to its available balance. The rest goes into the cash-flow calculation. [Healthcare inputs](model-guide.md#enter-healthcare-costs) · [IRS premium tax credit](https://www.irs.gov/credits-deductions/premium-tax-credit-ptc-overview) · [Social Security’s Medicare income rules](https://secure.ssa.gov/apps10/poms.nsf/lnx/0601101010)

## Asset_Allocation

Start with the investments you expect to hold in taxable, pre-tax, Roth and HSA accounts. Compare the working and retirement profiles on **Inputs**, then inspect the stock, bond, cash and real-asset weights on **Asset_Allocation**. Edit the blue cells and make each account column total 100%. The selected profile checks flag a missing or invalid mix before the projection runs.

Expected returns are annual rates before adjusting for inflation. The workbook weights the rates for each account, including the HSA, and applies the fee on Inputs. Taxable account yields also determine interest and dividends used in the tax calculation. Volatility and how assets move together affect the market simulations. Use these assumptions to explore how a different mix changes the plan’s funding result and balance range.

## Income_Targeting

Choose a retirement year. Enter zero for a no-conversion comparison or enter an amount in that year’s dollars. Use **Funding result** to check whether both cases can pay their required costs. When both are funded, compare total withdrawals and the change in tax and net healthcare. A conversion can increase income and reduce the Marketplace credit for that year. The sweep shows where a larger conversion raises costs or leaves bills unpaid.

The comparison starts from the active scenario’s balances and rules for one year. Enter the planned conversion on **Inputs** to see how it affects later balances, taxes and Medicare surcharges over the full timeline.

## Risk_Settings and Risk Results

List the living costs you could reduce, then decide how much you would cut from those costs and when you would act. The trigger compares the opening balance in a simulated year with the balance for that year in the base plan. Multiply the flexible share by the cut to that share to see the reduction to the full living budget. Mortgage, healthcare, tax and one-time bills remain part of the funding calculation.

Read the success rate and number of runs with spending cuts together. The rule tests the balance each year, so a cut can recur for several years. Check the early stock tests for the age of the first unpaid costs and whether the cut changes it. Excel opens with the fictional example’s results on **Summary**. Run the simulator locally after editing the plan. It returns a refreshed copy of the same workbook with the updated figures on **Summary**. [Market risk guide](risk-analysis.md)

## Tax_Tables

Use the links on the sheet to find federal, California, payroll, Medicare and Marketplace amounts. The tables identify the years used by the model. The projection uses published amounts for their applicable years and estimates later values according to the documented rules. Federal and California income taxes use separate rates, deductions and loss balances. The model estimates U.S. federal and California taxes for the scenario. [Tax assumptions](model-guide.md#tax-rates-and-healthcare-credits) · [Annual tax data review](annual-tax-review.md)

The federal estimate uses the additional standard deduction at age 65. The separate temporary enhanced senior deduction is listed in Tax_Tables as a reference. Eligible households can have lower federal tax than the model estimates for 2026–2028. [IRS enhanced senior deduction](https://www.irs.gov/newsroom/check-your-eligibility-for-the-new-enhanced-deduction-for-seniors)

## Checks and Report_Data

**Checks** compares income and withdrawals with cash payments and deposits, tax due with paid and unpaid tax, and account balances with transfers and withdrawals. It also checks Roth access, negative balances, the Marketplace credit, eligible HSA costs and taxable cost basis. **Review** means a difference exceeds one cent or the funding calculation flagged the year. The checks can balance in a year with unpaid bills because the model records the unpaid amounts.

**Report_Data** supplies the values for the chart and tables. It links back to the annual schedules and applies the selected age and dollar view. Change the assumptions on Inputs and the decision tools.

## Funding_Calc and Targeting_Calc

**Funding_Calc** tries different withdrawals to find where cash can cover spending and taxes. Each trial recalculates the taxes and Marketplace credit that would follow from the withdrawal. The model checks key income thresholds and refines the withdrawal amount. **Cash_Flow** shows the chosen row’s result. Blank rows are unused trials, optional calculations or years outside the plan.

**Targeting_Calc** runs the same funding calculation for the selected year and each requested conversion. The comparison holds the starting balances and required bills fixed. Blank rows arise when a trial does not apply or a requested amount falls outside the comparison range. The calculation tables display dollars and cents while retaining full precision. Read **Income_Targeting** for the selected-year result.
