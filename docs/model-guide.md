# Retirement model guide

## Run your first plan

1. **Save a private copy.** The saved scenarios use fictional figures.
2. **Pick a scenario.** On **Inputs**, choose 1 for the example, 2 to retire later or 3 to test 10% lower retirement spending. Edit the highlighted scenario column. **Jump to** takes you to each group of inputs.
3. **Enter your figures.** Review ages, pay, savings and spending. Choose the last age to show, up to 125. Use today’s dollars unless a note says otherwise. Use the **+** signs beside the rows to open optional settings.
4. **Review Summary.** Start with the outlook, then look at how the plan pays for retirement. If the status asks you to review an input, follow the note beneath it.

Set the year your starting ages and balances apply to in **Inputs B130**. The first full forecast year is the following year.

## Enter income, spending and savings

- **Pay and other income.** Enter wages before taxes and deductions. Enter the full annual raise, including inflation. A 5% raise takes $80,000 to $84,000 in the first projected year. Use row 52 for other taxable income and rows 86–88 for the share you can spend and the ages you receive it. The full amount counts as taxable income during those years. Use 0 for the starting age to begin in the first projected year. Use the same start and end age for a one-time payment.
- **Interest paid in cash.** Use rows 46 and 89 for outside tax-exempt interest and its spendable share. Count that interest once across your investment returns and outside income.
- **Spending.** Enter living costs in today’s dollars. The model adjusts those costs each year by the inflation rate in row 14. Include working-year healthcare in current spending and enter retirement healthcare separately. Include mortgage principal and interest in spending and repeat that payment in row 38 so the model can hold it fixed and remove it after the last payment. Keep property tax and insurance in ongoing spending.
- **One-time bills.** Enter the calendar year and cost in today’s dollars in rows 75–82. These expenses apply to all scenarios.
- **Long-term care.** Enter the starting age, annual cost and number of years in rows 58–60. The model adds one period of care to healthcare costs and adjusts the cost for inflation. Use 999 as the starting age to leave it out.
- **Investment fees.** Enter the annual fee in B131. Use zero if the returns already include the fee. The model applies the fee after investment growth.
- **Survivor spending and income.** Enter your age at the first death in row 54. Use 999 for a two-person plan. In row 55, enter the share of combined Social Security benefits the survivor will receive. Column B shows the resulting benefit. Set the survivor’s spending and healthcare shares in rows 56–57.

## Enter retirement account details

| Where on Inputs | What to enter |
| --- | --- |
| Rows 27–30, 90 and 92 | Choose Roth IRA or workplace Roth for the opening balance. Enter your share in row 90. For a Roth IRA, enter the contributions and conversions already available to withdraw and each owner’s first Roth IRA tax year. |
| Rows 35, 91 and 98 | Enter the combined annual Roth IRA savings goal and your share. The rest goes to your spouse. Choose 1 to include eligible savings or 0 to leave them out. |
| Rows 93–96 | Enter each owner’s first workplace Roth tax year and the age their plan allows withdrawals. Use 0 for an unknown tax year or 126 as the withdrawal age to exclude that balance from withdrawals. |
| Rows 26 and 97 | Enter one owner’s pre-tax traditional IRA and 401(k) balance. Select that owner for withdrawal rules, required distributions and the last conversion age. Conversions go into the owner’s Roth IRA. |

The opening Roth fields support one account type and apply the same owner split to the balance and the amount available today. Use a separate estimate for opening savings that mix Roth IRA and workplace Roth balances. The workbook uses one traditional balance. Separately inherited accounts need their own estimate.

The Roth IRA savings goal rises with inflation. Funding uses cash left after taxes and required spending and is limited by wages, each person’s IRA limit and the income rules for direct contributions. The model assumes these are the only IRA contributions and uses its own tax schedules to estimate income. Calculate backdoor Roth contributions and their tax separately.

## Read the results

- **Summary** shows whether the scenario covers its costs through the end age. A shortfall means at least one cent of spending or taxes goes unpaid. The five-page report includes the outlook, retirement cash flow, risk and decisions, and two pages that record your inputs.
- **Planning Metrics** groups the figures behind the report, including the savings target, accessible balances, funding, taxes and healthcare. Open the tab next to **Summary** or use its **Planning metrics** link. Dollar figures are labeled with their units.
- **Annual Detail** shows the same cash measures and portfolio balance over the full plan. It uses the dollar view on **Summary**. The opening row gives the starting balance. The values and row shading stop at the end age on **Inputs**.
- **Asset_Allocation**, **Income_Targeting**, **Risk_Settings**, **Risk Results** and **Tax_Tables** come before Appendix. Use them to review decisions and the tax rates the model uses.
- **Appendix** links to the detailed schedules that support the report.
- **Cash_Flow** shows tax, spending and withdrawals by year. Use the first shortfall year on Summary to find the unpaid amounts.
- **Roth_Access** shows what each person can withdraw. The early retirement amount on Summary shows the brokerage and Roth savings the model can draw from in the first retirement year.
- **Checks** compares cash, tax and account totals. A year can reconcile and still have unpaid bills. **Conversion Detail** shows the results for each conversion amount. **Targeting_Calc** and **Funding_Calc** show the steps behind the results. **Report_Data** holds the annual values used in Summary.

The ending portfolio can include workplace Roth savings the plan cannot draw by the end age. Summary shows that amount separately. The savings target uses annual spending after the mortgage ends, healthcare, and the withdrawal rate in row 41. Use the funding result to answer whether the scenario pays its modeled bills.

The example starts Social Security at age 67. Its annual benefit is $40,000 at opening-year prices. **Cash_Flow** shows the sources separately. [How cash pays the bills](workbook-schedules.md#cash_flow-and-projection)

### Print your plan

Choose a **Dollar view** on the right side of **Summary** before you print.

- **Today’s dollars** show what the amounts can buy at the prices in your opening year. Use this view to compare future costs with your current budget.
- **Future dollars** show the amounts you expect to pay or receive in each projected year. A $42,000 living budget becomes $43,050 after one year at 2.5% inflation. The two views show the same budget at different price levels.

Select **Summary** in desktop Excel, then choose **File → Print → Active Sheets**. The five pages include the results and both decision records. **View input decisions** on Inputs jumps to those records on Summary. Print **Annual Detail** separately for the full timeline in the same dollar view. Check the scenario, opening year, and dollar view before saving a PDF with personal figures. [Excel print instructions](https://support.microsoft.com/en-us/excel/print-a-sheet-or-workbook)


## Enter healthcare costs

1. **Choose a level of detail.** Keep row 101 at 0 to use a household budget. Enter yearly costs before and after Medicare in rows 42–43. The model splits the couple’s budget evenly and switches each person’s share at 65. The survivor’s share comes from row 57.
2. **Enter costs by person if needed.** Set row 101 to 1, then open rows 102–124. Use the coverage codes below and enter each person’s confirmed Medicare start age. Use 126 for someone who will not enroll in the model.
3. **Enter annual costs at today’s prices.** Use each person’s premium before credits, including their share of a joint policy. Enter basic Medicare premiums for Parts A, B and D or Medicare Advantage in rows 108–109. Use rows 110–111 for Medigap and other supplemental premiums. The model calculates Medicare income surcharges separately.
4. **Enter out-of-pocket costs after insurance.** Use rows 112–113 for medical bills eligible for health savings account (HSA) reimbursement and rows 114–115 for other bills. Enter premiums in the premium rows. Choose the HSA owner in row 121 and enter the share of long-term care costs confirmed eligible for reimbursement in row 124. Use 0% if unsure.
5. **Check Marketplace and Medicare inputs.** For Marketplace coverage, confirm whether each person qualifies for a credit and enter the applicable benchmark quotes. For Medicare surcharges, enter the requested years of income in rows 122–123. Open **Healthcare** to resolve missing inputs before reading Summary or running the risk report.

| Coverage code in rows 102–103 | Meaning |
| --- | --- |
| 0 | No coverage entered |
| 1 | Marketplace coverage |
| 2 | Other coverage |
| 3 | COBRA continuation coverage |
| 4 | Marketplace coverage with confirmed ineligibility for Medicare after age 65 |

For the Marketplace credit, enter 1 after confirming eligibility or 0 otherwise. Confirm eligibility against employer and government coverage, including Medicaid and Medicare. For rows 118–120, use annual quotes for the second-lowest-cost silver plan for both eligible people, you alone or your spouse alone, as applicable.

For the Medicare income entries, add federal adjusted gross income and tax-exempt interest. Row 122 uses the calendar year before the opening year. Row 123 uses the opening year. Enter -1 for an unknown amount. Enter 0 when the amount was zero.

Use **Inputs B132** for a separate medical inflation rate, or leave it blank to use general inflation. **Healthcare** shows yearly premiums, credits and HSA payments. **Cash_Flow** uses those amounts to show what remains to pay.

## Compare Roth conversions and market risk

- **Income_Targeting.** Enter a retirement age or use 0 for the first retirement year. Start with a zero conversion, then test a larger amount in that year’s dollars. Compare how much can convert, the withdrawals needed for bills, and the change in tax and healthcare costs. Select **Open conversion detail** for the full table. Apply a conversion on **Inputs** to see its effect over the full plan. [Conversion comparison](workbook-schedules.md#income_targeting)
- **Risk_Settings.** Estimate the share of retirement living costs you could reduce and how much you would cut from that share. The next row shows what that means for the full living budget. Refresh market risk after changes. Read the success rate and funding examples on Summary, then open **Risk Results** to see the details. [Portfolio risk analysis](risk-analysis.md)
- **Portfolio management.** Use the portfolio guide to record tax loss and gain harvesting decisions. [Portfolio management](portfolio-management.md)

## How the model works

### Withdrawals and taxes

Cash pays current and unpaid prior-year taxes first, then required spending. Roth IRA savings use the cash left over. Unpaid tax carries forward and reduces after-tax net worth. Unpaid spending is reported in the year it occurs. The model tests withdrawal amounts against available balances and skips a planned conversion when eligible savings cannot cover its costs.

Withdrawal order 1 uses taxable investments first. Order 2 uses traditional savings up to the ordinary taxable income limit in row 49, then taxable investments. With order 2, Roth conversions use any room left under that limit. The limit is entered after the standard deduction and rises with inflation. The model counts 85% of Social Security for this withdrawal limit and calculates actual tax from the benefit tax formula.

Set the age when required minimum distributions begin in row 67. The workbook uses the IRS Uniform Lifetime Table to calculate them. Required withdrawals above what the plan spends go into the taxable account. Use row 68 to model fixed early withdrawals under section 72(t). The payment uses the rate in row 69 and stays fixed for at least five years or until age 60, whichever is later. The workbook uses whole ages for the age-59½ rule and pauses Roth conversions during that period.

### When Roth savings become available

The workbook tracks regular Roth IRA contributions, conversion principal and earnings separately for each owner. Conversion principal generally becomes available in the model after five tax years or at age 60. Earnings require the five-year account period and age 60. The workbook uses age 60 for the age-59½ rule because it projects whole years.

Workplace Roth requires verified plan access and a qualified withdrawal. Use a separate estimate for rollovers, an age-55 separation exception or early nonqualified workplace Roth withdrawals. [IRS Roth account comparison](https://www.irs.gov/retirement-plans/ten-differences-between-a-roth-ira-and-a-designated-roth-account) · [IRS IRA distribution guidance](https://www.irs.gov/publications/p590b)

### How income affects Social Security taxes

The workbook uses other adjusted gross income, tax-exempt interest and half of Social Security benefits to estimate how much of the benefit is taxable. Up to 85% of the benefit can count as taxable income. Tax-exempt municipal interest also counts in the income tests for Marketplace credits and Medicare surcharges.

After the first death, the workbook uses single-filer rules from the following year. The survivor’s Social Security amount comes from the share of combined benefits entered in row 55. Spending and healthcare use the shares in rows 56–57.

### Tax rates and healthcare credits

**Tax_Tables** lists published rates and the years they apply. The model estimates later amounts where applicable. The tax calculation includes federal income tax, California income tax, payroll tax, the alternative minimum tax and the net investment income tax. See the [annual tax data review](annual-tax-review.md) for the sources and update steps. For higher-income Medicare enrollees, the workbook calculates the income surcharge using income from two years earlier. A Roth conversion can raise future Medicare costs.

The Marketplace estimate uses published 2026 and 2027 contribution percentages and keeps the 2027 percentages for later years. Income for the credit includes items such as traditional withdrawals, Roth conversions, investment gains and tax-exempt interest. A gain can raise that income even when its federal tax rate is zero.

The direct Roth IRA calculation uses published 2026 and 2027 amounts and estimates later amounts with inflation. [IRS 2026 retirement limits](https://www.irs.gov/pub/irs-drop/n-25-67.pdf) · [IRS 2027 retirement limits](https://www.irs.gov/pub/irs-drop/n-26-66.pdf)

The model treats qualified dividends and capital gains as ordinary income for California tax. It excludes Social Security from California taxable income, adds back HSA contributions, and applies the California exemption credit to tax. It includes the additional state tax on taxable income over $1 million and lists published disability insurance rates on **Tax_Tables**.

Payroll taxes use wages by earner. Confirm the wage split in row 126. The model tracks federal and California capital losses separately. Enter each unused loss from prior returns as a positive amount in rows 127–128.

### Investment mix and account values

Use rows 15–16 to choose investment profiles for working years and retirement. Review stocks, bonds, cash and real assets in each account on **Asset_Allocation**. Each account’s weights must total 100%. The model uses the mix to calculate account returns and income. The market simulations use the volatility and correlations on the same sheet. [Investment assumptions](workbook-schedules.md#asset_allocation)

For taxable investments, enter the cost basis from your brokerage records. Reinvested income and savings increase it. The model reduces it in proportion to what you sell and taxes price gains when realized. The model uses annual investment income and the applicable tax treatment for interest, real estate investment trust distributions and qualified dividends.

The workbook applies investment returns after deposits and withdrawals. The estimated tax rate in row 61 is used for after-tax net worth.

### Items to estimate separately

Use separate estimates for life insurance, pension survivor choices, the joint tax return for the year of death, the adjustment to cost basis at death and an HSA transfer after the owner’s death. The tax schedules also leave out qualified charitable distributions, medical-expense deductions, the federal senior deduction, state-law changes, California tax on HSA earnings and California’s separate treatment of certain municipal interest.

See the Phase 2 release record for sources, assumptions and verification. [Phase 2 release record](release-notes/phase-2.md)
