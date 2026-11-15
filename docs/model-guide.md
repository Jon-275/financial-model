# Retirement model guide

## Run your first plan

1. **Save a private copy.** The saved scenarios use fictional figures.
2. **Pick a scenario.** On **Inputs**, enter 1 in B4 for the example plan, 2 to retire later or 3 to test 10% lower retirement spending. Edit the blue-text cells in the matching column, D, E or F.
3. **Review your figures.** Check ages, wages, savings and spending. Enter the last age to show in row 53, up to 125. Use today’s dollars unless a note says otherwise. Open optional settings with the **+** signs beside the rows.
4. **Read Summary.** It shows whether the plan covers its costs, the first year it falls short and the ending balance in today’s dollars. Use **Cash_Flow** to see which bills went unpaid.

Set the year your starting ages and balances apply to in **Inputs B130**. The first full forecast year is the following year.

## Enter income, spending and savings

- **Pay and other income.** Enter wages before taxes and deductions. Enter the full annual raise, including inflation. A 5% raise takes $80,000 to $84,000 in the first projected year. Use row 52 for other taxable income and rows 86–88 for the share you can spend and the ages you receive it. The full amount counts as taxable income during those years. Use 0 for the starting age to begin in the first projected year. Use the same start and end age for a one-time payment.
- **Interest paid in cash.** Use rows 46 and 89 for outside tax-exempt interest and its spendable share. Count that interest once across your investment returns and outside income.
- **Spending.** Include working-year healthcare in current spending. Enter retirement healthcare separately. Include mortgage principal and interest in spending and enter that amount again in row 38 so the model can hold it fixed and remove it after the last payment. Keep property tax and insurance in ongoing spending.
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
| Row 97 | Choose the owner of the traditional retirement account. The model uses that person’s age for withdrawals, required distributions and the last year of Roth conversions. Conversions go into that owner’s Roth IRA. |

The opening Roth fields support one account type and apply the same owner split to the balance and the amount available today. Use a separate estimate for opening savings that mix Roth IRA and workplace Roth balances. The workbook uses one traditional balance. Separately inherited accounts need their own estimate.

The Roth IRA savings goal rises with inflation. Funding uses cash left after taxes and required spending and is limited by wages, each person’s IRA limit and the income rules for direct contributions. The model assumes these are the only IRA contributions and uses its own tax schedules to estimate income. Calculate backdoor Roth contributions and their tax separately.

## Read the results

- **Summary** shows the plan result first. The balance chart marks retirement and uses today’s dollars. The spending chart shows how income and withdrawals pay bills around retirement. A shortfall means at least one cent of required spending or tax goes unpaid. If you see **CHECK PLAN**, follow the note beside it.
- **Cash_Flow** shows tax, spending and withdrawals by year. Start with the first year shown on Summary and look at unpaid spending and unpaid tax.
- **Roth_Access** shows what each person can withdraw. For an early retirement, check withdrawals year by year. The amount available at retirement is a starting balance. Later income, taxes and the age when each account becomes available also affect funding.
- **Checks** compares the model’s cash, tax and account totals. A year can pass those checks and still have unpaid bills. **Funding_Calc** shows the withdrawal trials, **Targeting_Calc** shows the conversion trials and **Report_Data** supplies the Summary charts.

The ending portfolio can include workplace Roth savings the plan cannot draw by the end age. Summary shows that amount separately. The savings target is an estimate based on spending after the mortgage ends and the withdrawal rate in row 41. Use the funding answer to see whether the plan pays its actual bills.

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

- **Income_Targeting.** Enter a retirement age or use 0 for the first retirement year. Enter the conversion amount in that year’s dollars. Compare the amount the plan can convert, the withdrawals needed to cover costs, and the change in tax and healthcare costs. Account balances, required distributions and available cash can limit the conversion. Change the main scenario to see effects on later years.
- **Risk_Settings.** Use the separate risk report to see how often the plan pays all costs under different market returns. It compares fixed spending with temporary cuts to flexible spending under the same returns. Run it again after changing the plan or risk settings. [Portfolio risk analysis](risk-analysis.md)
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

Use rows 15–16 to choose the investment mix for working years and retirement. Check the stock share on **Asset_Allocation** when comparing profiles. The mix sets the return for each account and stays fixed within each period. The separate risk report uses the volatility and correlations on that sheet.

For taxable investments, enter the cost basis from your brokerage records. Reinvested income and savings increase it. The model reduces it in proportion to what you sell and taxes price gains when realized. The model uses annual investment income and the applicable tax treatment for interest, real estate investment trust distributions and qualified dividends.

The workbook applies investment returns after deposits and withdrawals. The estimated tax rate in row 61 is used for after-tax net worth.

### Items to estimate separately

Use separate estimates for life insurance, pension survivor choices, the joint tax return for the year of death, the adjustment to cost basis at death and an HSA transfer after the owner’s death. The tax schedules also leave out qualified charitable distributions, medical-expense deductions, the federal senior deduction, state-law changes, California tax on HSA earnings and California’s separate treatment of certain municipal interest.

See the Phase 2 release record for sources, assumptions and verification. [Phase 2 release record](release-notes/phase-2.md)
