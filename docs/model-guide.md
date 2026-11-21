# Retirement model guide

New to the planner? [Try the guided fictional example](getting-started.md) before entering household information.

## Run your first plan

1. **Save a private copy.** Download the planner and open it in desktop Excel for Mac. Keep the `.xlsm` format so the Run button works. On **Start Here**, choose **Open Inputs**. [Download the workbook (.xlsm)](../workbook/retirement-model.xlsm?raw=true)
2. **Pick a scenario.** On **Inputs**, select a scenario in **B4**. Edit the highlighted column. Column D holds the example, E compares a later retirement and F compares spending 10% less in retirement. **Jump to** takes you to each group of inputs.
3. **Enter your figures.** Review ages, pay, savings and spending. Choose the last age to show, up to 125. Use today’s dollars unless a note says otherwise. In **Inputs B130**, enter the year your starting ages and balances apply to. The forecast begins the following year.
4. **Review Summary.** With Excel set to Automatic calculation, the base projection updates as you edit. Start with the outlook, then look at how the plan pays for retirement. If the status asks you to review an input, follow the note beneath it.
5. **Run market risk when you’re ready.** You can make several edits before running the scenarios. On **Risk_Settings**, use 100 paired paths for a quick comparison or 1,000 for the final report. In desktop Excel with macros enabled, select **Run market risk** on **Summary** to start the full calculation. Print when the status says **Risk results are current**. [Risk instructions](risk-analysis.md#run-the-analysis)

To compare plans, save Summary as a PDF before changing an assumption. Run market risk again and compare the two reports. Each completed run replaces the last one.

When sharing the workbook, send a fresh copy with fictional figures and include a copy of the [license](../LICENSE). Your working file contains household inputs for all three scenarios and saved market results.

Select any formula cell to read its calculation in the formula bar. You can edit scenario inputs, investment assumptions, the conversion comparison, risk settings and published tax values while the sheets are protected. To change a formula in your own copy, choose **Review → Unprotect Sheet**. No password is needed.

## Enter income, spending and savings

- **Pay and other income.** Enter wages before taxes and deductions. Enter the full annual raise, including inflation. A 5% raise takes $80,000 to $84,000 in the first projected year. Use row 52 for other taxable income and rows 86–88 for the share you can spend and the ages you receive it. The full amount counts as taxable income during those years. Enter 0 to start in the first projected year, or use the same start and end age for a one-time payment.
- **Interest paid in cash.** Use rows 46 and 89 for outside tax-exempt interest and its spendable share. Count that interest once across your investment returns and outside income.
- **Living costs.** Enter spending in today’s dollars. Row 14 sets annual inflation. Include working-year healthcare in current spending and enter retirement healthcare separately.
- **Mortgage.** Include principal and interest in spending and enter the same payment in row 38. This keeps the payment fixed until the mortgage ends. Keep property tax and insurance in ongoing spending.
- **Blank or zero?** Leave pay and current spending blank if the plan starts in retirement. Leave retirement spending blank if the plan ends before retirement. Enter zero for a period you’ve included with no spending. The age-100 and age-105 market tests need retirement spending if they reach retirement.
- **One-time bills.** Enter the calendar year and cost in today’s dollars in rows 75–82. Bills in the same year are added together. The base plan and each market trial adjust the amount for inflation through the year it is due. A flexible spending cut never reduces a scheduled bill.
- **Long-term care.** Enter the starting age, annual cost and number of years in rows 58–60. The model adds one period of care to healthcare costs and adjusts the cost for inflation. Use 999 as the starting age to leave it out.
- **Investment fees.** Enter the annual fee in B131. Use zero if the returns already include the fee. The model applies the fee after investment growth.
- **Survivor spending and income.** Enter your age at the first death in row 54. The model applies the survivor assumptions in the following year. Use 999 for a two-person plan throughout. In row 55, enter the survivor’s share of combined Social Security benefits. Column B shows the resulting annual benefit. Set the survivor’s spending and healthcare shares in rows 56–57.

## Enter retirement account details

| Where on Inputs | What to enter |
| --- | --- |
| Rows 27–30, 90 and 92 | Choose Roth IRA or workplace Roth for the opening balance. Enter your share in row 90. For a Roth IRA, enter the contributions and conversions already available to withdraw and each owner’s first Roth IRA tax year. |
| Rows 35, 91 and 98 | Enter the combined annual Roth IRA savings goal and your share. The rest goes to your spouse. Choose 1 to include eligible savings or 0 to leave them out. |
| Rows 93–96 | Enter each owner’s first workplace Roth tax year and the age their plan allows withdrawals. Use 0 for an unknown tax year. The age displays **Access unconfirmed** until you enter the confirmed plan age. Enter 126 to restore that setting. |
| Rows 26 and 97 | Enter one owner’s pre-tax traditional IRA and 401(k) balance. Select that owner for withdrawal rules, required distributions and the last conversion age. Conversions go into the owner’s Roth IRA. |

Choose one type of Roth account for the opening balance. The workbook uses the same owner shares for the total and for the amount available to withdraw today. Make a separate estimate if you start with both Roth IRA and workplace Roth savings, or have separately inherited accounts.

The Roth IRA savings goal rises with inflation. Contributions come from cash left after taxes and required spending. Wages, each person’s IRA limit and the income rules for direct contributions can reduce the amount saved. The model assumes these are the only IRA contributions and uses its own tax schedules to estimate income. Calculate backdoor Roth contributions and their tax separately.

## Enter healthcare costs

1. **Choose a level of detail.** Keep row 101 at 0 to use a household budget. Enter yearly costs before and after Medicare in rows 42–43. The model splits the couple’s budget evenly and switches each person’s share at 65. The survivor’s share comes from row 57.
2. **Enter costs by person if needed.** Set row 101 to 1, then open rows 102–124. Use the coverage codes below and enter each person’s confirmed Medicare start age. Use 126 for someone who will not enroll in the model.
3. **Enter annual costs at today’s prices.** Use each person’s premium before credits, including their share of a joint policy. Enter basic Medicare premiums for Parts A, B and D or Medicare Advantage in rows 108–109. Use rows 110–111 for Medigap and other supplemental premiums. The model calculates Medicare income surcharges separately.
4. **Enter out-of-pocket costs after insurance.** Use rows 112–113 for medical bills eligible for health savings account (HSA) reimbursement and rows 114–115 for other bills. Enter premiums in the premium rows. Choose the HSA owner in row 121 and enter the share of long-term care costs confirmed eligible for reimbursement in row 124. Use 0% if unsure.
5. **Check Marketplace and Medicare inputs.** Confirm each person’s eligibility for a Marketplace credit and enter the benchmark premiums. For Medicare surcharges, enter the requested years of income in rows 122–123. Resolve any missing inputs on **Healthcare** before reviewing Summary or choosing **Run market risk**.

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

## Read the results

Start on **Summary**. A shortfall means at least one cent of spending or taxes went unpaid in the year it was due. The six-page report includes your results and two pages of input and spending decisions.

| What do you want to know? | Where to look |
| --- | --- |
| How much can I use in early retirement? | **Summary** shows the brokerage and Roth savings available in the first retirement year. **Roth_Access** shows what each person can withdraw. |
| When does money run short? | Find the first shortfall on **Summary**. Open **Cash_Flow** for that year’s income, withdrawals, taxes and unpaid bills. |
| What drives the result? | **Planning Metrics** shows the savings target, accessible balances, taxes and healthcare. Before **Appendix**, use **Asset_Allocation**, **Income_Targeting**, **Risk_Settings** and **Tax_Tables** to review the assumptions. |
| How does the plan change over time? | The first two tabs after **Appendix** are **Annual Detail** and **Risk Results**. They show the annual cash flow, market tests and range of balances in the dollar view selected on **Summary**. |
| How can I check a calculation? | **Checks** compares cash, taxes and balances. The totals can agree even when bills go unpaid. **Conversion Detail** lists conversion results, **Targeting_Calc** and **Funding_Calc** show the calculation steps, and **Report_Data** supplies the figures for **Summary**. |

The ending portfolio can include workplace Roth savings the plan cannot draw by the end age. Summary shows that amount separately. **Spending benchmark before income** on Planning Metrics divides the living and healthcare budget by the withdrawal rate in row 41. To see whether you can pay each year’s bills, use the funding result. It accounts for when benefits start, taxes come due and savings become available.

The example starts Social Security at age 67. Its annual benefit is $40,000 at opening-year prices. **Cash_Flow** shows the sources separately. [How cash pays the bills](workbook-schedules.md#cash_flow-and-projection)

### Print your plan

Choose a **Dollar view** on the right side of **Summary** before you print.

- **Today’s dollars** show what the amounts can buy at the prices in your opening year. Use this view to compare future costs with your current budget.
- **Future dollars** show the amounts you expect to pay or receive in each projected year. A $42,000 living budget becomes $43,050 after one year at 2.5% inflation. The two views show the same budget at different price levels.

Select **Summary**, then choose **File → Print → Active Sheets**. The six pages include results, stress tests and both decision records. On **Inputs**, **View input decisions** opens the record of account-access years and unconfirmed workplace Roth access. Print **Annual Detail** separately for the full timeline. Before saving a personal PDF, check the scenario, opening year, dollar view and risk status. [Excel print instructions](https://support.microsoft.com/en-us/excel/print-a-sheet-or-workbook)

## Compare Roth conversions and market risk

- **Income_Targeting.** Enter a retirement age or use 0 for the first retirement year. Start with a zero conversion, then test a larger amount in that year’s dollars. Compare how much can convert, the withdrawals needed for bills, and the change in tax and healthcare costs. Select **Open conversion detail** for the full table. Apply a conversion on **Inputs** to see its effect over the full plan. [Conversion comparison](workbook-schedules.md#income_targeting)
- **Risk_Settings.** B6 sets the portfolio level that triggers a spending cut. B7 sets the cut and B8 sets the share of non-mortgage living costs you can reduce. B10 shows the combined effect. For example, a 15% cut to a 30% flexible share reduces eligible living costs by 4.5% in a year when the rule applies. [Market risk guide](risk-analysis.md#spending-flexibility)
- **Portfolio management.** Use the portfolio guide to record tax loss and gain harvesting decisions. [Portfolio management](portfolio-management.md)

## How the model works

### Withdrawals and taxes

Cash pays taxes first, including prior years’ unpaid tax. It then pays overdue spending and current bills. Remaining cash can fund Roth IRA contributions. Unpaid taxes and bills carry forward and reduce what the household has left after tax. The model assumes zero financing interest on unpaid spending.

On **Cash_Flow**, columns AR–AW separate the opening unpaid bills, repayments, current payments and newly unpaid costs. **Cumulative newly unpaid spending** counts each new bill once. The funding result records the first shortfall even if later income repays it.

Eligible Roth funds can cover a cash need during working years. The model applies each owner's account-access rules and updates the remaining contribution and conversion basis. It tests withdrawals against the available balance and skips a planned conversion when eligible savings cannot cover its costs.

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

For each coverage year, the Marketplace uses the poverty guideline available when enrollment opens. The workbook uses the published guidelines for 2026 and 2027 coverage. It estimates later guidelines using inflation through the previous year. A price increase during the coverage year first affects the following year’s estimate. [IRS eligibility rules](https://www.irs.gov/affordable-care-act/individuals-and-families/eligibility-for-the-premium-tax-credit)

The direct Roth IRA calculation uses published 2026 and 2027 amounts and estimates later amounts with inflation. [IRS 2026 retirement limits](https://www.irs.gov/pub/irs-drop/n-25-67.pdf) · [IRS 2027 retirement limits](https://www.irs.gov/pub/irs-drop/n-26-66.pdf)

The model treats qualified dividends and capital gains as ordinary income for California tax. It excludes Social Security from California taxable income, adds back HSA contributions, and applies the California exemption credit to tax. It includes the additional state tax on taxable income over $1 million and lists published disability insurance rates on **Tax_Tables**.

Payroll taxes use wages by earner. Confirm the wage split in row 126. The model tracks federal and California capital losses separately. Enter each unused loss from prior returns as a positive amount in rows 127–128.

### Investment mix and account values

Use rows 15–16 to choose investment profiles for working years and retirement. Review stocks, bonds, cash and real assets in each account on **Asset_Allocation**. Each account’s weights must total 100%. The table at the bottom shows each account’s expected return after fees for each stage of the plan. The market simulations use the volatility and correlations on the same sheet. [Investment assumptions](workbook-schedules.md#asset_allocation)

For taxable investments, enter the cost basis from your brokerage records. When the model sells part of the account, it assigns that same share of the opening basis to the sale. New deposits and reinvested income on the remaining investments add to the closing basis. If it sells the entire account and deposits new money that year, the new deposit starts with its own cost basis.

The workbook applies investment returns after deposits and withdrawals. The estimated tax rate in row 61 is used for after-tax net worth.

### Items to estimate separately

The workbook leaves the following items for you to estimate separately.

- **Survivor planning.** Life insurance, pension survivor choices, the joint tax return for the year of death, changes to cost basis at death and the transfer of a health savings account after its owner dies.
- **Federal taxes.** Qualified charitable distributions, medical-expense deductions and the temporary enhanced senior deduction.
- **State taxes.** Changes in state law, California tax on HSA earnings and California’s treatment of certain municipal interest.

For the annual cash flow, account-access and funding calculations, see the [workbook schedules](workbook-schedules.md).
