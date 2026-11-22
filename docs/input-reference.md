# Inputs reference

This guide follows the rows on **Inputs**. Choose a scenario in **B4**, then edit its shaded cells. Column B shows the active values. Use today's dollars unless a cell says otherwise.

[Workbook guide](model-guide.md#inputs) · [Calculation rules](model-methods.md)

## Ages, pay and assumptions

**Rows 7–16.**

Name the scenario, enter both ages and choose your retirement age. Enter wages before taxes and deductions. Annual wage growth includes inflation: 5% takes $80,000 to $84,000 in the next projected year.

Choose investment profiles for working years and retirement in rows 15–16. Review each profile on **Asset_Allocation**.

## Accounts and savings

**Rows 25–35.**

Enter opening account balances and annual savings. Include each owner's first Roth IRA tax year and the Roth IRA contributions and conversions already available to withdraw.

Choose one type for the opening Roth balance: **1 = Roth IRA**, **2 = workplace Roth**. Use a separate estimate for a mix of both types or inherited accounts. Rows 90–98 set ownership and access details.

The annual Roth IRA savings target is for both people. It rises with inflation, but available cash, wages, contribution limits and income rules can reduce the amount saved. The model assumes these are the only IRA contributions. Estimate backdoor Roth contributions and tax separately.

## Spending and healthcare budgets

**Rows 36–43.**

| Input | What to include |
| --- | --- |
| Current spending | Working-year living costs, including healthcare. |
| Retirement spending | Living costs. Enter retirement healthcare separately. |
| Mortgage principal and interest | Include it in spending and enter the same payment in row 38. Enter its final year and months paid in rows 39–40. This keeps that payment fixed until it ends. Keep property tax and insurance in ongoing spending. |
| Withdrawal rate | Used for the spending benchmark on Planning Metrics. It does not establish whether each year's bills are covered. |
| Healthcare before and after Medicare | Household annual budgets, used when detailed healthcare is off. The model splits the budget evenly and switches each person's share at 65. |

Leave pay and current spending blank if the plan starts in retirement. Leave retirement spending blank if the plan ends before retirement. Enter **0** for an included period with no spending. Tests through age 100 or 105 still need retirement spending if they reach retirement.

## Benefits, conversions and end age

**Rows 44–53.**

Enter combined Social Security, claim age, outside tax-exempt interest, planned Roth conversions, withdrawal order and the taxable income ceiling. Enter the brokerage cost basis from your records and any other taxable income. Choose an end age up to **125**.

Rows 86–89 set how much outside income is available as cash and when other taxable income starts and stops. Count tax-exempt interest once across investment returns and outside income.

**Income_Targeting** compares a conversion in one year. Enter a planned conversion here to see its effect over the full plan. [Withdrawal and conversion rules](model-methods.md#withdrawals-and-taxes)

## Survivor, care and withdrawal settings

**Rows 54–69.**

- **Survivor:** enter your age at the first death in row 54, or **999** to keep a two-person plan. Survivor assumptions start the following year. Set shares of combined Social Security, spending and healthcare in rows 55–57. Column B shows the resulting annual Social Security amount.
- **Long-term care:** enter start age, annual cost and duration in rows 58–60. Use **999** as the start age to leave it out. The model adds one period of care and adjusts its cost for inflation.
- **Traditional account tax estimate:** row 61 sets the rate used for after-tax net worth.
- **Required withdrawals:** row 67 sets the start age.
- **Fixed early withdrawals:** rows 68–69 set the 72(t)/SEPP option and rate. Review the [model's withdrawal rules](model-methods.md#withdrawals-and-taxes) before using it.

## One-time expenses

**Rows 75–82.**

Enter a description, calendar year and cost in today's dollars. These entries apply to all scenarios. Bills in the same year are added together and adjusted for inflation until due. Flexible spending cuts do not reduce them.

## Cash and account access

**Rows 86–98.**

| Rows | What to enter |
| --- | --- |
| 86–88 | The cash share of other taxable income and its start and end ages. The full income counts as taxable during those years. Use start age **0** for the first projected year, or the same start and end age for a one-time payment. |
| 89 | The cash share of outside tax-exempt interest. |
| 90–92 | Your share of the opening Roth balance, your share of Roth IRA savings and your spouse's first Roth IRA tax year. The remaining shares belong to your spouse. |
| 93–96 | Each owner's first workplace Roth tax year and confirmed withdrawal age. Use tax year **0** if unknown. Age **126** means **Access unconfirmed**; replace it with the confirmed plan age. |
| 97 | The owner of the combined pre-tax IRA and 401(k) balance. This owner determines withdrawal rules, required distributions and the last conversion age. Conversions go to that owner's Roth IRA. |
| 98 | **1** to include eligible Roth IRA savings, **0** to leave them out. |

The opening Roth balance and its currently available amount use the same owner shares. The model projects whole years and uses age **60** for the age-59½ rule. [Account-access rules and limits](model-methods.md#when-roth-savings-become-available)

## Healthcare by person

**Rows 101–124.**

Keep row 101 at **0** to use household budgets in rows 42–43. Set it to **1** to enter costs for each person. Use annual costs at opening-year prices.

| Rows | What to enter |
| --- | --- |
| 102–103 | Pre-Medicare coverage code from the table below. |
| 104–105 | Confirmed Medicare start age. Use **126** for someone who will not enroll in the model. |
| 106–107 | Premiums before credits, including each person's share of a joint policy. |
| 108–111 | Basic Medicare or Medicare Advantage premiums, then Medigap and other supplemental premiums. Income surcharges are calculated separately. |
| 112–115 | Out-of-pocket costs after insurance, split into costs eligible for HSA reimbursement and other costs. Keep premiums in the premium rows. |
| 116–117 | **1** after confirming Marketplace-credit eligibility, **0** otherwise. Check employer and government coverage, including Medicaid and Medicare. |
| 118–120 | Annual second-lowest-cost silver-plan quotes for both eligible people, you only and your spouse only. |
| 121 | The HSA owner. |
| 122–123 | Federal adjusted gross income plus tax-exempt interest for the year before the opening year, then the opening year. Use **-1** if unknown or **0** if it was zero. |
| 124 | The share of long-term care costs confirmed eligible for HSA reimbursement. Use **0%** if unsure. |

| Coverage code | Meaning |
| --- | --- |
| 0 | No coverage entered |
| 1 | Marketplace |
| 2 | Other coverage |
| 3 | COBRA |
| 4 | Marketplace with confirmed Medicare ineligibility after 65 |

Open **Healthcare** and resolve missing coverage or income inputs before running market risk.

## Tax, year and fee inputs

**Rows 126–132.**

| Row | What to enter |
| --- | --- |
| 126 | Your share of combined wages, used for payroll taxes by earner. |
| 127–128 | Unused federal and California capital losses from prior returns, as positive amounts. |
| B130 | The year your starting ages and balances apply to. The forecast begins the following year. |
| B131 | Annual investment fee. Use **0** if the returns already include it. The fee is applied after investment growth and affects all scenarios. |
| B132 | Medical inflation, or blank to use general inflation. |

Check the status messages on Inputs before returning to **Summary**.
