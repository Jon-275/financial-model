# Calculation sheets

<a id="planning-metrics"></a>
<a id="asset_allocation"></a>
<a id="income_targeting"></a>
<a id="risk_settings-and-risk-results"></a>
<a id="tax_tables"></a>

These tabs appear after **Appendix**, in the order listed below. Use the links on Appendix to open a sheet. For the earlier tabs, see the [workbook guide](model-guide.md).

The base plan and market trials share the same annual calculation rules. [Model assumptions](model-methods.md) · [Market result definitions](risk-analysis.md#read-the-summary)

## Understand a calculation

| Calculation | Explanation |
| --- | --- |
| Withdrawals, bills and taxes | [Cash-flow rules](model-methods.md#withdrawals-and-taxes) |
| Account balances and cost basis | [Investment and account calculations](model-methods.md#investment-mix-and-account-values) |
| Roth withdrawals | [When savings become available](model-methods.md#when-roth-savings-become-available) |
| Tax and healthcare estimates | [Tax rates and credits](model-methods.md#tax-rates-and-healthcare-credits) |
| Monte Carlo simulation | [From random draws to the success rate](model-methods.md#how-the-monte-carlo-simulation-works) |

The sheet list below shows where to inspect each calculation in Excel.

## Report detail

| Tab | Use it to |
| --- | --- |
| Annual Detail | Follow income, spending, withdrawals, taxes and unpaid costs from the opening year to the end age. It follows Summary's dollar view and can be printed separately. |
| Risk Results | Read funding outcomes, early-loss tests and balances by age. It follows Summary's dollar view. |

## Annual schedules

| Tab | Use it to |
| --- | --- |
| Cash_Flow | Trace income, HSA payments, withdrawals, taxes and bills in each year's dollars. Find the year of a shortfall here. |
| Projection | Follow account balances and returns after deposits and withdrawals. Columns CE–CF track taxable cost basis removed and added. |
| Roth_Access | Check each owner's Roth contributions, conversions, earnings and available withdrawals. |
| Healthcare | Review premiums, Marketplace credits, Medicare surcharges, HSA payments and long-term care costs. |

### Cash_Flow and Projection

The HSA pays eligible healthcare first, up to its available balance. Income and available withdrawals then pay taxes, prior unpaid spending and current bills. During working years, cash left after these costs can fund eligible Roth IRA savings; the rest goes to taxable savings.

On **Cash_Flow**, use **Open income sources** to see pay after pre-tax contributions, Social Security and spendable outside income. The model also includes any Marketplace premium credit. Columns **AR–AW** separate prior unpaid spending, repayments, current bills paid and new unpaid amounts. Bills unpaid in their due year still count as shortfalls after repayment. Unpaid spending carries forward without interest.

On **Projection**, a sale removes the same share of taxable cost basis as the share of the account sold. New deposits and reinvested income on remaining investments add to basis. After a full sale, a new deposit starts with its own basis.

### Roth_Access

The opening Roth balance uses one account type, chosen on Inputs. One owner's pre-tax IRA and 401(k) savings also share a balance and withdrawal rules. See the [input reference](input-reference.md#cash-and-account-access) for ownership and access settings.

## Calculation checks

| Tab | Use it to |
| --- | --- |
| Checks | Reconcile cash, tax, accounts, cost basis, healthcare and HSA use. A **Review** status flags a difference or funding calculation to inspect. Balanced totals do not prove that all bills were paid. |

<a id="checks-and-report_data"></a>
<a id="funding_calc-and-targeting_calc"></a>

## Supporting data and calculations

| Tab | Use it to |
| --- | --- |
| Conversion Detail | Compare the conversion amounts tested for the year selected on Income_Targeting. |
| Report_Data | Trace the values and chart data shown on Summary in its selected dollar view. |
| Funding_Calc | Inspect the withdrawal tests for each projected year, including taxes, Marketplace credits and whether a conversion can be funded. |
| Targeting_Calc | Inspect the same funding tests for each proposed conversion. Each amount starts with the same balances and bills. |
| Risk_Data | Review saved market statistics and whether the completed run matches the current assumptions. |

In **Funding_Calc** or **Targeting_Calc**, filter **Calculation step** to **Selected draw** for the final withdrawal. Exclude **Skipped** to see the steps used. Keep rows in calculation order. Columns **BD onward** link to the inputs.

<a id="funding_paths-and-the-trial-schedules"></a>

## Market trial calculations

| Tab | Use it to |
| --- | --- |
| Funding_Paths | Follow annual income, withdrawals, balances and unpaid costs in the low, median and high paths. It follows Summary's dollar view. |
| Funding_Path_Data | Match a saved trial number and year to its source values in that year's dollars. |
| Risk_Trials | Select a trial in **B3**. Row 7 shows its result; saved results start in row 8. Enter **0** to return to the base plan. |
| Market_Paths | Read the selected trial's returns, inflation and spending cuts. |
| Random_Numbers | Inspect the seed and repeatable random-number calculation. |
| Trial_Inputs | Review the selected trial's inputs and stress-test changes. |
| Trial_Cash_Flow | Trace that trial's income, withdrawals, tax and unpaid costs. |
| Trial_Projection | Follow its account activity and balances. |
| Trial_Roth_Access | Check available Roth withdrawals by owner. |
| Trial_Healthcare | Check coverage, credits, HSA use and added medical reserves. |
| Trial_Checks | Compare closing balances with opening balances, transactions and returns. Columns **N–Y** should be zero in active years. |
| Trial_Funding_Calc | Inspect the trial's annual withdrawal and tax calculation. |

If **Trial_Checks J** says **Review**, inspect the differences in that year and **Trial_Cash_Flow AM**. On **Market_Paths**, AH is the allowed cut, AI is the balance threshold and AE is the actual cut. [Runner checks and saved data](macro-runner.md)
