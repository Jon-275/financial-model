# Workbook guide

Follow the tabs from left to right. Use **Inputs** to enter figures and **Summary** to read the plan. The tabs between **Summary** and **Appendix** let you review assumptions and compare choices. **Appendix** links to the calculation sheets.

[Try the example first](getting-started.md) · [Detailed input reference](input-reference.md) · [All documentation](README.md)

<a id="run-your-first-plan"></a>

## Start Here

Save a private copy of the workbook in `.xlsm` format. Open it in desktop Excel for Mac and keep calculation set to **Automatic**. Choose **Open Inputs**, or select the **Inputs** tab directly.

Macros are needed for **Run market risk**. Follow your organization's macro policy. Excel for the web can show saved results but cannot run the analysis. Release v1.0.0 was verified on desktop Excel for Mac 16.115. Windows is unverified.

<a id="enter-income-spending-and-savings"></a>
<a id="enter-retirement-account-details"></a>
<a id="enter-healthcare-costs"></a>

## Inputs

Choose a scenario in **B4** and edit its shaded cells. Column D holds the example, E compares later retirement and F compares spending 10% less in retirement. Column B shows the active values; some are calculated.

Use **Jump to** to move through the input groups. Enter ages, income, savings and spending first, then review account access, healthcare and tax inputs. Use today's dollars unless the cell says otherwise. The opening year is in **B130**; the forecast starts the following year.

The [input reference](input-reference.md) follows the rows on this sheet. It explains which amounts to include, blank versus zero, and special values such as **999 = off**.

Blue cells are editable. The highlighted scenario column supplies the active values used by the plan.

Resolve input-check messages before reading the final results or running market risk. **View input decisions** opens the report's record of account-access assumptions.

<a id="read-the-results"></a>

## Summary

Read the report in page order:

| Page | What to review |
| --- | --- |
| 1. Retirement outlook | Does the base plan cover spending and taxes through the end age? If not, when is the first shortfall? Then review the market success rate and balance range. |
| 2. Retirement cash flow | Which income and savings pay the bills around retirement? Are those accounts available when needed? |
| 3. Risk and decisions | Compare low, median and high market paths, spending cuts and early stock losses. Note the next change you want to test. |
| 4. Inflation and longevity | Review the separate inflation, longer-retirement and medical-budget tests. |
| 5. Input decisions | Check the ages, income, savings and account-access assumptions used in the report. |
| 6. Spending decisions | Check the spending, healthcare and account choices recorded with the plan. |

A **shortfall** means at least one cent of spending or tax was unpaid when due. Later repayment does not remove that shortfall. A positive ending balance does not prove that every year's bills were paid; some savings may have been unavailable. Summary lists workplace Roth savings that remain unavailable at the end.

Market results are separate from the base projection. Read them only when the status says **Risk results are current**. See [what the market results mean](risk-analysis.md#read-the-summary).

### Print your plan

Choose a **Dollar view** on Summary:

- **Today's dollars:** amounts at opening-year prices, for comparison with your current budget.
- **Future dollars:** amounts at each projected year's prices.

Check the scenario, opening year, dollar view and risk status. Select only **Summary**, then choose **File → Print → Active Sheets** to print or save all six pages as a PDF. Print **Annual Detail** separately for the full timeline.

To compare plans, save a report before changing an input. Run market risk again before saving the next report. Each completed run replaces the previous one.

## Planning Metrics

Use this sheet for more detail on the result: savings benchmarks, balances available at retirement, funding, taxes, healthcare and scenario comparisons.

**Spending benchmark before income** divides living and healthcare costs by the withdrawal rate entered on Inputs. Use the funding result to check whether the plan can pay each year's bills; it includes income timing, taxes and account access.

## Asset_Allocation

Review the investment profiles chosen in **Inputs rows 15–16**. Edit the blue cells for each account's mix of stocks, bonds, cash and real assets. Each account column must total **100%**.

Use **On this sheet** to jump between returns, account mixes, profile checks and the returns used by the plan. The same index on the other working tabs links to their main sections and back to Summary.

Check the expected returns, annual fee, volatility and correlations. Volatility sets how much returns vary; correlations set how closely assets move together. The table at the bottom shows expected returns after fees and before inflation.

<a id="compare-roth-conversions-and-market-risk"></a>

## Income_Targeting

Use this sheet to compare a Roth conversion in one retirement year.

1. Choose the retirement age to compare. Enter **0** for the first retirement year.
2. Compare a zero conversion with the **Conversion to test**, entered in that year's dollars.
3. Review funded costs, withdrawals, taxes and healthcare together. Choose **Open conversion detail** to see the range of tested amounts.
4. Enter a planned conversion on **Inputs** to see its effect over the full plan, including later Medicare costs.

The display limit controls the range shown in Conversion Detail and the charts below the comparison. Choose **View conversion charts** to see them.

## Risk_Settings

Review the spending-cut rule, number of paired paths, seed and inflation settings. Use **100 pairs** while comparing changes or **1,000 pairs** for a final report. Keep the seed unchanged when comparing plans.

Finish reviewing the assumptions, including **Tax_Tables**, then close other workbooks and return to **Summary**. Choose **Run market risk**. Progress appears in Excel's status bar. Wait for **Risk results are current** before printing. Use **Cancel** to stop a run and wait for confirmation.

The runner recalculates all open workbooks. Keeping only your plan open avoids extra work during each trial.

Runtime depends on your Mac and plan. [Measured runtimes](risk-analysis.md#measured-runtime)

## Tax_Tables

Review the years and values used for federal, California, payroll, retirement-account and healthcare calculations. Use **On this sheet** to reach a table.

Future values may be estimates. Check the year and source before changing a value. Maintainers should follow the [annual tax review](annual-tax-review.md).

## Appendix

Use this divider to open the supporting sheets in workbook order. Start with **Annual Detail** for the yearly plan or **Risk Results** for market outcomes. Use **Cash_Flow** to investigate a shortfall and **Checks** to investigate a calculation message.

The [calculation-sheet guide](workbook-schedules.md) explains every tab after Appendix and links to the calculation rules, including [how the Monte Carlo simulation works](model-methods.md#how-the-monte-carlo-simulation-works). Select a formula cell to read its calculation. To change a formula in your own copy, choose **Review → Unprotect Sheet**; no password is needed.

<a id="how-the-model-works"></a>
<a id="withdrawals-and-taxes"></a>
<a id="when-roth-savings-become-available"></a>
<a id="how-income-affects-social-security-taxes"></a>
<a id="tax-rates-and-healthcare-credits"></a>
<a id="investment-mix-and-account-values"></a>

## Calculation rules

See [how the model calculates the plan](model-methods.md) for withdrawal order, Roth access, taxes, account values and market assumptions.

## Items to estimate separately

The workbook leaves the following items for you to estimate separately.

- **Survivor planning.** Life insurance, pension survivor choices, the joint tax return for the year of death, changes to cost basis at death and the transfer of a health savings account after its owner dies.
- **Federal taxes.** Qualified charitable distributions, medical-expense deductions and the temporary enhanced senior deduction.
- **State taxes.** Changes in state law, California tax on HSA earnings and California’s treatment of certain municipal interest.

Also make separate estimates for mixed opening Roth IRA and workplace Roth balances, inherited accounts, backdoor Roth contributions, rollovers, the age-55 separation exception and early nonqualified workplace Roth withdrawals. [Account details](input-reference.md#cash-and-account-access)

For the calculation rules and tax assumptions, see [how the model works](model-methods.md). When sharing the workbook, use a fresh copy with fictional figures and include the [license](../LICENSE). Your working file contains all three scenarios and saved market results.
