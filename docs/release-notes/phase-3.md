# Phase 3 release record

## Published workbook

On November 15, 2026, the workbook’s quick start changed to four steps on one screen. The input notes use shorter instructions, and the model guide explains healthcare setup and the detailed withdrawal and tax rules. The survivor note now matches the share of combined Social Security entered in row 55.

The rebuilt workbook opened in desktop Excel. All 750,314 formulas and their saved results matched the previous workbook. The survivor benefit and mixed Medicare and Marketplace test cases passed, along with all required pull request checks. [Published workbook update](https://github.com/Jon-275/financial-model/pull/6)

GitHub requires Workbook checks, Greptile Review and Gemini Review on PRs into `main`. See the [test and release guide](../../tests/README.md) for the current checks and the [annual tax data review](../annual-tax-review.md) for maintenance steps.

## Phase 3 launch

Phase 3 first ran locally on November 6, 2026.

- The risk engine reads the workbook’s active scenario and simulates correlated asset returns with separate working and retirement allocations. Each path recalculates annual cash, taxes, healthcare costs and account access. A shortfall counts in the year it occurs, even if the portfolio later recovers.
- Fixed and flexible spending use the same return draws. The flexible policy can reduce only the eligible part of the recurring retirement budget. The report also tests the workbook’s early retirement stock-return scenarios.
- The workbook build now clears all saved formula results before LibreOffice recalculates. The previous build could preserve a cached formula after one of its inputs had been rebuilt. Recalculation corrected the first-year bracket headroom. Annual taxes, funded withdrawals and account balances matched the previous calculation. The workbook keeps its scenario inputs, formulas and charts.

## Verification

The full workbook build passed from the age-125 template. The repackaged workbook passed the existing independent workbook checks. With no market shocks, the risk engine matched all 6,570 cash, tax, healthcare and account values across 73 active years to within one cent. All 28 risk tests and 66 LibreOffice workbook fixtures passed. The engine reconciled 22,950 values across 62 valid fixture workbooks and rejected the four invalid cases. The tests cover repeated seeds, zero volatility, invalid workbook caches, access restrictions and funding around the Marketplace credit thresholds. Synthetic workbooks cover a 103-year projection ending at age 125, the spouse’s required distributions, Roth access, Medicare, HSA reimbursement, and a cash shortfall despite paying taxes in full.

The active case was also run with 1,000 paired market paths using seed `2026`. Both policies use the same workbook assumptions. The risk report is a separate snapshot with its source checksum and calculation engine recorded. The workbook’s built-in illustrations continue to use the base cash-flow schedule. LibreOffice was used for the initial workbook calculation.

[Risk analysis instructions](../risk-analysis.md)

## Microsoft Excel verification

On November 7, 2026, Excel for Mac 16.114.4 exposed two compatibility issues. Excel removed shared formulas that referred to other worksheets in the funding and conversion sheets. It also returned errors when the tax formula selected a bracket table inside an array calculation. The build now writes those cross-sheet formulas individually and selects the tax table outside the array calculation. The bracket rates, scenario inputs and withdrawal policy were preserved. Writing the formulas individually increased the workbook to about 24 MB.

The rebuilt workbook opened without repair. After a full calculation from blank caches, the active case passed the workbook checks and matched all 6,570 annual cash, tax, healthcare and account values to within a cent. Changing the end age to 125 on the Excel test copy produced 103 annual rows, reconciled 9,270 values and extended the portfolio chart to the new end year. The full build and seven targeted LibreOffice scenarios also passed. The companion risk report opened in Excel and its chart rendered.

## Wage growth and California disability insurance

On November 8, 2026, the wage growth input changed to a total annual increase that includes inflation. California disability insurance uses the published 1.4% rate in 2027 and carries it forward as a planning assumption. [California EDD rate schedule](https://edd.ca.gov/en/Payroll_Taxes/Rates_and_Withholding/)

The rebuilt workbook passed a full calculation in Excel for Mac from blank caches and reconciled 6,570 values across 73 years. All 28 risk tests passed, along with the wage growth, payroll and age-125 workbook cases. The risk report was regenerated from the updated workbook.

## Workbook instructions

The earlier quick start used three steps from choosing a scenario to reading the result. The Inputs notes explained common entries, and Summary pointed to Cash_Flow when the plan fell short. Rate notes pointed to Tax_Tables.

The README, Inputs and Summary layouts were checked in Excel. All 883,201 formula cells and the scenario assumptions matched the prior release. The active projection reconciled across 73 years, and the refreshed risk report reproduced the prior results.

## Workbook layout and reports

The workbook opens on **Start Here**. **Inputs** keeps the scenario selector and headings visible while scrolling, with optional details grouped under the plus signs. Each comparison scenario changes one input. One delays retirement by five years. The other cuts retirement spending by 10%.

**Summary** shows whether modeled spending and tax are paid, the first year with unpaid costs, and the ending portfolio in today’s dollars. It shows the savings target separately. Two charts show the portfolio by age and how income, withdrawals and health savings cover bills around retirement. **Risk_Settings** replaces the old market simulation tabs and keeps their report inputs. The regenerated companion risk report reproduces the prior simulation results for the same assumptions and seed.

The rebuilt workbook opened in Excel for Mac and passed the accounting checks after recalculation from blank caches. The 6,570 annual values across 73 years reconciled to within one cent. Switching to the later retirement scenario moved the retirement marker and changed the ages shown on the bills chart. Changing the end age to 125 extended the portfolio chart to age 125. The age-125 and late-shortfall fixtures also passed in LibreOffice.

All 66 workbook regression fixtures and 29 risk tests passed. Both saved comparison scenarios and the fixed-distribution shortfall fixture also passed.

## Reconciliation warning and one-cent shortfalls

Summary shows **CHECK PLAN** and directs readers to **Checks** when an active year’s reconciliation has not passed. Both sheets use the same reconciliation result from the cash-flow calculation. The acceptance tests use the workbook’s inclusive one-cent threshold for unpaid costs.

Excel for Mac recalculated the release from blank caches. A test difference switched the Summary headline from **YES** to **CHECK PLAN**. Separate cases checked the opening and final plan rows, a missing reconciliation result, and the amounts immediately below and at one cent. The base projection reconciled all 6,570 annual values.

## Workbook evaluation follow-up

The model uses an editable opening year and keeps published tax figures tied to their tax years. The forecast applies investment fees to account growth and can use a separate inflation rate for healthcare. Household healthcare changes for each person at 65. The survivor’s Social Security amount now follows the combined benefit and the share entered on Inputs.

Planned direct Roth IRA contributions use income limits for joint and single filers, including the published 2027 amounts. The model directs remaining savings to the taxable account when the Roth limit reduces the planned contribution. [IRS Notice 2026-66](https://www.irs.gov/pub/irs-drop/n-26-66.pdf)

The revised workbook restores the generated tax-table layout so the companion risk report can run. The report reconciled all 6,570 annual values and completed 1,000 paired simulations. The repository workbook uses fictional assumptions and exposes all 15 worksheet tabs. The build removes inherited document properties, sensitivity labels and revision identifiers.

## Printable retirement plan

**Start Here** links to setup and results. **Inputs** highlights the selected scenario and links to groups of assumptions. **Summary** reports whether the plan pays its bills, identifies the first shortfall when one occurs, and explains what to do if an input is missing.

The print package includes two pages of input decisions and three pages covering the outlook, retirement cash flow, and next decisions. **Annual Detail** prints the full selected timeline as a separate appendix. The funding result and the savings benchmark appear in different parts of the report.

The workbook generator writes the layout, report links, chart and print settings. **Appendix** marks where the supporting sheets begin and provides an index to the calculations and advanced tools. The published workbook uses fictional figures and all 17 sheets are visible.

Excel for Mac 16.115 opened the workbook and recalculated a scenario through age 125. The combined Inputs and Summary report exported as five pages on US Letter and A4. Annual Detail printed through the selected ending age with repeating headers. All 80 workbook scenarios and 29 risk tests passed during the redesign.
