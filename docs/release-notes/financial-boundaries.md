# Medicare surcharge, withdrawals and the first shortfall

November 17, 2026 · Risk engine 1.4.1 · Workbook schema v4

## What changed

The household healthcare budget now handles years when investment losses bring income below zero. The old Medicare surcharge lookup started at zero and could return a spreadsheet error. The workbook now assigns zero surcharge below the income brackets and keeps the actual income for tax calculations. [Social Security surcharge rule](https://www.ssa.gov/OP_Home/cfr20/418/418-1101.htm)

A withdrawal that uses an account in full now takes its exact available balance. Previously, subtracting other accounts from the total withdrawal could leave a fraction of a cent. That remainder could keep the cost basis of a fully sold brokerage account in later years. The workbook and the risk calculation now use the same withdrawal limits for brokerage, traditional and Roth accounts. Mandatory distributions count toward traditional withdrawals.

The first page of **Summary** now identifies the earliest year with at least one cent of unpaid costs. In Excel for Mac, an array lookup could fail and the report could then say the plan covered all costs. The report now looks up the year directly from the annual shortfalls whenever the funding calculation finds unpaid costs.

The workbook was rebuilt from the public fictional example. Use the workbook and risk runner from the same release. The on-demand Excel risk runner is tracked in the [improvement plan](../model-improvement-plan.md).

## Verification

- The two negative-income cases reproduce the error in the prior workbook. The full-liquidation case uses taxable investment income and a cost basis larger than the account balance. The earlier Python allocation reproduces the retained balance.
- The complete fictional example passed its release and privacy checks. Its 1,000 paired market paths refreshed successfully. With market shocks turned off, the workbook and independent calculation agreed on 5,760 annual values.
- All 37 risk tests passed, including both withdrawal orders, mandatory distributions and access to Roth funds. An additional 480,000 synthetic allocations checked account limits and total cash across partial and full withdrawals.
- Four fixture tests verified that editing test inputs preserves Excel's namespace declarations, formulas and other workbook parts. Nine privacy tests and 11 pull-request review tests passed.
- All three new cases passed on full-size workbooks, including the first shortfall year and the headline printed on page one. The risk refresh was also tested through age 125 and with unpaid spending carried forward.

Windows desktop verification remains open.
