# Funding corrections and smaller calculation tables

November 17, 2026

## Financial behavior

Eligible Roth IRA funds can cover required costs while the owner is working. Owner-specific basis, conversion seasoning and account age still control access. Confirmed qualified workplace Roth distributions use the plan's entered access age. An unknown workplace distribution age displays **Access unconfirmed** and excludes that money from available withdrawals. The whole-year model uses age 60 for the age-59½ boundary. [IRS Roth IRA ordering rules](https://www.irs.gov/publications/p590b) · [IRS plan distribution rules](https://www.irs.gov/retirement-plans/plan-participant-employee/when-can-a-retirement-plan-distribute-benefits)

Unpaid required spending carries into the next year. Payments cover taxes, earlier unpaid spending and current required costs in that order. New elective savings use only the remaining cash. Columns AR–AW on Cash_Flow show opening bills, repayments, current payments, newly unpaid bills, cumulative missed payments and the reconciliation. Outstanding tax and spending reduce after-tax net resources. Financing interest is zero.

Any year with at least one cent unpaid makes the path unsuccessful, including a bill repaid later. The cumulative missed-payment figure records bills that were late in their due year; the closing obligation records what remains owed.

Planning Metrics now labels the spending-divided-by-withdrawal-rate measure **Spending benchmark before income**. Summary explains income timing and follows the full annual funding calculation. The benchmark does not determine whether the plan pays its bills.

## Calculation and portability

Funding_Calc and Targeting_Calc link source values in visible columns BD onward. Trial formulas can then share their local references in Excel for Mac. The full age-125 capacity and 51 selected-year conversion comparisons remain available. Clear the trial filter to inspect every linked source row.

The measured candidate decreased from 24.0 MB to 6.1 MB. Expanded package contents decreased from 380 MB to 75 MB. Formula cells increased slightly because the new source links are explicit; repeated formula text is shared. The annual withdrawal algorithm and its credit-cliff tests are preserved.

Dashboard chart names now belong to Report_Data. They remain valid when a user renames the workbook. A renamed copy opened, recalculated, saved and printed the chart in Excel for Mac without an external-link warning.

## Verification

- All 85 financial scenarios passed with annual ledger reconciliation. Five new cases cover working-year IRA access, partial access, carried bills, later repayment and confirmed workplace access.
- The SEPP case covers five years of accumulating bills and repayment from a qualified withdrawal at age 60. It passed on the full calculation sheets.
- Full-size cases also passed for age 125, working-year Roth funding, missing retirement age and the earliest fundable Marketplace credit interval.
- The public build passed its release check and independent no-shock comparison across 60 years and 5,760 values. It refreshed the fictional example's 1,000 paired market paths.
- Excel for Mac recalculated the compact engine and extended the example from end age 95 to 125 without rebuilding. The native saved copy passed the independent annual reconciliation for all 90 projected years.

Mac UI observations were under one second for the full-calculation command and about 1.6 seconds for the end-age edit. These observations include UI automation overhead and are preliminary measurements. Windows performance and print verification remain required.

The risk engine and saved-result schema advance to versions 1.4.0 and v4. Use the workbook and development risk runner from this release together. Excel-only risk refresh remains a separate open phase of the improvement plan. [Plan and remaining checks](../model-improvement-plan.md)
