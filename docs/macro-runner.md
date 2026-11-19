# Market-risk runner

The **Run market risk** button on **Summary** refreshes the market results in desktop Excel. The visible worksheets calculate cash flow, taxes, healthcare, withdrawals and account balances. The macro selects trials, checks the results and updates the report. [Market risk guide](risk-analysis.md)

## Run and inspect

1. Open a private copy of `retirement-model.xlsm` in desktop Excel with macros permitted. Keep calculation on **Automatic** while editing. Excel for the web can display the saved results. [Excel for the web and VBA](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)
2. Check the base plan on **Summary**. On **Risk_Settings**, set **Paired market paths** in B29. Use 100 pairs while comparing edits and 1,000 for the final report. Each run adds 33 stress cases. Keep the seed in B30 fixed when comparing plans.
3. Choose **Run market risk** on the right side of **Summary**. Excel shows progress in its status bar. To request a stop, choose **Cancel** beside the run button and wait for the workbook to confirm cancellation.
4. Read or print the report when **Summary** says **Risk results are current**. If a run fails, read the message on **Risk_Settings** in A44. An interrupted run keeps the previous completed figures in the detail sheets and withholds the success rate until a new run finishes.

To inspect one trial, enter its number in **Risk_Trials** B3. Row 7 shows the selected result. Saved results start in row 8. Enter 0 to return to the main plan. On **Market_Paths**, AH shows the spending cut the flexible rule would allow, AI shows its balance threshold and AE shows the cut actually applied.

## How Excel verifies a run

Each random path has a fixed-spending result and a flexible-spending result. The two use the same returns and inflation. Excel can reuse the fixed result when it verifies that the cut is zero and the opening balance meets the threshold in every year. Otherwise it calculates the flexible result separately.

For each trial it calculates, Excel checks 12 closing values against the year's opening balances, transactions and returns. On **Trial_Checks**, N through Y should each show zero in active years. The saved nominal and inflation-adjusted balances must agree with the annual schedule. If a closing check fails, Excel recalculates the trial in year order. If the figures still differ, Excel stops the run before it updates the report. Use the trial number in the message to inspect that year. [Workbook schedules](workbook-schedules.md)

Excel ranks the fixed random paths by ending portfolio in today's dollars. It breaks ties by first shortfall age, largest unpaid balance and trial number. It selects complete funding histories at the 10th, 50th and 90th percentiles. At one pair, all three examples use the same path.

The runner checks that the assumptions still match before it replaces the saved results. It keeps backups until the workbook accepts the complete batch. It restores the inspection selector, sheet protection and Excel's calculation, event, display and cancellation settings when the run ends. [VBA source](../tools/vba/RunMarketRisk.bas)

## Saved results and maintainer checks

The six named ranges below belong to **Risk_Trials**. The funding ranges point to **Funding_Path_Data**.

| Data | Current calculation | Saved values | Saved IDs |
| --- | --- | --- | --- |
| Trial outcomes and annual balances | `RiskTrialCurrent` | `RiskTrialResults` | `RiskTrialIDs` |
| Three ranked funding histories | `RiskPathCurrent` | `RiskPathResults` | `RiskPathTrials` |

On **Risk_Data**, Z1 gives the report status. Z6 through Z11 hold the run state, completion time, assumptions, message, pair count and seed. Each saved trial includes up to 126 annual balances in each dollar view. Each of the three funding examples stores 38 values per year, including the opening year.

From the repository root, use these commands to compare representative saved trials with recalculated worksheets and check how the report handles incomplete results. They require the project dependencies and LibreOffice. Desktop Excel checks cover the macro itself, including refresh, cancellation and recovery. [Test instructions](../tests/README.md)

```sh
python3 tests/native_capture_acceptance.py workbook/retirement-model.xlsm --codes 1,1001,2001,2101,2205
python3 tests/native_integrity_acceptance.py
```
