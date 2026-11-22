# Market-risk runner

The **Run market risk** button on **Summary** refreshes the market results in desktop Excel. The visible worksheets calculate cash flow, taxes, healthcare, withdrawals and account balances. The macro selects trials, checks the results and updates the report. [Market risk guide](risk-analysis.md)

## Inspect a trial

For setup, running, cancellation and error messages, follow the [market risk guide](risk-analysis.md#run-the-analysis).

On **Risk_Trials**, enter a trial number in **B3**. Row 7 shows its result; saved results start in row 8. Enter **0** to return to the base plan.

- For a random fixed-spending trial, add **1,000** to select its flexible partner.
- For early-loss trials **2001–2013**, add **100** for the flexible partner.
- The seven inflation, longevity and medical tests use fixed spending.

Both paired records are saved even when their outcomes match. Columns **I–J** show years with cuts and the total cut in today's dollars. On **Market_Paths**, AH shows the allowed cut, AI the balance threshold and AE the actual cut.

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
