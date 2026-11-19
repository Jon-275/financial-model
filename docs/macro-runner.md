# Market-risk runner

In Excel for Mac, **Run market risk** calculates the selected market paths and required stresses. The macro coordinates the run. The visible worksheets calculate taxes, healthcare costs, withdrawals and account access.

With 1,000 pairs and 33 stress cases, the worksheet refresh took 7 min 52 sec through age 95 and 33 min 10 sec through age 125 on the test Mac. [Desktop timing and verification](native-risk-implementation.md#full-report-comparison)

## Run and inspect

1. Enter the plan on **Inputs**. Review the portfolio on **Asset_Allocation** and the market assumptions on **Risk_Settings**. Use 100 paired paths to compare edits and 1,000 for the final report.
2. Select **Run market risk** on Summary. Progress appears in Excel's status bar.
3. To request a stop, choose **Cancel** beside the run button. Check that the status changes to cancelled. Excel keeps the previous completed batch. Run the scenarios again to update the report.
4. Read the refreshed Summary. The workbook records the completed run's date, assumptions, seed and number of paths.
5. Inspect Risk_Trials for individual outcomes and Funding_Paths for the complete low, median and high paths. Their calculations remain visible behind Appendix.

A run uses the workbook's own formulas and needs no external service or Python process. To refresh the scenarios, open the workbook in Excel for Mac with macros permitted. The runner does not change macro security settings or use opening events. [Microsoft VBA compatibility guidance](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)

## Calculation and publication

The runner prevents a second overlapping run. The cancellation procedure requests a stop at the next checkpoint. Excel can also deliver a keyboard interruption as error 18. The recovery test verified the procedure. The Mac keyboard test timed out, and the button still needs a verified click during calculation. [Runner source](../tools/vba/RunMarketRisk.bas) · [Microsoft cancellation guidance](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablecancelkey)

A run records both spending policies for every requested market path. It also records 26 early-loss tests and seven inflation, longevity and medical tests. The runner collects the results in memory before updating the report.

Each fixed and flexible pair uses the same market returns and inflation. Spending can differ only in years when the flexible policy cuts living costs. While calculating the fixed path, the workbook checks each year's possible cut and the balance at which the cut would begin. If every possible cut is zero and each opening balance meets its threshold, the runner can copy the verified amounts to the flexible path's own result row. A balance exactly at the threshold passes. A missing, invalid or inconclusive check sends the flexible path through the full worksheet calculation.

To inspect a pair, enter its fixed trial number on **Risk_Trials**. On **Market_Paths**, column AH shows the possible flexible cut and AI shows its opening-balance threshold. Select the flexible trial to see the spending actually used in column AE. **Risk_Settings** shows how many flexible results reused their fixed partner in the last Excel run. Both policies count toward their respective funding rates.

The runner ranks fixed paths by ending balance in today's dollars. Ties are resolved by earlier failure age, larger unpaid balance and trial number. The low, middle and high funding examples are the paths nearest the 10th, 50th and 90th percentiles. A one-path run uses that path for all three examples.

The complete annual schedules for those three paths are captured before publication. The runner then replaces the primary results, their identifiers, the three full paths and the completed-run record. It keeps backups until the workbook confirms that the results are current. A publication error restores those backups and records failure. Ordinary cancellation before publication leaves them untouched.

Assumption records must match before each capture and immediately before publication. Every captured value is checked for Excel errors. A path with failed calculation checks stops the run. An interrupted or failed run remains visibly unavailable to the report until a successful replacement run.

### When an account balance fails a check

Excel recalculates each trial it evaluates and checks 12 annual account balances against the opening accounts, transactions and returns. The saved result must agree with the annual ledger in both dollar views. The three ranked funding examples must also agree with their selected market paths.

If a closing balance differs, a required check is missing or a changed path repeats the preceding positive ending balance, the runner recalculates that trial in year order. If the discrepancy persists, it stops before publishing new results. The message identifies the trial and the first balance or check that failed. Enter the trial number in **Risk_Trials B3** and look at columns N through Y on **Trial_Checks**. The appendix retains the worksheet formulas for inspection. [Checks and tests for maintainers](../tests/README.md)

## Workbook interface

All six names are local to Risk_Trials. Their ranges belong to the same workbook; the three full-path areas may be on Funding_Path_Data.

| Name | Responsibility |
| --- | --- |
| RiskTrialCurrent | One row of outcomes and annual balances for the selected trial |
| RiskTrialResults | Capacity for 1,000 fixed paths, 1,000 paired flexible paths and 33 stresses |
| RiskTrialIDs | Identifiers for the contiguous populated result rows |
| RiskPathCurrent | One row containing all years of the selected funding path |
| RiskPathResults | Three complete captured funding paths |
| RiskPathTrials | Their fixed-path identifiers |

The runner checks the full width of each range before it captures a trial. The primary row holds two sets of 126 annual balances. The funding row holds 38 fields for each of those 126 years. Result areas must match the corresponding row width and contain no formulas or overlapping ranges. The replacement batch clears unused result rows. The schema and engine version must match the runner.

Risk_Data stores the run state in Z6, a numeric Excel timestamp in Z7, the completed assumption record in Z8, the latest message in Z9, path count in Z10 and seed in Z11. The timestamp handles both Excel date systems. Current report status also depends on the workbook's formula checks; setting a status value alone cannot validate a batch.

## Excel settings and protection

The runner saves and restores calculation mode, events, screen updating, cancellation behavior, cursor and status-bar settings. It restores the original inspection selector. Excel's calculation method may also recalculate other open workbooks, as documented by Microsoft. [Calculation guidance](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculate)

For worksheets protected without a password, the runner temporarily allows macro writes while retaining user protection. It then restores the previous protection options. A custom password is not bypassed. No access to the VBA project object model is required. [Worksheet protection guidance](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protect)

## Verification

`tests/test_macro_runner_contract.py` checks the live compiler-to-runner range contract, field order, release identifiers, publication ordering, cleanup guards and prohibited automation interfaces. These are source checks, not execution evidence.

For native algorithm checks, append `tools/vba/tests/RunnerChecks.fragment` to a disposable copy of the module, import it into a test workbook and run `VerifyMarketRiskRunner`. The checks cover trial numbers, path order, flexible spending, cancellation and numeric controls. Excel for Mac confirmed that a numeric zero is a valid seed, and rejected a blank cell and a number stored as text. The earlier reader treated a blank as zero. Keep this test-only fragment out of the release module.

Native release tests must also cover a successful complete run, changed inputs during a run, Esc and button cancellation, a rejected trial, publication failure, a repeat run, both calculation modes and worksheet protection. Confirm prior results and global Excel settings after each failure. Compare seeded financial results with the independent model, then measure 10, 100 and 1,000 paired paths at the example and maximum horizons in Excel for Mac.

## Prepare the example for release

The public builder uses `tools/native_bootstrap.py` to populate the synthetic example. It reads the path count, seed, return assumptions and inflation settings from the workbook. It independently calculates the annual ledger, checks it against saved worksheet values and writes all selected paths and 33 stresses. The age-100 and age-105 stresses use a separate, extended calculation copy so the user's selected end age stays in effect for the main report.

The build recalculates the report and checks its completion status, number of captured rows and the three selected funding-path IDs. It preserves the original Excel package and transfers calculated values and chart data from the verification copy. To compare a built example with individual worksheet trials, run `python tests/native_capture_acceptance.py example.xlsx --codes 1,1001,2001,2101,2205`. Use `--codes all` to verify each captured trial. Excel users refresh their own plans with **Run market risk** on Summary.

Completion checks also verify that each row has its expected trial number. The three stored funding paths must match the workbook's independently calculated low, median and high ranks. If a trial identifier or ranked path is changed manually, the Summary marks the batch incomplete. To test the integrity checks, run `python tests/native_integrity_acceptance.py`.
