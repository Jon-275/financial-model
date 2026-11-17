# Market-risk runner

The desktop workbook uses **Run market risk** to calculate the selected market paths and required stresses. The macro coordinates visible worksheet formulas. It does not contain tax, healthcare, withdrawal or account-access rules.

The source-built project ran in desktop Excel for Mac. Its three-pair test included all 33 stresses and matched the independent example across 25,407 fields. The release checklist tracks interruption, larger runs and Windows testing. [Improvement plan](model-improvement-plan.md)

## Run and inspect

1. Finish Inputs, Asset_Allocation and Risk_Settings.
2. Select **Run market risk** on Summary. Progress appears in Excel's status bar.
3. To stop early, choose **Cancel** beside the run button or press Esc. Excel keeps the previous completed batch. Run the scenarios again to update the report.
4. Read the refreshed Summary. The workbook records the completed run's date, assumptions, seed and number of paths.
5. Inspect Risk_Trials for individual outcomes and Funding_Paths for the complete low, median and high paths. Their calculations remain visible behind Appendix.

A run uses the workbook's own formulas and needs no external service or Python process. VBA requires desktop Excel and permitted macros. The runner does not change macro security settings or use opening events. [Microsoft VBA compatibility guidance](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)

## Calculation and publication

The public procedures are `RunMarketRisk` and `CancelMarketRisk` in [RunMarketRisk.bas](../tools/vba/RunMarketRisk.bas). The first prevents a second overlapping run; the second requests cancellation at the next progress update. Esc uses Excel's normal error-18 interruption mechanism. [Microsoft cancellation guidance](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablecancelkey)

A run calculates only the requested fixed and paired flexible paths, followed by 26 early-loss tests and seven inflation, longevity and medical tests. All outputs are collected in memory. It ranks fixed paths by ending real balance, then earlier failure age, larger unpaid balance and trial identifier. The low, median and high examples use the paths nearest the 10th, 50th and 90th percentiles of ending balances. A one-path run uses that path for all three examples.

The complete annual schedules for those three paths are captured before publication. The runner then replaces the primary results, their identifiers, the three full paths and the completed-run record. It keeps backups until the workbook confirms that the results are current. A publication error restores those backups and records failure. Ordinary cancellation before publication leaves them untouched.

Assumption records must match before each capture and immediately before publication. Every captured value is checked for Excel errors. A path with failed calculation checks stops the run. An interrupted or failed run remains visibly unavailable to the report until a successful replacement run.

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

Unused result rows are cleared by the replacement batch. Current and result ranges must have matching widths, and result areas cannot contain formulas or overlap one another. The schema and engine version must match the runner.

Risk_Data stores the run state in Z6, a numeric Excel timestamp in Z7, the completed assumption record in Z8, the latest message in Z9, path count in Z10 and seed in Z11. The timestamp handles both Excel date systems. Current report status also depends on the workbook's formula checks; setting a status value alone cannot validate a batch.

## Excel settings and protection

The runner saves and restores calculation mode, events, screen updating, cancellation behavior, cursor and status-bar settings. It restores the original inspection selector. Excel's calculation method may also recalculate other open workbooks, as documented by Microsoft. [Calculation guidance](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculate)

For worksheets protected without a password, the runner temporarily allows macro writes while retaining user protection. It then restores the previous protection options. A custom password is not bypassed. No access to the VBA project object model is required. [Worksheet protection guidance](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protect)

## Verification

`tests/test_macro_runner_contract.py` checks the live compiler-to-runner range contract, field order, release identifiers, publication ordering, cleanup guards and prohibited automation interfaces. These are source checks, not execution evidence.

For native algorithm checks, append `tools/vba/tests/RunnerChecks.fragment` to a disposable copy of the module, import it into a test workbook and run `VerifyMarketRiskRunner`. The checks use independent expected trial coverage and a hand-specified path order. Keep this test-only fragment out of the release module.

Native release tests must also cover a successful complete run, changed inputs during a run, Esc and button cancellation, a rejected trial, publication failure, a repeat run, both calculation modes and worksheet protection. Confirm prior results and global Excel settings after each failure. Compare seeded financial results with the independent model, then measure 10, 100 and 1,000 paired paths at the example and maximum horizons on Mac and Windows.

## Prepare the example for release

The public builder uses `tools/native_bootstrap.py` to populate the synthetic example. It reads the path count, seed, return assumptions and inflation settings from the workbook. It independently calculates the annual ledger, checks it against saved worksheet values and writes all selected paths and 33 stresses. The age-100 and age-105 stresses use a separate, extended calculation copy so the user's selected end age stays in effect for the main report.

The build recalculates the report and checks its completion status, number of captured rows and the three selected funding-path IDs. It preserves the original Excel package and transfers calculated values and chart data from the verification copy. To compare a built example with individual worksheet trials, run `python tests/native_capture_acceptance.py example.xlsx --codes 1,1001,2001,2101,2205`. Use `--codes all` to verify each captured trial. Excel users refresh their own plans with **Run market risk** on Summary.

Completion checks also verify that each row has its expected trial number. The three stored funding paths must match the workbook's independently calculated low, median and high ranks. If a trial identifier or ranked path is changed manually, the Summary marks the batch incomplete. To test the integrity checks, run `python tests/native_integrity_acceptance.py`.

During a refresh, Excel completes the annual calculation before the runner reads a trial. The runner compares each year's balance in both dollar views with the trial ledger. A flexible path with no spending cuts must match its fixed counterpart. A changed path that repeats the previous positive ending balance is recalculated and checked again. Each of the three ranked funding paths must also match the selected fixed path’s annual closing balances. If the values disagree, the runner recalculates once. A second mismatch stops the run before it replaces the saved results. To check the saved no-cut pairs, run `python tests/native_capture_acceptance.py workbook/retirement-model.xlsm --codes pairs`.
