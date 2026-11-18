# Workbook risk refresh

November 17, 2026

The development workbook runs market scenarios in desktop Excel and includes the results in its six-page Summary. The release checklist tracks final testing and publication. [Improvement plan](model-improvement-plan.md)

## Using the workbook

After entering or changing assumptions, open **Summary** and choose **Run market risk**. The button sits outside the print area. The workbook shows progress during the run and the completion date afterwards. It keeps the previous results while it calculates. Summary uses the new results once the entire run passes its checks.

Changing a financial or risk assumption makes the saved results out of date. Switching between today's and future dollars changes how amounts appear and preserves the run. Excel for the web can display saved results. Running the scenarios requires a desktop version of Excel that permits VBA. [Microsoft's Excel for the web guidance](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)

## How the calculation works

1. **Risk_Settings** sets the number of paired paths, the random seed and the inflation assumptions. Each pair shares its returns and inflation so the comparison measures the effect of flexible spending.
2. **Random_Numbers** and **Market_Paths** generate one repeatable sequence at a time. The visible trial sheets use the main plan's annual formulas for income, tax, healthcare, withdrawals, account access and unpaid bills.
3. The macro runs each pair and 33 stress cases. It checks the assumptions and annual reconciliations before accepting each result.
4. The macro ranks the fixed-spending paths and captures the complete funding schedules for three examples. It replaces the previous batch after all calculations pass and restores Excel's settings and the inspection selector.
5. **Summary** shows the resulting funding rate, range of balances and stress outcomes. **Funding_Paths** traces the low, middle and high examples from income and spending through withdrawals and closing balances. The middle 80% of balances at an age comes from the distribution across paths at that age. A full funding example follows one path for its entire horizon.

The workbook builder uses an independent annual model to populate the fictional example and check results. People running the workbook use Excel. [Runner source and controls](macro-runner.md)

## Calculation evidence

| Check | Result |
| --- | --- |
| Repeatable market draws | The visible generator matched 8,500 independent reference values in both Excel for Mac and LibreOffice. |
| Annual financial policy | Eight selected market paths and stresses reconciled 47,040 annual values, including fixed and flexible spending under variable inflation. |
| Medicare surcharge | Excel for Mac recalculated the joint and single negative-income cases. Both had a zero surcharge and retained their funding shortfall. |
| Account liquidation | Exact withdrawals remove the closed holding and its basis. New deposits start their own basis. The combined 5% inflation, lower-return, age-105 stress reconciled 6,720 annual values between the corrected workbook in LibreOffice and the independent model. The largest difference was below one hundred-thousandth of a dollar. |
| Falling prices after Social Security starts | In two three-year examples, nominal benefits held at $950 when first claimed after deflation, and $1,000 when already claimed. The workbook and independent reader agreed on 576 annual financial values. |
| Brokerage sales | Three Excel for Mac cases passed all 20 checks for retained shares, a sale followed by a deposit and investment income when no shares were sold. Five direct worksheet trials also matched 15,803 fields in the full fictional example. |
| Native refresh and printing | Excel for Mac completed three paired paths and all 33 stresses in a development build. The 25,407 compared financial and funding fields agreed with the independent saved example to the cent. Excel printed Summary in six pages. |
| Earlier annual balance and funding checks | An early run carried values from the preceding path into seven fixed paths and changed all three selected examples. The runner now checks each year's captured balances against its ledger and retries a suspicious repeated ending. With those earlier checks, the Mac example matched all 573,306 trial fields, three selected path IDs, 378 annual balances and 14,367 detailed funding fields. |

## What takes time in Excel

A 100-pair run on the earlier worksheet runner took 70.13 seconds, including all 33 stresses. Excel spent 67.22 seconds recalculating worksheets. Reading and writing cells took 0.21 seconds, including publication. Updating the final report took 0.43 seconds.

These tests used the same example inputs and retained the visible annual worksheets as the reference.

| Excel for Mac experiment | Result |
| --- | --- |
| Remove unused withdrawal-search formulas for this plan | The saved test workbook had 173,186 search formulas, down from 560,571. The 100-pair run took 70.47 seconds, including 67.19 seconds of worksheet calculation and one full retry. All outcomes and ranked funding schedules matched. Removing unused formulas gave no measurable gain in this run. |
| Calculate in a temporary workbook containing only the 21 required sheets | Excel completed all 233 outcomes and three ranked funding schedules. Every saved result matched the reference run. This test recorded the completion time only, so its duration is unknown. |
| Calculate market paths and account balances in VBA arrays | Generating all 2,033 market paths took 2.49 seconds for 122,005 modeled years. Replaying 80,500 annual account closings took 0.19 seconds. Thirteen trials matched the worksheet reference for returns, prices and balances. The test used withdrawals and cash flows already calculated by the worksheets. Withdrawals, taxes and healthcare still ran on the worksheets. |
| Evaluate a proposed withdrawal in memory using the worksheet formulas | A test generated a VBA evaluator from the 30 formulas used for one withdrawal. In Excel for Mac, 410 proposed withdrawals matched the live formulas exactly across 12,300 results, including the funding and health-credit decisions. The evaluator processed 2,050 withdrawals in 1.81 seconds. Five funding decisions differed from the archived LibreOffice values where a tiny negative amount had been saved as zero. The experiment covers a single proposed withdrawal. A full annual search and market run still require separate tests. |
| Generate direct VBA for a proposed withdrawal | Excel for Mac compared the generated arithmetic, the interpreter and the live worksheet. All 12,300 amounts and funding decisions matched exactly. In two runs of 2,050 withdrawals each, direct VBA took 0.062 seconds both times. The interpreter took 1.305 and 1.303 seconds. These timings cover the proposed withdrawal only. |
| Compare every decision in an annual search | Across seven fictional years, the Mac test evaluated 231 proposed withdrawals using the existing worksheet search controls. It compared the original formulas with the in-memory withdrawal evaluator. All 32,718 cells, including each proposed draw, the final draw and the funded or unfunded result, matched exactly. A separate portable test ran the full search in memory. Five intermediate funding decisions differed near zero against LibreOffice, although the final funding status and conversion choices agreed. A full in-memory annual search still requires testing in Excel. |

For a quick comparison, set **Paired market paths** on **Risk_Settings** to 100 and run the scenarios. Set it to 1,000 and run again for the final report. Both runs include the 33 stress cases. More paths reduce the sampling variation in the reported funding rate.

### Fixed and flexible spending on the same market path

The runner checks whether flexible spending would change any year's cash flow. It compares the opening portfolio with the spending-cut threshold and reads the possible reduction from **Market_Paths**. When every check passes, it records the verified amounts under both trial numbers. If any year could differ, Excel calculates the flexible path separately. Users can still select either trial and inspect its annual formulas.

Excel for Mac 16.115 on macOS 26.6.2 ran the same fictional plan with the same seed and assumptions. The clock covered the workbook run through publication and restoration of Excel's settings. Timing stopped before Excel saved the test copy.

| Paired market paths | Original incremental calculation | Incremental calculation with the paired-path check | Flexible results reused | Results checked |
| ---: | ---: | ---: | ---: | --- |
| 100 | 71.16 seconds | 58.07 seconds | 46 | 233 outcomes and three detailed funding examples matched across 80,454 fields. |
| 1,000 | 12 min 7 sec | 7 min 55 sec | 515 | 2,033 outcomes and three detailed funding examples matched across 588,054 fields. |

The 100-pair comparison used the same timing procedure for both runs. The earlier 1,000-pair measurement was taken with other tests running. This experiment established that verified pairs can share results. The current runner adds a full calculation to each trial it evaluates because incremental calculation returned stale balances in the longer plan.

### Results with annual account checks

The current runner recalculates 12 closing amounts per year from the original financial formulas. It checks the brokerage and its cost basis, the traditional account, each spouse’s Roth IRA and workplace Roth, the HSA, and the portfolio totals. A missing or nonzero check triggers one fresh calculation from the starting position. If the second check fails, Excel stops before replacing the saved results. The differences are visible in **Trial_Checks** columns N through Y.

| Excel for Mac run | Result |
| --- | --- |
| 100 pairs, ages 35–95 | Completed in 107.218 seconds. Excel calculated 187 outcomes and three ranked funding schedules directly. It reused 46 verified flexible results. One closing check needed a fresh calculation. All 233 outcomes and three funding histories matched across 80,454 fields. |
| 1,000 pairs, ages 0–125 | Stopped at trial 250 after 447.64 seconds. Its closing amounts still failed to reconcile after a fresh calculation. The prior results and completion record were retained. The workbook marked the attempted refresh as failed. |

The maximum-horizon run did not complete. Its elapsed time measures the failed attempt. The first recorded closing difference was a brokerage balance on trial 42. That trial recovered. After saving in Excel, all 595,814 stored trial and funding cells matched an untouched results bank saved in the same version of Excel. The previous completion date and assumptions also matched.

A separate Mac test replayed 11 trial selections, including the transitions from 41 to 42, 114 to 115 and 249 to 250. All 11 matched the independent results on the first attempt. The short sequence did not reproduce the failure seen during the full run. Recovery needs to pass the longer sequence before release.

In a longer test, the first incorrect annual balance and the ending balance matched the preceding trial on 30 occasions. Rebuilding Excel's dependencies corrected 29. It stopped at trial 839 after 15 minutes 38 seconds when balances were missing from age 99 onward. The discrepancy was already in the worksheet before the macro read the result.

The next test calculated each year's visible formulas in dependency order. Eight selected trials and six selections that changed the end age matched the independent model. A 100-pair run using that order throughout took 160.64 seconds. All 80,454 compared fields matched on the first attempt, including the three complete funding histories.

The same method completed 1,000 pairs through age 125 in 46 minutes 21 seconds. All 2,033 outcomes, the three ranked path IDs and their funding histories matched across 588,054 fields. Excel reused 501 flexible results and accepted every directly calculated path on its first attempt. This schedule preserved the worksheet calculations and exceeded the five-minute target. The fixture had no scheduled one-time expenses. The release also needs tests on the revision that carries those expenses into each trial.

### Earlier timings by plan length

The capacity test starts at age zero and ends at 125. It exercises all 126 annual positions, including the opening balance. The example has 61 positions, from age 35 through 95. Each run used the same seed and included all 33 required stress cases. These measurements preceded the 12 annual account checks.

| Paired paths | Example, incremental | Ages 0–125, incremental | Example, full calculation | Ages 0–125, full calculation |
| ---: | ---: | ---: | ---: | ---: |
| 1 | 16 sec | 37 sec | — | — |
| 10 | 21 sec | 1 min 1 sec | — | — |
| 100 | 58 sec | 3 min 50 sec | 73.02 sec | — |
| 1,000 | 7 min 55 sec | Stopped at trial 1569 | 14 min 31 sec | 32 min 1 sec |

The first 1,000-pair capacity run stopped at flexible trial 1569. All 595,814 stored cells in the previous 100-pair batch were unchanged, including individual outcomes and ranked funding histories. The workbook marked the attempted refresh as failed. When run separately in Excel, the fixed and flexible versions of this path agreed on every financial field. In a separate diagnostic, Excel’s incremental calculation carried the previous trial’s last 14 annual balances into trial 130. The terminal result and its source ledger agreed with each other but differed from the independent model. A full calculation corrected the 30 affected fields.

A second trace checked results after the earlier runner had already accepted them. Trial 115 retained the prior trial’s balances from ages 109 through 125. Its investment returns were current. The captured balances agreed with the stored ledger, and its cash-flow status showed OK. Recalculating the brokerage close from the opening balance, transactions and return exposed the first error at age 109.

In a desktop replay, the new checks rejected the single stale year and the 17-year stale sequence beginning at age 109. The same test passed before the stale amounts were inserted and after the original formulas were restored.

Before the annual account checks were added, a full calculation for each directly evaluated trial completed the 1,000-pair capacity run in 32 minutes 1 second. All 2,033 outcomes, the three ranked path IDs and their detailed funding histories matched across 588,054 fields. The workbook verified that 501 flexible paths could use their fixed partners’ results.

An edit to the active wage assumption took 0.74 seconds in the example and 1.69 seconds at maximum capacity. A full deterministic recalculation took 0.24 and 0.59 seconds, respectively. The market simulation runs separately on request.

### Interruption and recovery

In a disposable copy, Excel for Mac cancelled a run at the first reusable pair, detected a changed seed, rejected an invalid trial and recovered from an error after three result ranges had been written. Each test preserved the previous outcomes, funding examples, completion date, inspection selector and protection settings. A comparison after saving and reopening matched all 595,814 stored trial and funding cells to the original example.

When a required comparison formula was removed, Excel evaluated all 14 flexible trials directly. After the formula was restored, it reused all 14 and produced the same results. The harness called the public cancellation procedure. In a separate test, neither Esc nor Command–Period sent through desktop automation interrupted a live 45-second checkpoint. The previous financial results and Excel settings were preserved. The visible Cancel button still needs a verified click during a run.

## Release checks

First reproduce trial 250 from the preceding sequence and reconcile the first failing account against the independent annual results. Then test 1, 10 and 100 pairs on both plan lengths. Compare the 100-pair example with and without verified reuse under the same calculation and checking method. Require every outcome, rank and funding history to agree before interpreting the timing. Finally, run 1,000 pairs at each horizon and repeat the interruption, protection and six-page print checks on the same release revision. The targets remain two minutes for the example and five minutes for the maximum horizon. The owner uses Excel for Mac. A native Windows test remains open. [Release checklist](model-improvement-plan.md)
