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
| Annual balances and selected funding examples | An early run carried values from the preceding path into seven fixed paths and changed all three selected examples. The runner now checks each year's captured balances against its ledger and retries a suspicious repeated ending. The full Mac example matched all 573,306 trial fields, three selected path IDs, 378 annual balances and 14,367 detailed funding fields. |

## What takes time in Excel

A 100-pair run on the earlier worksheet runner took 70.13 seconds, including all 33 stresses. Excel spent 67.22 seconds recalculating worksheets. Reading and writing cells took 0.21 seconds, including publication. Updating the final report took 0.43 seconds.

These tests used the same example inputs and retained the visible annual worksheets as the reference.

| Excel for Mac experiment | Result |
| --- | --- |
| Remove unused withdrawal-search formulas for this plan | The saved test workbook had 173,186 search formulas, down from 560,571. The 100-pair run took 70.47 seconds, including 67.19 seconds of worksheet calculation and one full retry. All outcomes and ranked funding schedules matched. Removing unused formulas gave no measurable gain in this run. |
| Calculate in a temporary workbook containing only the 21 required sheets | Excel completed all 233 outcomes and three ranked funding schedules. Every saved result matched the reference run. This test recorded the completion time only, so its duration is unknown. |
| Calculate market paths and account balances in VBA arrays | Generating all 2,033 market paths took 2.49 seconds for 122,005 modeled years. Replaying 80,500 annual account closings took 0.19 seconds. Thirteen trials matched the worksheet reference for returns, prices and balances. The test used withdrawals and cash flows already calculated by the worksheets. Withdrawals, taxes and healthcare still ran on the worksheets. |

For a quick comparison, set **Paired market paths** on **Risk_Settings** to 100 and run the scenarios. Set it to 1,000 and run again for the final report. Both runs include the 33 stress cases. More paths reduce the sampling variation in the reported funding rate.

### Fixed and flexible spending on the same market path

The runner checks whether flexible spending would change any year's cash flow. It compares the opening portfolio with the spending-cut threshold and reads the possible reduction from **Market_Paths**. When every check passes, it records the verified amounts under both trial numbers. If any year could differ, Excel calculates the flexible path separately. Users can still select either trial and inspect its annual formulas.

Excel for Mac 16.115 on macOS 26.6.2 ran the same fictional plan with the same seed and assumptions. The clock covered the workbook run through publication and restoration of Excel's settings. Timing stopped before Excel saved the test copy.

| Paired market paths | Earlier worksheet runner | With the paired-path check | Flexible results reused | Results checked |
| ---: | ---: | ---: | ---: | --- |
| 100 | 71.16 seconds | 58.07 seconds | 46 | 233 outcomes and three detailed funding examples matched across 80,454 fields. |
| 1,000 | 12 min 7 sec | 7 min 55 sec | 515 | 2,033 outcomes and three detailed funding examples matched across 588,054 fields. |

The 100-pair comparison used the same timing procedure for both runs. The earlier 1,000-pair measurement was taken with other tests running. All runs include the 33 stress cases. The full example still exceeds the two-minute target.

### Longest supported plan

The capacity test starts at age zero and ends at 125. It exercises all 126 annual positions, including the opening balance. The example has 61 positions, from age 35 through 95. Each run uses the same seed and includes all 33 required stress cases.

| Paired paths | Example | Ages 0–125 |
| ---: | ---: | ---: |
| 1 | 16 sec | 37 sec |
| 10 | 21 sec | 1 min 1 sec |
| 100 | 58 sec | 3 min 50 sec |
| 1,000 | 7 min 55 sec | Under investigation at trial 1569 |

The first 1,000-pair capacity run stopped at flexible trial 1569. All 595,814 stored cells in the previous 100-pair batch were unchanged, including individual outcomes and ranked funding histories. The workbook marked the attempted refresh as failed. When run separately in Excel, the fixed and flexible versions of this path agreed on every financial field.

An edit to the active wage assumption took 0.74 seconds in the example and 1.69 seconds at maximum capacity. A full deterministic recalculation took 0.24 and 0.59 seconds, respectively. The market simulation runs separately on request.

### Interruption and recovery

In a disposable copy, Excel for Mac cancelled a run at the first reusable pair, detected a changed seed, rejected an invalid trial and recovered from an error after three result ranges had been written. Each test preserved the previous outcomes, funding examples, completion date, inspection selector and protection settings. A comparison after saving and reopening matched all 595,814 stored trial and funding cells to the original example.

When a required comparison formula was removed, Excel evaluated all 14 flexible trials directly. After the formula was restored, it reused all 14 and produced the same results. The harness called the public cancellation procedure. In a separate test, neither Esc nor Command–Period sent through desktop automation interrupted a live 45-second checkpoint. The previous financial results and Excel settings were preserved. The visible Cancel button still needs a verified click during a run.

## Release checks

The full risk refresh still needs faster annual calculations. Recheck financial results, controls and the six-page printout on the final release revision. The owner uses Excel for Mac. A native Windows test remains open. [Release checklist](model-improvement-plan.md)
