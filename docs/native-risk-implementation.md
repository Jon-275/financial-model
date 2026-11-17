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
| Mac run times | Ten market pairs and 33 stresses completed in 64 seconds. One hundred pairs and 33 stresses completed in 84 seconds. Both timings include other tests running on the Mac. |
| Repeated trial values | In the earlier 1,000-pair desktop run, seven fixed paths carried values from the preceding path and changed the three selected funding examples. The revised runner reconciles the annual capture and retries a repeated positive ending after a changed history. A fresh 100-pair Mac run completed all 233 outcomes in 78 seconds. All 65,706 primary fields and 378 ranked annual balances matched the independent calculation. |
| Full native run | Excel for Mac completed 1,000 pairs and 33 stresses in 12 minutes 7 seconds with other tests running. All 2,033 outcomes agreed with the independent example, including 573,306 primary fields, all three selected path IDs, 378 annual balances and 14,367 funding-ledger fields. The run was compared with the finished report workbook. The final change only adjusted the nonprinting Run and Cancel labels. |

## What takes time in Excel

A separate Excel for Mac profiling run covered 100 pairs and all 33 stresses. Of its 70.13 seconds, Excel spent 67.22 seconds (95.85%) calculating and waiting for worksheets. Explicit worksheet reads and writes, including publication, took 0.21 seconds. The final report recalculation took 0.43 seconds and progress updates took 0.66 seconds. There were no retries. Improving worksheet calculation time has the greatest potential to bring the full run closer to the two-minute goal.

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

## Release checks

Verify a successful run after cancellation, changed assumptions, error handling, repeated runs, saved results after reopening and the unchanged base plan. Measure runs at the longest supported horizon and investigate the difference between the 12-minute full example run and the two-minute goal. Record the Excel version and test conditions. Recheck the finished workbook and printed report on the supported desktop platforms. [Release checklist](model-improvement-plan.md)
