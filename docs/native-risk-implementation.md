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

## Release checks

Verify a successful run after cancellation, changed assumptions, error handling, repeated runs, saved results after reopening and the unchanged base plan. Measure 1,000 paired paths for the example and repeat the 10-, 100- and 1,000-pair runs at the longest supported horizon. Record the Excel version and test conditions. Recheck the finished workbook and printed report on the supported desktop platforms. [Release checklist](model-improvement-plan.md)
