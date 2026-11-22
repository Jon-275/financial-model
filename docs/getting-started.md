# Try the example

Use fictional figures for your first run. You do not need Python or a copy of the source code.

[Download for Excel for Mac](https://github.com/Jon-275/financial-model/releases/latest/download/retirement-model.xlsm) · [View the example report](examples/retirement-plan-example.pdf)

## Start Here: save a copy

Open the workbook in desktop Excel for Mac and save a private copy in `.xlsm` format. Keep calculation set to **Automatic**. Macros are required to run market risk; follow your organization's policy.

Choose **Open Inputs**. If a navigation button does not work, select the sheet tab directly. Excel for the web can show saved results but cannot run the analysis. Windows and other spreadsheet apps are not verified release targets.

## Inputs: choose a scenario

Select the example in **B4** and keep its fictional figures. Column D holds the example, E compares later retirement and F compares spending 10% less in retirement. Edit only the shaded input cells for the selected scenario.

## Summary: read and compare

Read whether the base plan covers spending and taxes through the end age. If not, find the first shortfall age and year.

Return to **Inputs**, select the later-retirement or lower-spending scenario, then compare **Summary**. A positive ending balance does not mean every year's bills were paid. Check when income starts and savings become available.

## Asset_Allocation and Risk_Settings: review and run

Review the investment mix on **Asset_Allocation**. On **Risk_Settings**, choose **100 paired paths** for a first run. Keep the seed unchanged when comparing plans.

Close other workbooks, then return to **Summary** and choose **Run market risk**. Watch Excel's status bar and wait for **Risk results are current**. [Settings and result definitions](risk-analysis.md)

Use 1,000 pairs for a final report. The documented final example took about 8 minutes; the full age range took 33 minutes. Your Mac and plan affect runtime.

## Summary: save the report

Choose today's or future dollars. Select only **Summary**, then **File → Print → Active Sheets** to print or save all six pages as a PDF.

For your own plan, save another private copy and follow the [workbook guide](model-guide.md). Review all inputs before relying on the results. Save a report before changing a plan; a new market run replaces the old results.

## If something does not work

| Problem | What to do |
| --- | --- |
| A button does nothing | Select the sheet tab directly. Check your macro policy before running market risk. |
| Results do not update | Set Excel calculation to Automatic. |
| Summary asks you to review an input | Follow the note beneath the status and correct the input. |
| Risk results are out of date | Finish editing, then run market risk again. |
| A run fails | Read the message in **Risk_Settings A44**, correct the issue and rerun. |
| You need to stop a run | Choose **Cancel** and wait for confirmation. |

[Ask for help](https://github.com/Jon-275/financial-model/issues/new/choose) with the planner version or download date, Excel and macOS versions, steps and exact message. Use fictional figures. Do not attach a personal workbook or report.
