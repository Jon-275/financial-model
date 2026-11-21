# Try the retirement planner

Start with fictional figures so you can learn the workflow before entering household information. No Python installation or repository clone is needed.

[Download for Excel for Mac](https://github.com/Jon-275/financial-model/raw/main/workbook/retirement-model.xlsm) · [Preview the six-page example report](examples/retirement-plan-example.pdf)

## Check compatibility

Use desktop Excel for Mac and save in `.xlsm` format. The market-risk calculation uses VBA macros; follow your organization's macro policy. Windows, LibreOffice and other spreadsheet apps are not verified release targets. Excel for the web can display saved results but cannot run the VBA analysis.

The base projection updates with Excel calculation set to **Automatic**. You can review inputs and the base result before running market risk. If your macro policy prevents the navigation buttons from working, select the **Inputs** and **Summary** sheet tabs directly.

## 1. Open a private example copy

Save the download somewhere private with a name you recognize, such as `retirement-plan-learning.xlsm`. Open **Start Here**, then **Open Inputs**. Select the example scenario in **Inputs B4**. The highlighted column is the editable scenario; column D holds the example, E compares later retirement and F compares spending 10% less in retirement.

Keep the fictional figures for this walkthrough. Your personal workbook will contain inputs for all three scenarios and saved market results.

## 2. Read the base result

Open **Summary**. Start with whether the base plan covers spending and taxes through the selected end age. If it does not, look for the first shortfall age and year. The base result uses the entered return assumptions; market simulations are a separate analysis.

The published example begins at age 35, retires at 60 and projects through 95. Its figures illustrate one set of assumptions. Your copy's saved report can differ if its version or settings differ.

## 3. Compare one change

Select the later-retirement scenario on **Inputs**, then return to **Summary**. Compare retirement age, funding and when income or savings pays the bills. Next try the lower-spending scenario.

For a controlled comparison, save the baseline Summary as a PDF, return to the example and change only the highlighted retirement-spending input. Review the base result again. Editing inputs makes the previous market-risk result out of date; it does not silently rerun simulations.

Do not interpret a positive ending balance alone as proof that every year's bills were paid. Account access and timing matter. [Reading the results](model-guide.md#read-the-results)

## 4. Run market risk when ready

Review **Asset_Allocation** and **Risk_Settings**. Use **100 paired paths** for an exploratory run, keeping the seed unchanged when comparing plans. On **Summary**, choose **Run market risk** with macros permitted.

Excel displays progress in its status bar. Wait for **Risk results are current**. For a final report, use 1,000 pairs. The documented example took about 8 minutes at 1,000 pairs; the full age range took 33 minutes. Runtime depends on the Mac and plan. Use the workbook's **Cancel** button if needed and wait for confirmation.

Read the success rate together with balance ranges, spending cuts and stress tests. It describes the simulated outcomes under your assumptions. [Market-risk interpretation](risk-analysis.md#read-the-summary)

## 5. Save a report, then build your own plan

Choose today's or future dollars on **Summary**. Select only that sheet, then **File → Print → Active Sheets** to print or save all six pages as a PDF. Check the scenario, dollar view and risk status first.

Create another private workbook copy for personal planning. Begin with ages, income, balances and spending, then review every relevant assumption, including healthcare and account access. The [full model guide](model-guide.md) explains the details and [items to estimate separately](model-guide.md#items-to-estimate-separately).

## If something does not work

| What you see | What to try |
| --- | --- |
| A navigation button does nothing | Open the named sheet tab directly. Check your organization's macro policy before running market risk. |
| Base results do not update | Check Excel calculation is set to Automatic. |
| Summary asks you to review an input | Follow the message beneath the status and resolve it before running risk. |
| Market results are out of date | Finish editing, then run market risk again. |
| A run fails | Read **Risk_Settings A44**, correct the reported issue and rerun. A failed run is not a current report. |
| The file lost its Run button or macro functionality after saving | Return to the original download and retain `.xlsm` format. |

[Ask for help](https://github.com/Jon-275/financial-model/issues/new/choose) with your Excel version, macOS version, steps and the exact message. Use fictional figures and remove personal information from screenshots. Do not upload a personal workbook or report.
