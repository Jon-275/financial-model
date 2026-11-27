# Try the example

Open the fictional plan and compare retirement at age 60 and 65. Both base plans cover spending and taxes through age 95 under the example assumptions. You can see the comparison before running market risk.

[Download v1.0.1 for Excel for Mac](https://github.com/Jon-275/financial-model/releases/download/v1.0.1/retirement-model.xlsm) · [View the v1.0.1 example report](https://github.com/Jon-275/financial-model/releases/download/v1.0.1/retirement-plan-example.pdf)

## Start Here: save a copy

Open the workbook in desktop Excel for Mac and save a private copy in `.xlsm` format. Keep calculation set to **Automatic**. Release v1.0.1 was verified on desktop Excel for Mac 16.115.2. Windows is unverified.

On **Start Here**, choose **Open Inputs**. You can also select the **Inputs** sheet tab directly.

## Inputs: choose a scenario

In **Scenario to edit (1–3)** at **B4**, select **1** and confirm **Example plan**. Keep the fictional figures and choose **Open Summary** or the **Summary** tab.

## Summary: read and compare

1. On **Summary**, find the retirement age, the end age and the base result under those ages. **Example plan** retires at 60 and covers spending and taxes through 95.
2. Return to **Inputs**, select **2** in **B4** and confirm **Retire at 65**. Open **Summary** again. This plan retires at 65 and also covers spending and taxes through 95.
3. Check the risk status. **Choose Run market risk to update the results** means the previous market run is out of date for this scenario. The base result has already updated.

If a plan shows a shortfall, read the first unpaid cost and the age and year it occurs. A positive ending balance can coexist with an earlier unpaid bill. [Read the first-shortfall example](model-guide.md#first-shortfall)

## Asset_Allocation and Risk_Settings: review and run

When you want to test market outcomes, review the investment mix on **Asset_Allocation**. On **Risk_Settings**, enter **100** in **Paired market paths** (**B29**) for exploration. For comparisons, keep the seed, pair count, return convention and inflation settings consistent across plans.

If your organization permits macros, close other workbooks, return to **Summary** and choose **Run market risk**. Watch Excel's status bar and wait for **Risk results are current**. [Settings and result definitions](risk-analysis.md)

Use 1,000 pairs for a final report. Runtime depends on your Mac and plan. [Measured runtimes](risk-analysis.md#measured-runtime)

## Summary: save the report

On **Summary**, choose **Dollar view**. **Today's dollars** use opening-year prices. **Future dollars** use projected prices in each year. To include market results, confirm **Risk results are current**. You can still share the base projection while market results need updating.

Select only the **Summary** sheet. Choose **File → Print → Active Sheets** and check that the preview has six pages before printing or saving a PDF.

For your own plan, save another private copy and follow the [workbook guide](model-guide.md). Review all inputs before relying on the results. Save a report before changing a plan; a new market run replaces the old results.

## If something does not work

### Inputs need attention

Follow the message below the Summary result and check the selected scenario on **Inputs**. If results are slow to update, confirm calculation is set to **Automatic**.

### Risk is out of date

Finish editing and start a fresh run if you want updated results. For a failed run, check the message in **Risk_Settings!A44**. If **Summary** blanks while the status bar shows progress, allow the run time to finish. [Market risk and recovery](risk-analysis.md#run-the-analysis)

### Printing

If you see more or fewer than six pages, select only **Summary** and check **Active Sheets** in the print dialog. [Print your plan](model-guide.md#print-your-plan)

### Macros are blocked

Keep your existing macro policy. Use **Inputs** and **Summary** directly for the base projection. Excel for the web can show saved results. Running market risk requires desktop Excel with macros permitted.

[Ask for help](https://github.com/Jon-275/financial-model/issues/new/choose) with the planner version or download date, Excel and macOS versions, steps and exact message. Use fictional figures. Do not attach a personal workbook or report.
