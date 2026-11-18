# Printable plan and workbook schedules

November 16, 2026

## Workbook changes

- This edition brought all five report pages onto **Summary**, including the input, spending and account decisions. The example opens with its saved market risk results. **Annual Detail** and **Risk Results** follow the dollar view on **Summary**. The base projection and risk simulator use the HSA’s own investment mix. [Printable report update](https://github.com/Jon-275/financial-model/pull/12)
- **Income_Targeting** holds the Roth conversion comparison and its charts. **Tax_Tables** is before **Appendix**. **Conversion Detail** is behind it and shows the results at each tested amount. **Report_Data** has one row per modeled year. [Schedule layout update](https://github.com/Jon-275/financial-model/pull/13)
- **Planning Metrics** sits next to **Summary** with grouped supporting figures, labeled dollar units and links to the annual schedules. [Planning Metrics and documentation update](https://github.com/Jon-275/financial-model/pull/14)
- Annual report and schedule shading ends with the data. The solver tabs label skipped trials and final draws, and the **Calculation step** filter controls which rows are visible. **Risk Results** shows the simulation settings. **Risk_Data** holds the saved annual balances, early-loss tests, funding examples, and input comparison. [Workbook schedules](../workbook-schedules.md)

## Verification

Excel for Mac print preview showed five pages from **Summary** on US Letter and A4. The schedule update was also checked in Excel for Mac. Shading ended at age 95 on **Annual Detail** and **Risk Results**, the conversion charts displayed, and the solver filter hid skipped rows on a test copy. The current workbook has six report pages. [Print instructions](../model-guide.md#print-your-plan)

The **Planning Metrics** update matched all 37 moved calculation results and the printed content, apart from the preparation date. The new tab was inspected in Excel for Mac, and US Letter and A4 still previewed five pages. [Planning Metrics verification](https://github.com/Jon-275/financial-model/pull/14)

All 81 workbook scenarios and 35 risk tests passed. The refresh test covered a plan ending at 125, both dollar views, scenario changes, and preservation of the source workbook. The schedule update matched 12,444 saved numeric values across **Projection**, **Cash_Flow**, **Roth_Access**, **Healthcare**, and **Checks**. Workbook checks, Greptile Review, and Gemini Review passed on the schedule update. [Schedule checks and reviews](https://github.com/Jon-275/financial-model/pull/13)
