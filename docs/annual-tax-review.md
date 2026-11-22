# Annual tax data review

For maintainers. [Documentation index](README.md) · [Tax assumptions](model-methods.md#tax-rates-and-healthcare-credits)

Check tax and benefit figures before the first release for each tax year and whenever an agency changes a figure the workbook uses. Record the date, source and effective year even if the figures stay the same.

## Check the published figures

Open **Tax_Tables** and use **Find a rate or limit** to reach each table. Check the figure and its effective tax, payroll or coverage year against the agency release. The workbook notes identify the source and explain estimates for later years.

| Check | Where to verify |
| --- | --- |
| Federal income tax brackets, deductions, capital gains rates and alternative minimum tax | [IRS Internal Revenue Bulletin](https://www.irs.gov/irb) |
| Social Security wage limit | [Social Security taxable wage limits by year](https://www.ssa.gov/oact/cola/cbb.html) |
| California brackets, deductions, credits and disability insurance | [Franchise Tax Board](https://www.ftb.ca.gov/file/personal/tax-calculator-tables-rates.asp) · [Employment Development Department](https://edd.ca.gov/en/payroll_taxes/rates_and_withholding) |
| IRA contribution and income limits, and required withdrawals | [IRS retirement contribution guidance](https://www.irs.gov/retirement-plans/plan-participant-employee/retirement-topics-contributions) · [IRS distribution guidance](https://www.irs.gov/publications/p590b) |
| Marketplace credit percentages and the poverty guidelines used for the coverage year | [IRS Internal Revenue Bulletin](https://www.irs.gov/irb) · [HHS poverty guidelines](https://aspe.hhs.gov/topics/poverty-economic-mobility/poverty-guidelines) |
| Medicare income surcharges and income thresholds | [How Social Security uses income](https://www.ssa.gov/OP_Home/handbook/handbook.25/handbook-2504.html) · [2026 Medicare premium tables](https://www.cms.gov/newsroom/fact-sheets/2026-medicare-parts-b-premiums-deductibles) · [2027 Medicare premium tables](https://www.cms.gov/newsroom/fact-sheets/2027-medicare-parts-b-premiums-deductibles) |

Use the agency table for the year you are reviewing. If a figure is not yet available, label it as an estimate and record how the model calculates it.

## Update the workbook source

- Update federal, California and payroll figures at their source. Check the formulas when a rate changes how or when the tax applies. [Published tax data](../tools/published_tax.py) · [Tax formulas](../tools/tax_policy.py)
- Update IRA limits in the contribution calculations. Update Marketplace figures and coverage years in the healthcare calculations. [Roth contributions](../tools/roth_contributions.py) · [Healthcare](../tools/healthcare.py)
- Check the Medicare surcharge and required withdrawal tables in `workbook/build-template.xlsx`. If the template changes, run the content and privacy checks before approving its checksum. [Workbook privacy checks](../tests/README.md#share-an-updated-workbook)
- Apply the same effective years in the independent risk calculation. If **Tax_Tables** moves a figure, update the workbook reader and its tests. [Independent tax calculation](../tools/risk_engine/tax.py) · [Workbook reader](../tools/risk_engine/scenario.py)

## Verify and record the release

1. Build the updated workbook from the fictional template. [Build instructions](../tests/README.md#build-the-workbook)
2. In LibreOffice Calc, recalculate cases before and after each changed figure takes effect. Compare the results with the published figures and the independent financial calculation. Include the base plan and a market trial with the plan’s assumed returns. [Financial and market checks](../tests/README.md#check-a-change)
3. Verify the updated workbook in desktop Excel for Mac. Submit the source, updated test cases and rebuilt fictional workbook in one pull request. [Excel checks](../tests/README.md#verify-a-changed-workbook-in-excel)
4. Record the review date, effective years, agency links, estimates and test results. If the workbook changed, link the merged pull request.

To restore a release, use a pull request to restore its matching source, test data and workbook together. Run the release checks on that set before publishing.
