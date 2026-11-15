# Annual tax data review

Review the tax data each year before publishing the workbook for the new tax year. Review it again when an agency publishes or revises figures the model uses. Record the review even when the figures stay the same.

## Check the published figures

Start on **Tax_Tables**. The notes name the sources, the years the figures apply to, and how the model estimates later years. Confirm each figure against the agency release for the correct tax, payroll or coverage year.

| Check | Where to verify |
| --- | --- |
| Federal income tax brackets, deductions, capital gains rates and alternative minimum tax | [IRS Internal Revenue Bulletin](https://www.irs.gov/irb) |
| Social Security wage limit | [Social Security Administration](https://www.ssa.gov/oact/cola/cbb.html) |
| California brackets, deductions, credits and disability insurance | [Franchise Tax Board](https://www.ftb.ca.gov/file/personal/tax-calculator-tables-rates.asp) · [Employment Development Department](https://edd.ca.gov/en/payroll_taxes/rates_and_withholding) |
| IRA contribution and income limits, and required withdrawals | [IRS retirement contribution guidance](https://www.irs.gov/retirement-plans/plan-participant-employee/retirement-topics-contributions) · [IRS distribution guidance](https://www.irs.gov/publications/p590b) |
| Marketplace credit percentages and the poverty guidelines used for the coverage year | [IRS Internal Revenue Bulletin](https://www.irs.gov/irb) · [HHS poverty guidelines](https://aspe.hhs.gov/topics/poverty-economic-mobility/poverty-guidelines) |
| Medicare income surcharges and income thresholds | [Social Security Handbook](https://www.ssa.gov/OP_Home/handbook/handbook.25/handbook-2504.html) · [CMS newsroom for the year’s Medicare premiums fact sheet](https://www.cms.gov/newsroom) |

If an agency has yet to publish a value, record the estimate and how the model carries it forward. Keep published figures tied to the year they apply to.

## Update the workbook source

- Update federal, California and payroll values in [published tax data](../tools/published_tax.py). Check [tax formulas](../tools/tax_policy.py) when an effective year or calculation changes.
- Update IRA limits in [Roth contributions](../tools/roth_contributions.py). Update Marketplace figures and coverage-year logic in [healthcare](../tools/healthcare.py).
- Check any affected values in `workbook/build-template.xlsx`, including the Medicare surcharge and required withdrawal tables. Follow the [template privacy check](../README.md#keep-private-workbook-copies-out-of-git) if the template changes.
- Check the [risk engine’s tax rules](../tools/risk_engine/tax.py) for matching year boundaries. If the **Tax_Tables** layout changes, update the [workbook reader](../tools/risk_engine/scenario.py) and its tests.

## Verify and record the release

1. [Build a candidate workbook](../README.md#rebuild-the-workbook) with the fictional scenario.
2. Run the [workbook and risk tests](../tests/README.md) against the candidate. Exercise the years before and after each changed figure takes effect. Check the workbook against the risk report with returns set to the plan assumptions.
3. Complete the [desktop Excel release check](../tests/README.md#microsoft-excel-release-check) and open a PR with the source changes, fixtures and rebuilt workbook.
4. Record the review date, the years checked, the agency links, any values or estimates that changed, and the checks that passed. Link the merged PR from the [release records](release-notes/phase-3.md).

If a release needs to be restored, use a PR to restore the matching source, test data and workbook from the last verified release.
