# Release readiness

The macro-enabled workbook is available for testing in desktop Excel for Mac. Users enter their assumptions, run market risk when ready and print the six-page Summary. [Workbook and first steps](../README.md)

## Work before publication

| Area | Next step |
| --- | --- |
| First use | The owner will follow the first-use guide on a private copy, refresh market risk and preview Summary. Clarify anything that is hard to find or interpret before sharing a fresh fictional copy with colleagues. [First-use guide](model-guide.md#run-your-first-plan) |
| Refresh time | The targets for 1,000 paired paths are two minutes for the example and five minutes through age 125. Worksheet runs on the test Mac took 7 min 52 sec and 33 min 10 sec, respectively. Measure which parts take the most time before making further changes. |
| Privacy | Resolve one outstanding commit-author email item privately. Verify that the workbook and sample PDF contain only fictional data. |
| Licensing | The owner’s license decision is pending. Confirm rights to the code, workbook and documentation and review third-party notices before applying a license. |
| Final check | Open the exact release file in Excel for Mac. Check calculation, the saved report and print output on Letter and A4. Automated cancellation is verified. The manual Cancel-button check is deferred. [Test and release guide](../tests/README.md#microsoft-excel-release-check) |
| Publish | Confirm the required reviews have passed, verify that the sample PDF matches the release workbook, and update the download and its guide together. |

## Maintenance

Review tax data annually and when a source agency changes a figure used in the workbook. Have a second maintainer rehearse the build, verification and recovery process with fictional data. [Annual tax review](annual-tax-review.md) · [Build and release checks](../tests/README.md)
