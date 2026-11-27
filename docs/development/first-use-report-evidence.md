# First-use documentation and report evidence

This record covers the November 26–27, 2026 verification for the v1.0.1 workbook and report. The fictional Example plan is used throughout. The frozen study baseline is v1.0.0. Windows remains unverified.

## Published release controls

| Change | Evidence |
| --- | --- |
| Release facts and links | [PR #66](https://github.com/Jon-275/financial-model/pull/66) pins the public workbook and PDF to v1.0.0 and gives measured run times one home. |
| Study baseline | [PR #67](https://github.com/Jon-275/financial-model/pull/67) pins the archived product page, compatible guides, workbook and PDF. Setup verified the workbook and PDF checksums. |
| Website and guide flow | [PR #68](https://github.com/Jon-275/financial-model/pull/68) covers acquisition, the 60-to-65 comparison, current and stale results, and printing. Headless Chrome checked keyboard skip and FAQ controls, report images and reflow at widths of 320, 768 and 1,440 CSS pixels. |
| Release identity | [PR #69](https://github.com/Jon-275/financial-model/pull/69) uses a manifest to check the workbook, report, license and preview digests before packaging. The staged v1.0.0 download set matched the live GitHub release byte for byte. |

## Candidate workbook and example report

| Item | Identity |
| --- | --- |
| Workbook authoring source | `a8f24256d1afc3c11db95391718b2d20fc5be6d7` ([PR #70](https://github.com/Jon-275/financial-model/pull/70)) |
| Fictional workbook SHA-256 | `8c9dade568ebf44df73d313adb21eef247dcdd3aafe8ccd680cebbd5a15034e1` |
| Tagged example PDF SHA-256 | `17f0ce69e706789894480b9282da13668d6d7db02e7074e98b471cb0dc17e414` |
| Report tooling | [PR #71](https://github.com/Jon-275/financial-model/pull/71) |

### Workbook checks

The release and candidate contain the same model inputs and calculation formulas. The example comparison reconciled 2,033 outcomes and 589,709 fields, including the three complete ranked funding ledgers. The embedded macro project verified against source. The privacy audit verified both the exact published v1.0.0 package and the candidate.

On Excel for Mac 16.115.2 (16.115.26112215), the first candidate (`e3fadd09654b9ccccfabae50c0e5f4f1a7c19fcb797c143b971d394f5165ba26`) opened and recalculated. Switching to scenario 2 kept the base projection funded to age 95 and suppressed the stale market count. Returning to the exact inputs restored the saved count for the active scenario. Changing dollar view kept that count current. Calc also exercised the invalid scenario selector and the long scenario-name presentation.

### Native print and tagged PDF

Desktop Excel exported the first candidate's macro-free print projections on Letter and A4. Each Summary had six pages in both dollar views, with 335 Summary cells matched per export. The band, median and dashed base line matched workbook values within 0.847 PDF points across the native exports. The separate Annual Detail outputs covered all ages 35 through 95 in two pages per paper size.

The review workbook differs from the first candidate in one Start Here instruction and six refreshed timestamp fields. The model, chart, saved financial values and VBA match. The tagged example used a separate Calc export of a disposable workbook projection. It verified 4,072 repeated assumption records before using the temporary Calc encoding. The exported PDF kept its vector content, human-readable age ticks, six bookmarks and headings, scoped data-table headers and a chart alternative. The workbook-to-PDF check matched six Letter pages, 335 printed cells and 244 chart values, with a maximum difference of 0.033 points. PNG previews came from pages 1, 3 and 5 of that tagged PDF.

## Mac release verification follow-up

Excel for Mac 16.115.2 exported the review workbook’s four Summary variants on Letter and A4 in both dollar views. All four contained six pages and matched 335 Summary cells per export. The vector chart comparisons had a maximum difference of 0.847 PDF points. Both Annual Detail exports contained two pages and covered ages 35 through 95.

The Excel Accessibility Checker found three charts without alternative text and 31 merged-cell navigation warnings in the review workbook. Keyboard testing reached the scenario selector with **Control+G**, entered scenario 2 in **Inputs B4**, and reached the base funding answer in **Summary B8**. The presentation source now describes all three charts and gives the **Open Summary** link in **Inputs G4** its own cell. The remaining merged layouts need review on the rebuilt workbook. [Excel keyboard shortcuts](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)

The rebuilt fictional workbook has SHA-256 `5fd56b2259bac7afe740b1b8cf789aebd919e45d3443f588a1e6261d5db24d45`. Its matching tagged PDF has SHA-256 `e3ab7ae134edae2775dbbfd71b94e9bb41a0de7a50b1519cc3f01e0130edb581`. The PDF matched 335 Summary cells and 244 chart values. The workbook privacy and embedded macro checks passed. The saved trial, funding ledger, chart data, and base projection and cash flow worksheets matched the prior candidate. Native and accessibility checks are tracked by exact workbook hash.

The facilitator materials include the checksum-verified frozen v1.0.0 site, shortfall fixture and coded forms for baseline and candidate sessions. Participant scheduling and coded observations are pending.

### Saved native risk run

On November 26, 2026, Excel for Mac 16.115.2 (16.115.26112215) on macOS 26.7.3 (25G342) ran the rebuilt workbook in isolation with 1,000 pairs and seed 2026. The exact starting workbook was `5fd56b2259bac7afe740b1b8cf789aebd919e45d3443f588a1e6261d5db24d45`. The fresh saved capture was `5ab69cd3c864918a7d0baeac2d6bc8fc9dddb8f9d7c33d992c563416768f3f9d`. The capture stayed local as verification evidence.

The independent comparison reconciled all 2,033 outcomes and 588,054 fields, including all three ranked funding ledgers. The completed batch reports 515 verified flexible path reuses. The saved Summary reads **999 of 1,000 paths** and displays current risk, with a November 26 run time of 21:13. The saved file reopened in desktop Excel with the same SHA-256, batch timestamp, complete state, 1,000-pair count and visible result.

A separate 100-pair Mac run reconciled 233 outcomes and 80,454 fields, including the three ranked funding ledgers. The Cancel command timed out while Excel was calculating. On November 27, a run started from the Summary button reached 8 of 236 checks in the Excel status bar. The sheet body was blank at that point, including the Cancel control. The button and keyboard attempts ended with a completed 100-pair batch. The saved button-test result independently reconciled all 233 outcomes and 80,454 fields. The retained-batch and retry checks require an observed cancellation. The v1.0.0 and v1.0.1 workbooks have byte-identical embedded VBA projects. Track the Mac interruption limitation and its native recovery test in D3.

The exact rebuilt workbook also supplied all six Mac print projections. Each Summary export contained six pages in the expected dollar view and paper size, with 335 matching printed cells. The desktop band, median and dashed base line matched within 0.847 PDF points. Both Annual Detail outputs contained two pages and all ages 35–95.

### Release packaging and accessibility

A local rehearsal rejected the rebuilt workbook under the existing v1.0.0 manifest, verified the candidate workbook, PDF and all previews together under their own digests, then restored the approved v1.0.0 files and verified their digests. The rehearsal ran in a temporary staging directory.

The rebuilt PDF has one main heading, five section headings, six bookmarks, four data tables, 22 scoped column headers and a description for the chart. macOS PDFKit read the six titles and the funding, simulation count and freshness text. The earlier Excel checker’s merged-cell flags remaining after the source fix cover section headings, captions and notes, including the first-shortfall explanation on Summary. On November 27, the project owner scoped this release’s accessibility evidence to the documented keyboard, Excel Accessibility Checker and tagged PDF checks. VoiceOver was outside the release gate.

## Release and follow-up

The v1.0.1 presentation release uses the reconciled Mac run, save and reopen, print matrix, workbook and PDF accessibility checks described above, and an exact-byte publication and rollback check. The release notes and risk guide disclose the Mac cancellation limitation. D3 carries the cancellation fix and recovery proof for the native runner.

D2 runs in parallel with the release. Use the frozen v1.0.0 files for the staffed walkthrough and consented baseline sessions. Use the verified v1.0.1 files for the follow-up. Feed the coded observations into U1.

Windows remains unverified. Windows desktop Excel was unavailable for testing.
