# First-use documentation and report evidence

This record covers the November 26, 2026 documentation rollout. The public download is v1.0.0. The workbook and PDF candidates use the fictional Example plan. Windows remains unverified.

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

## Promotion gate

The following evidence remains open. Keep the public downloads tied to v1.0.0 until the candidate qualifies.

- D2 staffed walkthrough, eligible consented baseline sessions and the follow-up study on the candidate.
- Native save and reopen of the candidate with a fresh reconciled market batch, and the desktop cancellation and recovery check. The first Mac candidate opened and recalculated. Save attempts produced no verifiable saved bytes. Repeat the native print matrix for the final workbook hash.
- Hands-on Excel keyboard and VoiceOver checks and assistive-technology reading of the tagged PDF.
- The cross-format release rehearsal for the candidate, including published workbook, PDF, guide and preview identity and rollback verification.

Windows remains unverified. Windows desktop Excel was unavailable for testing.
