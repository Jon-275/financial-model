# Publication source cleanup

November 17, 2026

## Changes

The build template and distributed example use synthetic household inputs and investment profiles. The profiles illustrate diversified growth, a balanced retirement mix and a larger reserve. Their weights are defined in the source fixture.

The template stores formulas without financial result caches. Its chart definitions also omit saved financial values. The privacy guard checks inputs, allocations, stored text, unused shared strings, external data sources and inherited metadata independently of approved checksums.

The build uses Python and publicly installable libraries. The native dashboard layout is stored as an inspectable XML definition. The build and source audit preserve the existing workbook formulas, charts and sheet visibility.

All 12 completed remote branches were deleted after their final commits were matched to merged PRs. Automatic branch deletion after future merges is enabled.

## Verification and remaining work

Nine privacy regression tests cover deliberately altered inputs, allocations, unused text, metadata, extended company and manager properties, formula caches and additional notes outside the input ranges. The release artifact also has a full-file approval checksum.

The public build passed the release check, the no-shock comparison across 60 years and 5,400 values, and all 35 risk tests. The rebuilt example includes 1,000 paired paths. Excel for Mac opened it without a repair prompt, displayed the native portfolio chart and synthetic profiles, and previewed five Summary pages on US Letter and A4.

Retained history was rewritten and the reviewed synthetic files were restored. The current source tree matched before and after the rewrite. A fresh clone passed the content audit and contained only the two approved workbook blobs. The owner's checkout was replaced; 12 earlier task clones were disabled for normal pushes. The original required reviews and branch rules were restored. Automatic branch deletion was verified with the merged publication cleanup PR.

GitHub's read-only PR references and cached views remain pending Support review. The colleague's review clone still needs replacement and confirmation. [Cleanup and clone instructions](../history-cleanup.md) · [Improvement plan](../model-improvement-plan.md)
