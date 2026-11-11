# Phase 1 release record

Phase 1 replaces the workbook’s separate working and retirement funding shortcuts with an annual cash ledger. The selected horizon still stops at the end age entered on `Inputs`, up to 125. Existing scenario inputs remain intact. Confirm the new defaults in rows 86–98 before making decisions from the results.

| Area | Change |
| --- | --- |
| Working cash | The ledger separates taxable income from spendable cash. It records actual Roth IRA savings, any unmet savings target, and cash shortfalls separately. |
| Tax timing | A bounded search estimates tax on actual withdrawals for the current year, including the first and final retirement years. Unpaid tax carries forward and reduces after-tax net worth. Optional conversions are skipped when funding cannot support them. |
| Retirement bridge | Summary and annual draws use the same Roth eligibility calculation. The model keeps IRA and workplace Roth balances separate for each owner. |
| Summary and risk | Summary flags input and calculation problems and identifies the first year with a spending or tax shortfall. Risk sheets withhold headlines when the base plan cannot fund its obligations. |
| Traceability | `Cash_Flow`, `Roth_Access` and `Checks` expose cash, access and reconciliation. The hidden `Funding_Calc` sheet contains the draw search. |

## Build and rollback

Use the pre-Phase 1 horizon workbook saved in local Git. The original repository workbook predates that horizon change and is not the build baseline. Recover the baseline to a local scratch path, then run the build from the repository root:

```sh
ARTIFACT_NODE_MODULES=/path/to/workspace/node_modules \
  tools/build_model.sh /path/to/baseline.xlsx /path/to/retirement-model.xlsx
```

The build needs Python with the dependencies pinned in `tools/requirements.txt`, Node with the workspace Artifact Tool installed, and headless LibreOffice. Set `PYTHON_BIN`, `NODE_BIN` and `SOFFICE_BIN` if they are not on your path. The wrapper compiles and authors workbook formulas, preserves workbook objects from the baseline, recalculates a separate copy, and runs `tests/model_acceptance.py` before publishing the file. Use a new output path named `retirement-model.xlsx` to preserve the existing chart links.

To roll back, extract the prior workbook revision to a separate path with `git show REVISION:workbook/retirement-model.xlsx`, verify it, and then replace the working workbook. Keep the pre-Phase 1 baseline local. Share only synthetic test inputs.

## Verification scope

On November 6, 2026, the saved workbook passed independent reconciliation across all active years. All 31 synthetic cases passed after recalculation in LibreOffice. They include taxable draws that cross a bracket, first- and final-year tax, carried tax payable, owner-specific Roth access, the five-tax-year boundary, a shortfall at age 90, and the age-125 limit. The scripted build also passed end to end.

Run the [acceptance tests](../../tests/README.md) after modifying funding logic. The workbook retains its five chart objects. Native Excel recalculation and interactive chart refresh are unverified; check both before distributing to Excel users.

Remaining phases are listed in the [implementation plan](../../PLAN.md). Phase 1 does not implement a new tax engine, per-path risk tax and access, or nonqualified workplace Roth draws.

Local build baseline: the original age-125 template (`workbook/retirement-model.xlsx`).
