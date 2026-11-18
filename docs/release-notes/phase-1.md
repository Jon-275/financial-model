# Phase 1 release record

Phase 1 replaces the workbook’s separate working and retirement funding shortcuts with an annual cash ledger. The selected horizon still stops at the end age entered on `Inputs`, up to 125. Existing scenario inputs remain intact. Confirm the new defaults in rows 86–98 before making decisions from the results.

| Area | Change |
| --- | --- |
| Working cash | The ledger separates taxable income from spendable cash. It records actual Roth IRA savings, any unmet savings target, and cash shortfalls separately. |
| Tax timing | A bounded search estimates tax on actual withdrawals for the current year, including the first and final retirement years. Unpaid tax carries forward and reduces after-tax net worth. Optional conversions are skipped when funding cannot support them. |
| Retirement bridge | Summary and annual draws use the same Roth eligibility calculation. The model keeps IRA and workplace Roth balances separate for each owner. |
| Summary and risk | Summary flags input and calculation problems and identifies the first year with a spending or tax shortfall. Risk sheets withhold headlines when the base plan cannot fund its obligations. |
| Traceability | `Cash_Flow`, `Roth_Access` and `Checks` expose cash, access and reconciliation. The `Funding_Calc` sheet contains the draw search. |

## Build and rollback

Use the [current build and test instructions](../../README.md#build-the-next-edition).

## Verification scope

On November 6, 2026, the saved workbook passed independent reconciliation across all active years. All 31 synthetic cases passed after recalculation in LibreOffice. They include taxable draws that cross a bracket, first- and final-year tax, carried tax payable, owner-specific Roth access, the five-tax-year boundary, a shortfall at age 90, and the age-125 limit. The scripted build also passed end to end.

Run the [acceptance tests](../../tests/README.md) after modifying funding logic. The workbook retains its five chart objects. Phase 1 verification used saved workbook values and LibreOffice. See [Phase 3](phase-3.md#microsoft-excel-verification) for later Excel verification.

Remaining phases are listed in the [implementation plan](../../PLAN.md). Phase 1 does not implement a new tax engine, per-path risk tax and access, or nonqualified workplace Roth draws.
