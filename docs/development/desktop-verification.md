# Desktop Excel verification and performance

Use the [enhancement ledger](enhancement-plan.md) to track the implementation and evidence. A Calc or Python pass supplies model evidence. Record repair, execution, cancellation, save/reopen and pagination from desktop Excel on each supported platform.

## D1. Build fixtures and check evidence

The desktop benchmark builder creates a disposable `.xlsm` from a workbook that matches both its published SHA-256 and the tracked workbook at the specified Git revision. It writes the declared inputs and records the resulting input hash and resolved commit in `fixture.json`. Replay uses that commit even if the original branch name moves. The released fictional snapshot is the seed-2026 reference for the standard example, including a 100-pair subset with independently ranked funding histories. For an edited case, the builder independently authors and packages a fresh pre-VBA `.xlsx` from the pinned source and template. It applies the same overrides and gives the empty result batch to `native_bootstrap`. [Benchmark builder](../../tests/desktop_benchmark.py) · [Bootstrap](../../tools/native_bootstrap.py)

| Fixture ID | Inputs or action | Evidence |
| --- | --- | --- |
| `example-95` | Released Example plan, opening age 35, end age 95, seed 2026. Run 100 and 1,000 pairs. | 233 and 2,033 outcomes respectively, including the 33 stresses. Independent reference, reports and timings. |
| `extended-125` | Same released inputs with `Inputs!D53=125`. Run 100 and 1,000 pairs with seed 2026. | Ending age, report and independently rebuilt reference for the extended scenario. |
| `full-0-125` | Explicit synthetic maximum-capacity case. Set fictional `Inputs!D8=0`, `D9=0` and `D53=125` on the release fixture and preserve the rest of the declared inputs. | Maximum model capacity, closure checks and an independently rebuilt reference. Preserve this workload apart from `extended-125`. |
| `stale-and-invalid` | Change an active input, swap scenarios, change dollar view, enter invalid age, restore the original input, and edit an inactive scenario. | Correct readiness and current/stale behavior. Inactive edits and the dollar view leave matching results eligible. |
| `recover-and-reopen` | Start from a complete batch. Exercise visible Cancel at a computation yield point. Inject a controlled error during staged capture, guarded writes or post-write verification, then save, reopen and retry. | Application settings, protection, prior batch and selection restored. Only a complete matching retry appears current. |
| `format-and-boundary` | Managed download and local save, macro denied, non-target workbook, supported locale, marked 1904 date-system fixture, Letter and A4. | Actual download trust behavior, base projection usable, comparable report values, six pages, no repair or unrelated workbook edits. |

Planning preflight recalculated `extended-125` and `full-0-125` in Calc using seed 2026 and a 100-pair setting. Both accepted the intended horizons and the risk settings. The independent deterministic ledger matched 8,640 and 12,000 annual values respectively within the existing one-cent tolerance. Native Excel execution and timing remain D1 and D3 evidence.

### Prepare and record a benchmark

Run from the repository root with the [maintainer dependencies](../../tests/README.md). Use the `.xlsm` and `SHA256SUMS.txt` from the same [published release](https://github.com/Jon-275/financial-model/releases/tag/v1.0.0), and fetch its Git tag if your clone lacks it. The fixture command accepts all three benchmark cases and either pair count. The independent reference for an edited horizon also requires LibreOffice Calc.

```sh
release_sha256=$(awk '$2=="retirement-model.xlsm" {print $1}' scratch/SHA256SUMS.txt)
fixture_dir=scratch/benchmarks/example-95-100
python3 tests/desktop_benchmark.py prepare scratch/retirement-model.xlsm "$fixture_dir" \
  --revision v1.0.0 --sha256 "$release_sha256" --fixture example-95 --pairs 100
python3 tests/desktop_benchmark.py reference "$fixture_dir"
cp "$fixture_dir/desktop-input.xlsm" "$fixture_dir/mac-01.xlsm"
```

Open only the run copy in desktop Excel and run **Run market risk**. Measure the run separately from file open and save. After **Risk results are current**, save the copy in `.xlsm` format. Keep the original `desktop-input.xlsm` in the directory. Record the observed timing and environment in a separate JSON file. Use `cold` for the first run in a fresh Excel process, or `warm` after the process has already run the workbook. Count repeats within each phase.

```json
{
  "platform": "mac",
  "phase": "warm",
  "repeat": 1,
  "seconds": 120.0,
  "excel_build": "16.115.26110837",
  "excel_channel": "not observed",
  "excel_bitness": 64,
  "os_build": "macOS 26.7.3 (25G342)",
  "architecture": "arm64",
  "machine": "MacBook Pro, Apple M5 Pro, 24 GB",
  "distribution": "local copy of fictional release",
  "macro_policy": "per-file prompt",
  "timing_method": "launch to visible completion"
}
```

The JSON shows the field format with an illustrative duration. Record the actual measurement and method. Run both commands below to save and replay the evidence.

```sh
python3 tests/desktop_evidence.py record "$fixture_dir" mac-01.xlsm scratch/timing.json \
  --output evidence-mac-01.json
python3 tests/desktop_evidence.py check "$fixture_dir/evidence-mac-01.json"
```

`DesktopEvidence/v1` binds the source, fixture, desktop input, reference and Excel capture by hash. The checker requires a fresh batch and verifies the saved workbook's VBA source and library IDs against the pinned input. It replays the independent comparison and rejects missing checks, mismatched inputs, and duplicate saved batch timestamps for the same candidate, fixture and machine, even when the two save files have different hashes. Timing and environment fields describe the observed run. Keep evidence bundles local because Excel can add personal document metadata on save. Share sanitized summaries and hashes in PRs. The D3 release record also requires the hands-on compatibility, recovery and print checks below. [Evidence checker](../../tests/desktop_evidence.py)

### November 24, 2026 baseline

| Platform | Fixture | Observation | Reconciliation |
| --- | --- | --- | --- |
| MacBook Pro, Apple M5 Pro, 24 GB, macOS 26.7.3 (25G342), Excel 16.115.26110837 | `v1.0.0`, `example-95`, seed 2026, 100 pairs | One timed warm repeat, **135 seconds** from `RunMarketRisk` invocation to the workbook's saved batch timestamp. Excel showed current results and saved the `.xlsm`. | **233 outcomes**, with 80,454 compared fields and zero differences. The runner reported 46 verified fixed-result reuses. |
| Windows | Desktop Excel | Unverified | No Windows capture |

An earlier untimed run on the same Mac also reconciled. The timed value covers computation through the captured timestamp. Final verification and file save follow that timestamp. The Mac timing matrix appears below. Native cancellation and report exercises remain part of D3.

### November 25, 2026 Mac matrix

The Mac workload matrix has one cold and three warm runs for each scenario and pair count. Each cold run starts in a fresh Excel process. The warm repeats use the same process and a prior run of the same workload. The first standard-case warm repeat opened a fresh copy of the fixture in that process. Subsequent warm repeats keep the workbook open with the previous published batch. Inputs and the runner are fixed, and each run starts with only the benchmark workbook open. Two extended-horizon 100-pair attempts were discarded when unrelated workbook prompts interrupted capture. Cold retries start in a fresh Excel process.

Timing uses a monotonic clock from starting the synchronous Excel command for `RunMarketRisk` through its return, including final verification and restoration. Workbook opening, saving and independent replay take place outside that interval. The November 24 observation ends at the batch timestamp inside the macro and stays separate. Peak RSS is the maximum of two-second process samples where collected. Raw captures stay local.

**Mac matrix coverage is 24 of 24 independently reconciled runs.** Windows remains unverified.

The fictional `v1.0.0` workbook at `6573e1e4bd86eb99bbce6678c6fdf9f575c2b681` uses seed 2026. The test Mac was a MacBook Pro with Apple M5 Pro, 24 GB RAM, macOS 26.7.3 (25G342), and 64-bit Excel 16.115.2 (16.115.26112215). The Excel update channel was unobserved. Process samples identify the Excel build for the November 25 matrix and profiles.

| Fixture | Pairs | Cold (s) | Warm 1 (s) | Warm 2 (s) | Warm 3 (s) | Warm median (s) |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `example-95` | 100 | 143.965 | 130.199 | 161.823 | 135.364 | 135.364 |
| `example-95` | 1,000 | 1,108.198 | 1,081.553 | 1,109.471 | 1,146.757 | 1,109.471 |
| `extended-125` | 100 | 201.229 | 204.682 | 188.917 | 180.185 | 188.917 |
| `extended-125` | 1,000 | 1,575.026 | 1,644.752 | 1,629.388 | 1,608.097 | 1,629.388 |
| `full-0-125` | 100 | 274.019 | 298.783 | 271.708 | 259.099 | 271.708 |
| `full-0-125` | 1,000 | 2,063.465 | 2,415.751 | 2,508.750 | 2,549.974 | 2,508.750 |

Each capture passed freshness, pinned runner identity and independent reconciliation. The full matrix covers 27,192 outcomes and 8,022,096 field comparisons across the repeated runs. Field counts include trial, balance, ranked path ID and funding ledger values.

| Fixture and pairs | Repeat | Outcomes | Fields checked | Peak sampled RSS (MiB) | Capture SHA-256 |
| --- | --- | ---: | ---: | ---: | --- |
| `example-95`, 100 | cold 1 | 233 | 80,454 | — | `6bfdf0872eb40dc75bfb89c31ff33f7883b282645f1f018ff5fa74c839a1952e` |
| `example-95`, 100 | warm 1 | 233 | 80,454 | 993 | `24d3e6bc620116e1486c32f8cbfb894c9ad686eb45dc9c47650277bc559dd916` |
| `example-95`, 100 | warm 2 | 233 | 80,454 | 1,213 | `e3f3e9145429d54e36d9e66049645ff2d2185b9a4f8d50f07608a3847a25718f` |
| `example-95`, 100 | warm 3 | 233 | 80,454 | 970 | `568380daa0aa0821641bbe1702ae0af1418339c738620e3ca7fb7817a76b19cb` |
| `example-95`, 1,000 | cold 1 | 2,033 | 588,054 | 1,805 | `b36a3a839de196af00274b62b565533d503a9838254264bea70cd5f8e08b7e4a` |
| `example-95`, 1,000 | warm 1 | 2,033 | 588,054 | 2,480 | `ae46812ae9959e53b95fedf40a1e8219b8c572527a30cb0431b97707b6487494` |
| `example-95`, 1,000 | warm 2 | 2,033 | 588,054 | 2,722 | `c4ac4c09e35672503211f0e21536efc65f17e21a04f3ee43e10fb301c0805f30` |
| `example-95`, 1,000 | warm 3 | 2,033 | 588,054 | 2,451 | `071b7f57a78c8ca934ab85c30c549ae27ff2753720ac736d390dee76a998dfc9` |
| `extended-125`, 100 | cold 1 | 233 | 80,454 | 1,237 | `1c31e9b0daba118a17267a9f6007a2829571fcac763e7ffa6bf2156726b7bc80` |
| `extended-125`, 100 | warm 1 | 233 | 80,454 | 1,558 | `3499c2cd62e27e863d7f8f2eab70cd2d477bc6b2117a71fbe0f49c5e4eb6db68` |
| `extended-125`, 100 | warm 2 | 233 | 80,454 | 1,502 | `8d69ea56a4ba028219a9a4b73b8bbcae222db167aaa350728ac0539eb053c853` |
| `extended-125`, 100 | warm 3 | 233 | 80,454 | 1,608 | `84555000918c69878d8e3ad38d0c6c6199e748b47cb7c30cedaa2abab3c0d8f7` |
| `extended-125`, 1,000 | cold 1 | 2,033 | 588,054 | 2,505 | `6987152e765a09154c5d1792f22aa063a599e5f60f9c3feeb38aa6dc14fe4c5e` |
| `extended-125`, 1,000 | warm 1 | 2,033 | 588,054 | 2,866 | `b929d75def00265a68a9392581432adac74e87e20a28e1831b6c5d301b50f36f` |
| `extended-125`, 1,000 | warm 2 | 2,033 | 588,054 | 2,728 | `c25cfdc661afd419d1589f4eeed3fd484c06487432de232a011a7c833f38e531` |
| `extended-125`, 1,000 | warm 3 | 2,033 | 588,054 | 2,748 | `2f62d38a44bebda4b7279a8a0347c782fd7b95ef49c1b473985a084d9fd5a674` |
| `full-0-125`, 100 | cold 1 | 233 | 80,454 | 1,446 | `3187571f73ac3d36bb8da2534a3a98b48f9da526b0d655092757b7cbdfbd5c8f` |
| `full-0-125`, 100 | warm 1 | 233 | 80,454 | 1,677 | `821e6f93786d585abca2613214485786f54ce63e769a706e975274c35ad6b097` |
| `full-0-125`, 100 | warm 2 | 233 | 80,454 | 1,697 | `60dade3311bccbf995bcb97b97002471a8e21e17cc846a790e4e0671bb68dce4` |
| `full-0-125`, 100 | warm 3 | 233 | 80,454 | 1,719 | `c1749b4d45121ab4fa0b3e77c314df5a4b9eea2322612c6132ad2ca740aac2a3` |
| `full-0-125`, 1,000 | cold 1 | 2,033 | 588,054 | 3,112 | `92c8739b1b2555966e035eb634062e4479465a10604ca0bd7a9143dc467163aa` |
| `full-0-125`, 1,000 | warm 1 | 2,033 | 588,054 | 3,571 | `0630091ddeaea2facfb911410ac0a4ea420a55d460a8691501a6a843951e70ca` |
| `full-0-125`, 1,000 | warm 2 | 2,033 | 588,054 | 3,763 | `3d992a6210495f71a841bb715198d5f4a43209a891cb2e798bf1c716de03bc52` |
| `full-0-125`, 1,000 | warm 3 | 2,033 | 588,054 | 3,496 | `ce2110230e8f83cd9e4769e9d5a938c789157b709173bf11e69ae9821a2906a3` |

Run the existing comparison from the repository root against saved captures. Replace the paths below with the fixture artifacts from the manifest. Use `--expect-reused` only when the independent case supplies that expected count.

```sh
python3 tests/native_excel_acceptance.py "$reference" "$excel_saved_copy"
python3 tests/native_excel_acceptance.py "$before" "$after" --check-restoration
```

## D3. Execute on supported desktop builds

Record Microsoft 365 desktop Excel on Apple Silicon and 64-bit Windows first. Name the supported update channels and versions before release. Add other architectures or Windows/macOS channels only when their own test evidence exists. Exercise the public download under normal policies. In E1b, repeat acceptance on E1a's signed file under its target distribution policy. Preserve application security settings in either lane. The released v1.0.0 workbook has prior Mac compatibility evidence. **Windows remains unverified.** [Maintainer walkthrough](../../tests/README.md#verify-a-changed-workbook-in-excel) · [Windows macro behavior](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)

1. Open a disposable copy through the supported user route, record any security or repair prompts, and verify the normal path to **Inputs** and **Summary**. Check protected formulas and editable inputs. Edit the end age and an assumption from an active scenario. Verify that the base Summary and Annual Detail update and saved risk becomes stale. Undo, switch scenarios, change only dollar view and verify status after each step.
2. Run on the worksheet button and verify progress, visible Cancel, failure status, retry, last-audit fields, save/reopen and the inspection selector. In a separate exercise with a second disposable workbook, record application-wide recalculation and verify that the runner makes no unintended edits or settings changes. Confirm the base projection works when macros are blocked. Repeat essential button and print checks interactively if app automation is used. Dispatch only reviewed commits with synthetic inputs to a licensed managed test device or private interactive host.
3. Compare the saved results with the independent workbook output. Check the fixed and flexible counts, annual ledgers, ranked examples, preserved report choices and all stress cases. Inspect raw desktop exports in both dollar views on Letter and A4, recording page count, layouts, extracted values and both decision pages. Use `example_pdf_acceptance.py` for the prepared public fictional release PDF with its metadata and structure requirements. [PDF comparison](../../tests/example_pdf_acceptance.py)

## R1 and R2. Measure and remove the dominant cost

Time the native macro independently of open, fixture preparation and save. For `example-95`, `extended-125` and `full-0-125`, collect 100- and 1,000-pair timings with one cold run and three warm repeats per supported platform on fixed hardware and settings. Close other workbooks during benchmark runs. Record the individual values and warm median. Report percentiles only with a sample that supports them. Capture peak memory where safely measurable and record progress and cancel latency separately. Budget for the existing long-running fixtures as an integration lane, with full grids on candidates touching execution and smaller smoke coverage on unrelated changes.

Instrument preflight/snapshot, trial preparation, recalculation, verified capture, fallback, ranking, publish and restore. Count executed and reused trials, fallback reasons, and all 33 stresses. Store timing and counters in test evidence outside plan identity and financial outputs. Measure the primary full calculation in `FinishTrialCalculation`, the year-ordered retry and workbook ranking separately. Inspect formula dependencies with `native_order_acceptance.py` before trying to narrow the calculation scope. [Native runner](../../tools/vba/RunMarketRisk.bas) · [Order acceptance](../../tests/native_order_acceptance.py) · [Excel calculation performance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance)

### Capture a stage profile

Prepare a pinned fictional fixture and its independent reference using the commands in D1. The profiling tool derives a second `.xlsm` in that directory and replaces only its VBA project. It checks every other package part against the original. Use a separate working copy of `desktop-profile-input.xlsm` for Excel. Keep this diagnostic workbook local to the benchmark environment.

```sh
python3 tests/desktop_stage_profile.py prepare "$fixture_dir"
```

1. Open the working copy of the diagnostic workbook as the only workbook in desktop Excel. Measure the synchronous `RunMarketRisk` invocation from dispatch through return with an independent monotonic clock. Use the run timing metadata from D1 and record the diagnostic distribution.
2. Call `ReadMarketRiskProfile` in the same workbook after that invocation returns. Capture the returned text using Excel's AppleScript `run VB Macro` command on Mac or `Application.Run` on Windows. Save that text locally as `profile.txt`. Save the workbook with a distinct `.xlsm` filename in the fixture directory and confirm the report shows current results before recording.
3. Record and replay the evidence. Keep the original diagnostic input untouched. Use a distinct capture, timing file, text payload and evidence name for each repeat.

```sh
python3 tests/desktop_stage_profile.py record "$fixture_dir" capture-cold-01.xlsm \
  "$fixture_dir/timing.json" "$fixture_dir/profile.txt" --output profile-evidence-cold-01.json
python3 tests/desktop_stage_profile.py check "$fixture_dir"/profile-evidence-*.json
```

Each transition assigns elapsed time to one stage. Preflight covers input checks and snapshotting Excel settings and protections. Preparation covers trial selection and staging trial arrays. Primary calculation times the full Excel calculation and its existing wait. Verified capture covers the source-range calculation, reads and reconciliations. Ordered calculation covers the retry caused by a balance check or an unexpectedly repeated ending. Ranking covers sorting, selection and staging of the three ranked ledgers. Their calculations and capture use the same timing stages as other trials. Reuse verification covers the fixed-path proof calculation and its paired-flexible checks. Progress covers status-bar updates and the existing cancellation yield. Publish covers the prior-batch backups, guarded writes and workbook verification. Restore includes returning to the inspection selection, cleanup and rollback where applicable.

The profile records computed and reused trial totals, all 33 stresses, three ranked ledgers, full calculation entries and ordered retries by reason. It binds the exact diagnostic fixture and capture to a fresh published batch and the independent reference. A missing clock, cancellation or any runner failure prevents a profile from counting as completed timing evidence. Verify interrupted runs separately with the restoration comparison under D3. Report the diagnostic stage totals with the external wall times and compare the latter with matched runs of the release runner before attributing a performance difference.

The diagnostic runner uses a monotonic raw clock on 64-bit Mac and requests the performance counter on 64-bit Windows. Leave Excel and system security policies in place. Record the diagnostics as unavailable if platform policy blocks the clock. Windows execution remains unverified. [Apple monotonic clock](https://developer.apple.com/documentation/kernel/1646199-mach_continuous_time) · [Windows performance counter](https://learn.microsoft.com/en-us/windows/win32/sysinfo/acquiring-high-resolution-time-stamps)

### November 25, 2026 Mac stage profiles

The diagnostic runner at `042d6a7593fa14485e22eef4e1967d844856d33e` profiled the pinned `v1.0.0` fictional `example-95` fixture with 100 pairs and seed 2026. The Mac and Excel build match the release matrix above. One cold run started in a fresh Excel process. The three warm repeats used that process and kept the workbook open. The timings are for the diagnostic fixture. Use the 24-run release matrix above for release-run timings.

| Run | Macro wall time (s) | Total time assigned to stages (s) |
| --- | ---: | ---: |
| Cold | 133.941 | 133.757 |
| Warm 1 | 150.119 | 149.892 |
| Warm 2 | 154.158 | 153.972 |
| Warm 3 | 129.035 | 128.858 |

Each saved batch independently reconciled 233 outcomes and 80,454 fields, with 187 computed outcomes, 46 verified reuses, 33 stresses, three ranked histories and zero ordered retries. The warm median wall time was 150.119 seconds. Across the three warm runs, median full calculation was 55.029 seconds, median trial preparation was 38.701 seconds and median reuse verification was 38.947 seconds. These are medians for individual stages. Preparation and reuse verification each include multiple operations, so use finer probes to choose what to optimize.

On disposable copies of the saved cold batch, an injected cancellation at the second progress yield and an injected failure after the publication writes each preserved 595,814 stored result fields and restored inspection, application controls and worksheet protection. A profiling-clock fault yielded unavailable diagnostics and a fresh financial batch that reconciled. All three subsequent retries published fresh current batches and reconciled all 233 outcomes. The cancellation exercise invoked the runner directly. D3 still covers the visible Cancel button, Esc and print behavior.

The diagnostic workbook SHA-256 is `19fdee2f6b84cf0965622b66af4395aa1eaf548ee033a0cc537b335cac4f16e3`. Capture hashes and follow-up profiling are recorded with the reviewed source. The raw workbooks stay local. Windows remains unverified. [PR #37 evidence](https://github.com/Jon-275/financial-model/pull/37)

### Guarded Mac calculation experiment

Use separate copies of the `full-0-125` fixture and its independent reference. Select 100 pairs for a smoke test or 1,000 pairs for the full-scale workload. Match the count and manifest across the comparison arms. Prepare the full-first control with the existing stage profile command. Prepare the natural candidate and the fault-injection copy from untouched copies of the same fixture. The candidate attempts Excel's normal calculation first. A balance mismatch or a suspect repeated ending triggers the original full calculation, then the original ordered recalculation if needed. Recheck the selector, assumptions and closing balances on every attempt. Hard calculation or input errors and invalid accepted capture records stop the run. The candidate's `primary_calculation` bucket includes normal and full fallback time.

```sh
python3 tests/desktop_guarded_calculation.py prepare "$candidate_dir"
python3 tests/desktop_guarded_calculation.py prepare "$canary_dir" --mode canary
```

Open a working copy of the prepared workbook as the sole open workbook. Use a fresh Excel process for each cold control or candidate run. Measure through the return of `RunMarketRisk`. Before saving, read `ReadMarketRiskProfile` and `ReadGuardedCalculationCounts` from the same workbook and save each returned string separately. Use the D1 timing format and keep the verified template separate from the fresh capture.

```sh
python3 tests/desktop_guarded_calculation.py record "$candidate_dir" capture-cold-01.xlsm \
  "$candidate_dir/timing.json" "$candidate_dir/profile.txt" "$candidate_dir/counts.txt" \
  "$candidate_dir/guarded-evidence-cold-01.json"
python3 tests/desktop_guarded_calculation.py check "$candidate_dir/guarded-evidence-cold-01.json"
```

For the canary, supply `--mode canary` when recording. It corrupts a captured value in memory on the first fictional trial after both normal and full calculation. A passing result requires the ordered retry, a fresh saved batch, and independent reconciliation. Measure performance with the natural candidate on the same fixture, independent reference and desktop build as the control. Use distinct saved captures and require a newer batch serial for every warm run in the same session. Trial-specific reconciliation, fallback counts and the saved capture are required for each recorded run. This is a test-only Mac path. Windows remains unverified.

Replay the complete comparison before reporting the timing pairs. Pass cold and warm evidence in that order for each arm. The comparison writes a summary of the reconciled captures, timings and fallback counts. It reports canary evidence separately from the performance observations.

```sh
python3 tests/desktop_guarded_calculation.py compare \
  --control "$control_dir/profile-evidence-cold-01.json" "$control_dir/profile-evidence-warm-01.json" \
  --candidate "$candidate_dir/guarded-evidence-cold-01.json" "$candidate_dir/guarded-evidence-warm-01.json" \
  --canary "$canary_dir/guarded-evidence-cold-01.json" --output guarded-comparison.json
```

### November 25, 2026 guarded Mac observations

Both arms used the pinned fictional `full-0-125` workbook with 100 pairs and seed 2026, Excel for Mac 16.115.26112215 and the same independent reference. Each arm used one fresh Excel process followed by one repeat in the open workbook. Wall time covers synchronous macro dispatch through return.

| Run | Full-first (s) | Smart-first (s) | Full fallbacks after smart | Ordered retries, full-first / smart-first |
| --- | ---: | ---: | ---: | ---: |
| Cold | 256.856 | 188.764 | 1 | 1 / 0 |
| Warm | 275.567 | 207.077 | 5 | 5 / 0 |

Each run independently reconciled 233 outcomes and 80,454 fields. Each smart-first run attempted 190 smart calculations and retained 46 verified reuses. These are initial paired observations. A production decision still requires repeated runs, the 1,000-pair workload, checks for internally consistent stale captures, and cancellation and publication recovery on the candidate. Windows remains unverified.

For a replicated warm comparison, start each arm in its own fresh Excel process and keep its one diagnostic workbook open for one cold run and three warm repeats. Save and independently reconcile each run before dispatching its successor. Alongside each evidence file, save a timing receipt named `raw-cold-01.json` or `raw-warm-01.json` through `raw-warm-03.json`. Record the Excel process ID as a decimal string in `excel_pid`, the SHA-256 of the open workbook's saved file before the macro in `starting_sha256`, the original run metadata in `timing`, and the native profile and fallback payloads in `profile` and `counts`. Use an empty string for the control's fallback payload.

```sh
python3 tests/desktop_guarded_calculation.py compare-warm \
  --control "$control"/profile-evidence-{cold-01,warm-01,warm-02,warm-03}.json \
  --candidate "$candidate"/guarded-evidence-{cold-01,warm-01,warm-02,warm-03}.json \
  --output "$scratch/matched-warm.json"
```

The comparison replays all eight saves, checks each warm run against the preceding capture in its arm, and calculates medians from the three warm runs. Keep the separate recovery canary as part of the release evidence.

The separate canary rejected both injected captures on the first trial, reached the ordered retry, and independently reconciled the fresh saved batch across the same 233 outcomes and 80,454 fields. It recorded two injected mismatches, two full fallbacks overall, and one ordered balance retry. Its timing is excluded from the paired comparison.

### Replicated 100-pair warm Mac comparison

On November 26, 2026, the full-first control and the guarded candidate each ran the pinned `full-0-125` fixture with 100 pairs and seed 2026. The same M5 Pro ran Excel for Mac 16.115.26112215 on macOS 26.7.3 (25G342). Each arm started in a fresh Excel process and kept its one diagnostic workbook open across one cold run and three warm repeats. The replay linked each warm timing to the preceding save in its arm. Each of the eight saved batches reconciled 233 outcomes and 80,454 fields, with 46 verified flexible reuses and three ranked ledgers.

| Run | Full-first wall time (s) | Guarded wall time (s) | Full-first ordered retries | Guarded full fallbacks | Guarded ordered retries |
| --- | ---: | ---: | ---: | ---: | ---: |
| Cold | 251.650 | 197.415 | 1 | 4 | 1 |
| Warm 1 | 273.308 | 212.114 | 2 | 5 | 0 |
| Warm 2 | 276.899 | 204.315 | 8 | 11 | 0 |
| Warm 3 | 263.036 | 208.584 | 9 | 13 | 0 |
| Warm median | 273.308 | 208.584 | | | |

The guarded warm median was 23.7% lower for this 100-pair fixture. The matching tool reverified all eight evidence records and native run receipts. These measurements apply to the existing diagnostic overlay. Replicated 1,000-pair timing and production repair validation remain open. Windows remains unverified.

Independent reference SHA-256 `ce8b2f723f8eeff9bdba8637431f28a73fb94b85f6a743c62f6680770132cf7d`. Full-first diagnostic input SHA-256 `59cfcbd256ce5870fc2069f904236978c0c8da17874891f45adfbf8f61f71164`. Guarded input SHA-256 `82dd6338e98add64085cb23f25e93dd9216089f3f5b298bb38a0fb39705b92a5`. The aggregate comparison SHA-256 is `278c4f843bac2fe020cbb056da1900b96368cd30abc2d48e1e42876e743e17e8`.

### Earlier full-horizon brokerage probes

Earlier full-first probes on this same fictional fixture observed brokerage closing mismatches at ages 112 and 118. At age 118 in trial 20, the live closing value exceeded the recomputation from opening balance, withdrawal, deposit and return by $102,062,082.42. Ordered recalculation recovered. An intentional cancellation after that trial preserved the previous 595,814 stored result fields. A large failed reconciliation requires calculation-order investigation. Keep the comparison and ordered retry in place while tracing the first stage at which the live closing and its operands disagree.

### Full-scale guarded Mac observations

The `full-0-125` 1,000-pair fixture used seed 2026, Excel for Mac 16.115.26112215 and macOS 26.7.3 on an M5 Pro. The profiled historical full-first cold run and the guarded cold run used the same fixture and independent reference. Wall time ends when the synchronous macro returns.

| Run | Wall time (s) | Full calculations | Ordered balance retries |
| --- | ---: | ---: | ---: |
| Historical full-first cold | 1,989.144 | 1,535 | 40 |
| Guarded cold | 1,523.699 | 40 fallbacks | 3 |
| Guarded warm | 1,604.108 | 41 fallbacks | 0 |

Each saved batch independently reconciled 2,033 outcomes and 588,054 fields, including three full ranked ledgers. The guarded cold run made 1,535 smart attempts and retained 501 verified reuses. The cold comparison is a 23.4% wall-time reduction. The profiled warm full-first control and a replicated warm median remain release gates. The existing 100-pair canary covers forced recovery, and the full-scale guarded cold run exercised ordered recovery on naturally observed mismatches. Windows remains unverified.

The source workbook and runner are pinned to `6573e1e4bd86eb99bbce6678c6fdf9f575c2b681`. The diagnostics and replay at `0c9c35c06045bdcce278a85b14b29cc82c0d1849` reproduced both diagnostic input hashes. The independent reference SHA-256 is `858c4723327a08c72c844c1b25c0150f75e6aec9afd69a7d209d8762267e2688`. The raw workbooks stay local.

| Run | Diagnostic input SHA-256 | Saved capture SHA-256 |
| --- | --- | --- |
| Historical full-first | `4f0a3ca409e7aa477819c2495683d90efa60e16e8cd6aff1eade931953f77d5e` | `d566f0aedcfb4270ea10a5dd4a708dd10c09c988b9ca46bd8c777856f8c53f13` |
| Guarded cold | `9ee1e57655b01140b2d6f74eec234ee90d67bf0c42b1374b8fa1230ed5315e77` | `aef39bd27e0e9596e68ecdabda8cd19b14eaad8ff4f5f0cd8961980e02b4e276` |
| Guarded warm | `9ee1e57655b01140b2d6f74eec234ee90d67bf0c42b1374b8fa1230ed5315e77` | `342050dab9f9040b7d58777854c8a74f3e9d8e390775de798ef0ebf143179a0a` |

The warm run used the same open guarded workbook after the verified cold capture. Its saved batch timestamp matches the warm profile and follows the cold batch timestamp. The warm run retained 501 verified reuses.

### Brokerage calculation trace on Mac

On November 26, 2026 the adaptive trace reproduced the first closing-balance rejection at trial 10, age 114, in `Trial_Checks!N119`. After writing selector 10 and before calculating, the observed brokerage cells still described trial 9. When `Application.CalculateFull` returned, the opening balance and return matched trial 10, while `Trial_Projection!BC119` and `V119` still matched trial 9's closing balance and withdrawal. The deposit retained its previous snapshot value and was within one cent of both independent paths. Those values persisted through capture-range recalculation and the closing check. The runner's stored guard residual and the subsequent cell read agreed at −$85,426,746.79. The existing ordered recalculation restored a consistent trial 10 and a zero residual.

The trace identifies the full-calculation return as the first sampled point with mixed-trial values. The Excel-internal dependency or scheduling cause remains open. Both frozen and reread residuals and the independent trial replay support the observation. The separate diagnostic stopped after fixed trial 81 and independently verified the prior inspection, inactive funding detail and all 595,814 stored result fields. Its reads were excluded from performance measurements. Trace input SHA-256 `50bb370cebe964408ac1467a158f77a8d655840bdfde212c3145bae8afb67865`. Saved capture SHA-256 `ccad2bad445d4bdda4eef416dbcfc2067b806dbe8a4cede1c05a49a4a5bbe6bd`.

#### Brokerage funding solver

A second disposable 1,000-pair trace reproduced stale values after switching from fixed trial 6 to 7. The age-121 opening brokerage balance and return matched trial 7 when full calculation returned. The withdrawal and closing value still matched trial 6. At the first brokerage closing rejection, the funding solver's selected gross draw also matched trial 6. The selected withdrawal in `Trial_Funding_Calc!H14884` exceeded its own formula evaluated with its live precedents by $658.21. This places a stale value inside the funding solver as well as in the projection.

| Cumulative recalculation at age 121 | Selected gross draw ($) | Brokerage withdrawal ($) | Brokerage close ($) | Brokerage check residual ($) |
| --- | ---: | ---: | ---: | ---: |
| First brokerage rejection | 3,991,289.26 | 3,984,479.92 | 347,487,591.62 | −35,408,938.13 |
| Projection withdrawal and close | 3,991,289.26 | 3,984,479.92 | 382,896,529.76 | 0.00 |
| Selected solver withdrawal and projection | 3,991,289.26 | 3,983,821.71 | 382,897,338.32 | 0.00 |
| Localized inputs, whole same-year solver block and projection | 3,761,432.87 | 3,753,965.32 | 383,179,699.55 | 0.00 |
| Existing full ordered retry | 3,761,432.87 | 3,753,965.32 | 383,179,699.55 | 0.00 |

Projection-only recalculation produced a zero value in `Trial_Checks!N126` while the withdrawal still exceeded the independent result by $230,514.60. The complete same-year solver recalculation and the ordered retry both produced the independently replayed withdrawal and close to the cent. The diagnostic preserved the original rejection, performed the full ordered retry and cancelled before publishing. This establishes the limit of the single brokerage closing equation. A release repair still needs evidence that it prevents stale funding throughout the capture, preserves cancellation and publication recovery, and passes the matched warm gate.

The diagnostic stopped after fixed trial 81 and independently restored the inspection, inactive funding detail and 595,814 stored fields. It was excluded from timing. Input SHA-256 `bfedd28d0b4c77d843b92ed751837262782f8851e6fdd19a0f021ad5f0f23caf`. Saved capture SHA-256 `aae764dc8232d9348c63b6cb3bd594f66e8a9340ae65f4926e9e511995b85fd0`. The independent reference is the same as the full-scale benchmark above. Windows remains unverified.

#### Offline funding selection regression

In a separate fictional trace, trial 7 at age 121 retained trial 6's gross requested draw of $3,991,289.26141221. Recalculating the selected brokerage withdrawal and closing balance made the closing check read zero. The ordered recovery selected $3,761,432.87464338. The trace capture SHA-256 is `aae764dc8232d9348c63b6cb3bd594f66e8a9340ae65f4926e9e511995b85fd0`. Its independent reference SHA-256 is `858c4723327a08c72c844c1b25c0150f75e6aec9afd69a7d209d8762267e2688`.

Offline repricing using only trial 7's pinned income, account and tax inputs gave a $200,440.86300932 cash margin on the stale request and approximately zero on the recovered request. One cent below the stale request still funded the year with a $200,440.85428906 margin. One cent below the recovered request left a $0.00872003 deficit. A fresh funded quote below the requested draw by more than the solver's half-cent convergence tolerance refutes an earliest-funded selection within the same effective conversion policy. A missing witness leaves the selection open, including at zero, within numerical tolerance, or near an ACA discontinuity. Verify unfunded choices against the applicable grid, checking whether a funded candidate exists and otherwise checking best margin and the smaller-draw tie rule.

Run `PYTHONPATH=tools:tests python -m unittest test_risk_funding_witness -v` to reproduce the offline comparison. The test uses pinned fictional inputs and the existing Python funding solver. [Funding witness tests](../../tests/test_risk_funding_witness.py)

#### Native coherent-stale canary

The Mac canary intentionally reused trial 6's selected gross draw on trial 7's current age-121 inputs. After recalculating the selected brokerage withdrawal and close, `Trial_Checks!N126` read zero. Pricing one cent below that stale draw left a $200,440.85 cash surplus. Pricing one cent below the independently recovered draw left a $0.00872003 deficit under the same conversion policy. The native quotes agreed with the Python reference. The test restored the solver formula and sheet protection, reran the ordered calculation, then cancelled before publishing. The saved cancellation independently reconciled the restored inspection, inactive funding detail and 595,814 stored fields.

This is evidence for the funding quote check and ordered correction in the controlled case. Release validation still needs that check integrated before capture with coverage for conversion decisions, discontinuities and unfunded paths. Canary input SHA-256 `6b3ecaaffda894bd7aa2cc6be5dd0e8aacfc6f1e670573d509eced04429d115b`. Saved capture SHA-256 `fda900ce432c54359f2e9655c84818dc60621041062dca968e9c0f55b57d5f51`.

### Disposable native quote formulas

Run `PYTHONPATH=tools python tests/native_funding_witness_formulas.py 126` to emit the quote formulas for trial projection row 126 as JSON. Omitting the row emits every annual row from 6 through 130. Place the formulas at those coordinates on a separate test worksheet. The generator reuses the native projection's candidate draw, conversion, tax and ACA formulas. It links directly to the annual sources, including pre-funding cash `Trial_Cash_Flow!AB`, and computes its own withdrawal breakpoints.

Fill columns A, B and C with the claimed funding status, gross draw request and independently established conversion mode for that year. Keep the raw mode in the native trace. Mode 1 can produce a zero conversion when the current year has no eligible request. Leave C blank when the current policy branch is uncertain. An unfunded claim needs the applicable grid. A nonpositive or near-zero lower quote does not settle a claim, including when the quote crosses an ACA discontinuity. A fresh positive lower quote is a counterexample under the established policy.

Refresh and verify the current trial's annual sources before calculating D and G, then the breakpoints AM:AO, then the candidate economics H:AL and the result AP. Read the repriced margin in Z. The Mac canaries first priced both requests using the original solver range, then independently repriced the generated row 126. The generated lower quote gave a $0.00872003 deficit for the recovered draw and a $200,440.85428905 surplus for the stale draw. It also exercised zero draw, unsupported conversion mode and the unfunded disposition. The prior batch was preserved after cancellation. The generated-row capture SHA-256 is `d4b334b80f0f76bd00569ee44baf3d73bc67288a07f097be6c209ab3c6c7f6c9`. [Formula generator](../../tests/native_funding_witness_formulas.py) · [Structure tests](../../tests/test_native_funding_witness_formulas.py)

Prepare and verify a disposable full-span bank from the repository root.

```sh
proof_work=$(mktemp -d)
python tests/desktop_benchmark.py prepare workbook/retirement-model.xlsm "$proof_work/fixture" \
  --revision v1.0.0 --sha256 f5faea8318773c911800462d43488c5063dc20b9363f1121151454d5e8007af4 \
  --fixture full-0-125 --pairs 100
PYTHONPATH=tools python tests/native_funding_proof_package.py \
  "$proof_work/fixture/desktop-input.xlsm" "$proof_work/proof-bank.xlsm"
PYTHONPATH=tools python tests/native_funding_proof_package.py \
  "$proof_work/fixture/desktop-input.xlsm" "$proof_work/proof-bank.xlsm" --check
```

The builder pins the names and decompressed bytes of every package part, including the original VBA. ZIP timestamps, compression and member ordering may differ. It changes only `Trial_Checks` worksheet XML. The check verifies the exact generated formulas and all original package parts before the separate diagnostic VBA is layered in. [Package builder](../../tests/native_funding_proof_package.py)

The bank spans BA6:CP130. BA and BB read the selected solver status in AC and gross request in D. BC certifies numeric zero only when the current annual inputs preclude conversion, such as a zero request or an ineligible year. Unresolved years stay inconclusive. Calculate BA:BC, BD:BG, CM:CO, BH:CL and CP in that order after confirming the source year is current. BZ holds the independent margin. At projection row 126 the selected solver row is 14884. The generated worksheet still needs a combined native capture, recovery and full-batch comparison.

### Performance release gate

Take one optimization per PR. If recalculation dominates, evaluate scoped calculation using the proven dependency order. Keep primary full calculation for the baseline and retain it as a recovery path only after proving the scoped path, with the existing ordered retry available. If marshaling dominates, batch reads and writes at the same verified boundary. If ranking dominates, preserve exact tie breakers and revalidate the three complete funding ledgers. Compare the baseline and candidate on each platform where the change will be enabled. A recurring fallback has an investigated cause and a regression case. Narrowed calculation may become the default only after the regular and fallback paths independently reconcile, including cancel and failure recovery. Retain the baseline runner strategy for rollback while the candidate is under evaluation.

Set release latency budgets after collecting raw timings. Initial comparison targets are at least 25% better warm median for full-range 1,000 pairs, no more than 10% regression in any other fixture, and visible cancel acknowledgement within five seconds at normal yield points. Track median 100-pair time as the exploration outcome and require an explicit budget on each supported device before claiming the new fast path. Record the prior 8- and 33-minute figures as historical context. Investigate and correct the cause of every new reconciliation or recovery failure. Keep the candidate Mac-only if Windows evidence is absent and label the missing tests in the release decision.
