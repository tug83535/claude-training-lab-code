# Release Readiness Checklist

Use this checklist before any executive or company-wide demo.

## 1) Workbook readiness

- [ ] Demo workbook opens without repair warnings.
- [ ] Macros are enabled and trusted on presenter machine.
- [ ] Required demo sheets exist (`Assumptions`, `Checks`, `P&L - Monthly Trend`, etc.).
- [ ] Source sample files in `samples/` are untouched.

## 2) Feature readiness

- [ ] `BuildDemoCommandCenter` runs successfully.
- [ ] `RunDemoReconciliation` updates `Checks` with timestamped output.
- [ ] `GenerateDemoVarianceNarrative` creates/refreshes `Exec_Variance_Narrative`.
- [ ] `RunDemoWhatIfScenarios` creates/refreshes `Scenario_Compare`.
- [ ] `BuildDemoExecutiveBriefPack` creates `Exec_Brief` and PDF export where applicable.

## 3) Universal toolkit readiness

- [ ] `BuildCommandCenter` builds `UTL_CommandCenter`.
- [ ] Sanitizer preview and full run complete without runtime errors.
- [ ] Compare/consolidate flow creates output sheets.
- [ ] One-pager generation succeeds.

## 4) Python and SQL readiness

- [ ] `bash scripts/run_stage_smoke.sh` passes fully.
- [ ] Python utilities run with expected output files.
- [ ] SQL templates reviewed for environment-specific schema changes.

## 5) Branding and communications

- [ ] Arial font only in all visible report outputs.
- [ ] Brand palette matches `BRAND.md` values.
- [ ] Executive outputs contain clear timestamp and ownership.
- [ ] Video scripts align with current feature behavior.

## 6) Risk controls

- [ ] Presenter has a clean backup copy of workbook.
- [ ] Offline fallback plan exists (PDF exports and screenshots prepared).
- [ ] Recovery steps reviewed from `guides/troubleshooting-reference.md`.

## 7) Final sign-off

- [ ] Finance owner sign-off
- [ ] Reviewer sign-off
- [ ] Demo dry-run completed within target time
