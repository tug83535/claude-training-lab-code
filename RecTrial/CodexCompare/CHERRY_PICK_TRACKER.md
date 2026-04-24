# Cherry-Pick Tracker — Codex → Project A

**Source:** Comparison report at `COMPARISON_REPORT.md` (Section 8 + supporting sections).
**Purpose:** Track which ideas from the Codex build we're porting into Project A's universal toolkit + other areas. Each row lists what it is, where it lands, effort, risk, and current status.

**Started:** 2026-04-20
**Project A source of truth for VBA/Python:** `C:\Users\connor.atlee\RecTrial\`
**Repo (sync target after):** `C:\Users\connor.atlee\.claude\projects\claude-training-lab-code\`

---

## Conventions when porting

- **VBA headers:** every ported file gets Project A's header comment block (see top of `modUTL_Core.bas`). Codex's VBA has no headers — don't inherit that.
- **Brand colors:** use iPipeline brand RGBs (documented at top of `modUTL_Branding.bas`): Blue `RGB(11,71,121)`, Navy `RGB(17,46,81)`, Innovation Blue `RGB(75,155,203)`, Lime Green `RGB(191,241,140)`, Aqua `RGB(43,204,211)`, Arctic White `RGB(249,249,249)`, Charcoal `RGB(22,22,22)`. No inline off-brand `RGB(...)` values.
- **Path A pattern:** any new sub that shows dialogs needs a `DirectorXxx` silent wrapper.
- **Naming:** sheets created by toolkit prefix with `UTL_` (e.g., `UTL_RunReceipt`, `UTL_QualityScorecard`, `UTL_Intelligence`).
- **Option Explicit:** required on every `.bas`.

---

## Item tracker

| ID | Item | Type | Destination | Effort | Risk | Status |
|----|------|------|-------------|--------|------|--------|
| **K** | Margin-threshold narrative labels (`MarginVerdict` + `AppendMarginVerdictRow`) | VBA (demo) | `RecTrial\DemoVBA\modWhatIf_v2.1.bas` | S | None | 🟢 Done 2026-04-20 |
| **B** | `CreateRunReceiptSheet` helper | VBA (universal) | `RecTrial\UniversalToolkit\vba\modUTL_Audit.bas` | S | Low (sheet name `UTL_RunReceipt` must be excluded from compare tools) | 🟢 Done 2026-04-20 |
| **C** | `UTL_DetectHeaderRow` helper | VBA (universal) | `RecTrial\UniversalToolkit\vba\modUTL_Core.bas` | S | None | 🟢 Done 2026-04-20 |
| **A** | `modUTL_Intelligence.bas` — MaterialityClassifier / ExceptionNarratives / DataQualityScorecard (3 universal subs) | VBA (universal, new module) | `RecTrial\UniversalToolkit\vba\modUTL_Intelligence.bas` | M | Overlap with `modDataQuality_v2.1` verified none | 🟢 Done 2026-04-21 |
| **D** | `UTL_QuickRowCompareCount` + `BuildRowHashMap` helpers (Scripting.Dictionary pipe-delim keys) | VBA (universal) | `RecTrial\UniversalToolkit\vba\modUTL_Compare.bas` | S | None — additive only | 🟢 Done 2026-04-21 |
| **E1** | `profile_workbook.py` (stdlib-only workbook inventory) | Python (zero-install) | `RecTrial\UniversalToolkit\python\ZeroInstall\` | S | Stdlib-only verified | 🟢 Done 2026-04-21 |
| **E2** | `sanitize_dataset.py` (stdlib CSV cleanup) | Python (zero-install) | Same | S | Stdlib-only verified | 🟢 Done 2026-04-21 |
| **E3** | `compare_workbooks.py` (stdlib xlsx diff → CSV) | Python (zero-install) | Same | S | Stdlib-only verified | 🟢 Done 2026-04-21 |
| **E4** | `build_exec_summary.py` (CSV → markdown summary with talking points) | Python (zero-install) | Same | S | Stdlib-only verified | 🟢 Done 2026-04-21 |
| **F** | `variance_classifier.py` — Actual vs Baseline labels | Python (zero-install) | `RecTrial\UniversalToolkit\python\ZeroInstall\` | S | None | 🟢 Done 2026-04-21 |
| **G** | `scenario_runner.py` — stdlib % shocks to a metric column | Python (zero-install) | Same | S | None | 🟢 Done 2026-04-21 |
| **H** | `sheets_to_csv.py` — extract named sheets to CSVs (stdlib) | Python (zero-install) | Same, renamed from Codex's `pnl_data_extract.py` | S | None | 🟢 Done 2026-04-21 |
| **I/J** | `build_talking_points()` + `--talking-points` CLI flag | Python (existing module) | `RecTrial\UniversalToolkit\python\word_report.py` | S–M | None — opt-in, default off | 🟢 Done 2026-04-21 |
| **L** | Dual-logging pattern (local `VBA_AuditLog` sheet + universal logger) | VBA (demo + universal) | `RecTrial\DemoVBA\modLogger_v2.1.bas` + pattern doc | S (per module) | Moderate — integrate carefully with existing `LogAction` signature (13 prior bugs) | ⬜ Planned (Batch 4) |
| **9** | Top-level `CONSTRAINTS.md` | Docs | `claude-training-lab-code\CONSTRAINTS.md` | S | None | ⬜ Planned (Batch 5) |
| **10** | Top-level `BRAND.md` | Docs | `claude-training-lab-code\BRAND.md` | S | None | ⬜ Planned (Batch 5) |
| **4** | Release-readiness checklist | Docs (guide) | `RecTrial\Guide\RELEASE_READINESS_CHECKLIST.md` | S | None | ⬜ Planned (Batch 5) |
| **5** | User-facing `TROUBLESHOOTING.md` | Docs (guide) | `RecTrial\Guide\TROUBLESHOOTING.md` | S–M | None | ⬜ Planned (Batch 5) |

## Items NOT being ported (and why)

| ID | Item | Why skipped |
|----|------|-------------|
| Section 4 #16 | 456-LOC stage2_smoke_check.py | Overkill; scope creep risk. Could revisit later with a smaller, targeted check. |
| Section 4 #17 | 8-column RunLog schema replacement | Touches existing logger (13 LogAction bugs' worth of scar tissue). Defer; consider only if dual-logging (L) doesn't cover the gap. |
| Tier 2 #11 | Workbook-mapping CoPilot prompt addition | Parked for later — not a toolkit code add. |
| Tier 2 #12 | Git branch push quickstart | Not urgent — no coworker contributions planned yet. |
| Tier 2 #13 | Architecture overview one-pager | Archive/docs already has similar; nice to have, not impactful. |
| Tier 2 #14 | Makefile | Requires make on Windows; PowerShell equivalent if desired later. |
| Tier 2 #15 | GitHub Actions CI | Out of scope for current work. |
| Tier 2 #18 | SQL extract templates | Parked for Video 4 prep if needed. |
| Tier 2 #20 | Top-level `STARTER_PROMPT.md` | Already have one for AP_CodexVersion side repo; adapt later if useful. |
| Tier 3 #24 | Named error-label style | Stylistic only. |
| Section 3.1 | Universal command center port | Project A's is already richer (1,155 LOC vs 180 LOC); nothing to steal. |
| Section 3.5 | What-If full port | Project A dominates; only the margin labels (item K above) are worth borrowing. |

## Active-video deferral rule (2026-04-21)

Videos 1 and 2 are already recorded — safe to change anything used by them. Videos 3 (in Gemini review) and 4 (pending manual Python recording) are still active.

**Active-video touch inventory:**

- **Video 3** touches the 11 UTL modules imported into `Sample_Quarterly_ReportV2.xlsm` — Director calls their `Director*` silent wrappers. Also uses `modUTL_Core` indirectly (via `UTL_TurboOn` etc. called by those wrappers).
- **Video 4** is manual Python from Command Prompt. Uses these 8 scripts specifically: `aging_report.py`, `bank_reconciler.py`, `compare_files.py`, `forecast_rollforward.py`, `fuzzy_lookup.py`, `pdf_extractor.py`, `variance_analysis.py`, `variance_decomposition.py`.

**Rule for cherry-pick ports:**
- Ports that add NEW subs/functions to existing Video 3/4 files are SAFE (additive — no existing behavior changes).
- Ports that modify existing Video 3/4 code behavior → DEFER until both videos are recorded.
- Port source code freely. Delay the *re-import into the .xlsm* / *modification of a Video 4 script* only when it would otherwise introduce noise or risk during active recording.

**Items flagged as defer-re-import-only (source code done, hold the re-import):**

**Update 2026-04-21:** Deferrals resolved. Video 3 shipped + all Batch 1+2 files re-imported into `Sample_Quarterly_ReportV2.xlsm`, tested, and live. Intelligence registered in Command Center static category at position 6. Narrative-on-wrong-column bug found and fixed (commit 8eff337).

| Item | File | Status |
|------|------|--------|
| C | `modUTL_Core.bas` (UTL_DetectHeaderRow) | 🔵 Live in .xlsm |
| D | `modUTL_Compare.bas` (UTL_QuickRowCompareCount) | 🔵 Live in .xlsm, auto-discovered in Command Center |
| B | `modUTL_Audit.bas` (CreateRunReceiptSheet) | 🔵 Live in .xlsm |
| A | `modUTL_Intelligence.bas` (Materiality + Narratives + Scorecard) | 🔵 Live in .xlsm, position 6 in Command Center, all 3 tools tested |
| K | `modWhatIf_v2.1.bas` (MarginVerdict + AppendMarginVerdictRow) | Source done; demo file (Video 2), import at leisure |

**Items to build after videos are done (video code would be modified):**

- *(empty — Batch 4 dual-logging is additive to demo file, Batch 5 is docs only)*

---

## Status legend

- ⬜ Planned — listed, not yet started
- 🟡 In progress — currently being ported
- 🟢 Done — ported, synced, committed
- 🔵 Done + verified — ran against sample file, no bugs

## Batch plan

| Batch | Items | Risk | Feel |
|-------|-------|------|------|
| 1 | K, B, C | All S-effort, low risk, additive | Quick wins, warm up |
| 2 | A, D | One M, one S | The big capability add |
| 3 | E1-E4, F, G, H, I/J | Python only | Standalone effort |
| 4 | L | Moderate, interacts with existing logger | Careful integration |
| 5 | 4, 5, 9, 10 | Docs only | No code risk |

## Log

| Date | Batch | Items | Notes |
|------|-------|-------|-------|
| 2026-04-20 | — | Tracker created | Plan approved, Batch 1 about to start |
| 2026-04-20 | 1 | K, B, C done | `MarginVerdict` + `AppendMarginVerdictRow` appended to `modWhatIf_v2.1.bas`; `CreateRunReceiptSheet` appended to `modUTL_Audit.bas`; `UTL_DetectHeaderRow` appended to `modUTL_Core.bas`. All additive, no existing code touched. Ready for next test-run import. |
| 2026-04-21 | 2 | A, D done | New `modUTL_Intelligence.bas` created (231 LOC) with `MaterialityClassifierActiveSheet`, `GenerateExceptionNarrativesActiveSheet`, `DataQualityScorecardActiveSheet` + 5 private helpers. `UTL_QuickRowCompareCount` + `BuildRowHashMap` appended to `modUTL_Compare.bas`. Project A header block + brand-styled output sheets. No Director wrappers yet (coworker manual use only). |
| 2026-04-21 | 3 | E1-E4, F, G, H, I/J done | New `RecTrial\UniversalToolkit\python\ZeroInstall\` folder with 7 stdlib-only scripts + README (`profile_workbook.py`, `sanitize_dataset.py`, `compare_workbooks.py`, `build_exec_summary.py`, `variance_classifier.py`, `scenario_runner.py`, `sheets_to_csv.py`). `word_report.py` gained `build_talking_points()` helper + opt-in `--talking-points` CLI flag. All 8 Python files verified ast-parse clean. None overlap with Video 4's 8 active scripts. |

---

*Update this file as each batch completes. Keep it short — this is an index, not a report.*
