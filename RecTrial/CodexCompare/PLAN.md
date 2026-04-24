# Project Plan — TO BE FILLED BY CODEX (Stage 1)

> **Instructions for Codex:** This file is a template. At Stage 1, replace every `<<FILL IN>>` block below with your proposed answers based on what you learned from `README.md`, `CONTEXT.md`, `CONSTRAINTS.md`, `BRAND.md`, and the two files in `samples/`.
>
> Do not write code yet. Do not delete or reorder the sections below. Add extra sections at the bottom if useful, but preserve this structure.
>
> When finished, commit this file as `PLAN.md` and wait for the user to reply with the literal word **"approved"** before moving to Stage 2.
>
> Every stage after Stage 1, re-read this file first and check your work against it. If you need to change the plan mid-project, propose the edit here and wait for re-approval before executing.

---

## 1. Sample File Inventory

### `samples/ExcelDemoFile_adv.xlsm`
Workbook profile: advanced P&L model with 15 sheets, 8 workbook-level named ranges, hidden audit log, and embedded VBA project (`xl/vbaProject.bin`).

Sheet-by-sheet inventory:

1) **CrossfireHiddenWorksheet**  
- Visibility: Visible  
- Purpose (inferred): raw GL-like transactional feed used as source data for reconciliation and rollups.  
- Data shape: ~511 rows x 7 columns (`A1:G511`), tabular with headers in row 1.  
- Key columns: `ID`, `Date`, `Department`, `Product`, `Expense Category`, `Vendor`, `Amount`.  
- Existing VBA relevance: likely read by reconciliation/allocation modules (see module names below).

2) **Disclaimer**  
- Visibility: Visible  
- Purpose: legal/demo disclaimer and context text.  
- Data shape: ~28 rows x 14 columns (`A1:N28`), mostly formatted text blocks.

3) **Assumptions**  
- Visibility: Visible  
- Purpose: driver table and model assumptions.  
- Data shape: ~33 rows x 4 columns (`A1:D33`), structured key-value table.  
- Key named-range links: `AWS_Shared_Pct`, `InsureSight_Split`, `Revenue_Share_Pct_iGO` (plus related driver cells).

4) **Data Dictionary**  
- Visibility: Visible  
- Purpose: reference master data for products, departments, vendors.  
- Data shape: ~54 rows x 5 columns (`A1:E54`) with multiple sections.  
- Key named-range links: `Products_Table`, `Departments_Table`, `Vendors_Table`.

5) **AWS Allocation**  
- Visibility: Visible  
- Purpose: model for AWS pool allocation by product/driver.  
- Data shape: ~42 rows x 6 columns (`A1:F42`), sectioned model sheet.

6) **Report-->**  
- Visibility: Visible  
- Purpose: executive summary/landing report with KPI cards.  
- Data shape: ~29 rows x 10 columns (`A1:J29`), formatted report blocks.

7) **P&L - Monthly Trend**  
- Visibility: Visible  
- Purpose: consolidated monthly P&L trend.  
- Data shape: ~52 rows x 19 columns (`A1:S52`), matrix (line items x months/summary columns).

8) **Product Line Summary**  
- Visibility: Visible  
- Purpose: product-level P&L summary by month and sections.  
- Data shape: ~80 rows x 18 columns (`A1:R80`), multi-section matrix.

9) **Functional P&L - Monthly Trend**  
- Visibility: Visible  
- Purpose: function-level and product-level trend detail.  
- Data shape: ~147 rows x 18 columns (`A1:R147`), dense matrix across monthly columns.

10) **Functional P&L Summary - Jan 25**  
- Visibility: Visible  
- Purpose: January functional summary by product/total.  
- Data shape: ~37 rows x 5 columns (`A1:E37`).

11) **Functional P&L Summary - Feb 25**  
- Visibility: Visible  
- Purpose: February functional summary by product/total.  
- Data shape: ~37 rows x 5 columns (`A1:E37`).

12) **Functional P&L Summary - Mar 25**  
- Visibility: Visible  
- Purpose: March functional summary by product/total.  
- Data shape: ~37 rows x 5 columns (`A1:E37`).

13) **US January 2025 Natural P&L**  
- Visibility: Visible  
- Purpose: natural-account style January P&L by product + total.  
- Data shape: ~77 rows x 5 columns (`A1:E77`), sectioned financial statement.

14) **Charts & Visuals**  
- Visibility: Visible  
- Purpose: dashboard visuals and product selector-driven charts.  
- Data shape: ~265 rows x 20 columns (`A1:T265`), chart support ranges and dashboard layout.

15) **Checks**  
- Visibility: Visible  
- Purpose: reconciliation/validation results dashboard.  
- Data shape: ~28 rows x 9 columns (`A1:I28`) with log-style rows (`Check Name`, `Expected`, `Actual`, `Status`, etc.).

16) **VBA_AuditLog**  
- Visibility: **Very Hidden**  
- Purpose: system audit trail for macro actions and status.  
- Data shape: currently header-only ~1 row x 6 columns (`A1:F1`).  
- Key columns: `Timestamp`, `User`, `Module`, `Procedure`, `Message`, `Status`.  
- Named range: `_xlnm._FilterDatabase` points to this sheet header region.

Workbook named ranges found:
- `_xlnm._FilterDatabase` → `VBA_AuditLog!$A$1:$F$1`
- `AWS_Shared_Pct` → `Assumptions!$B$12`
- `Departments_Table` → `'Data Dictionary'!$A$14:$C$20`
- `InsureSight_Split` → `Assumptions!$B$6`
- `_xlnm.Print_Area` → `Disclaimer!$B$2:$M$19`
- `Products_Table` → `'Data Dictionary'!$A$6:$D$9`
- `Revenue_Share_Pct_iGO` → `Assumptions!$B$8`
- `Vendors_Table` → `'Data Dictionary'!$A$40:$D$54`

Existing VBA (best-effort inventory from embedded `vbaProject.bin` strings; module export not yet performed):
- Module names detected (high-confidence): `modAdmin`, `modAllocation`, `modAuditTools`, `modAWSRecompute`, `modConsolidation`, `modDashboard`, `modDashboardAdvanced`, `modDataGuards`, `modDataQuality`, `modDataSanitizer`, `modETLBridge`, `modExecBrief`, `modFormBuilder`, `modForecast`, `modLogger`, `modMasterMenu`, `modMonthlyTabGenerator`, `modPDFExport`, `modProgressBar`, `modReconciliation`, `modScenario`, `modSensitivity`, `modSplashScreen`, `modTrendReports`, `modUtilities`, `modVarianceAnalysis`, `modVersionControl`, `modWhatIf`.
- Procedure names detected (sample): `modReconciliation.RunAllChecks`, `modPDFExport.ExportReportPackage`, `modDashboard.BuildDashboard`, `modFormBuilder.BuildCommandCenter`, `modSplashScreen.ShowSplash`, `modVersionControl.SaveVersion`, `modLogger.LogAction`.

---

### `samples/Sample_Quarterly_ReportV2.xlsm`
Workbook profile: coworker-style reporting file with mixed data quality patterns, 9 sheets, no workbook named ranges, and embedded VBA project (`xl/vbaProject.bin`).

Sheet-by-sheet inventory:

1) **Cover**  
- Visibility: Visible  
- Purpose: front page and report metadata/instructions.  
- Data shape: ~35 rows x 8 columns (`A1:H35`), formatted text sections.

2) **Pivot_SalesByRegion**  
- Visibility: Visible  
- Purpose: summary pivot output of revenue by region.  
- Data shape: ~8 data rows in `A3:B10` with `Row Labels`/`Sum of Amount` structure.

3) **Pivot_SalesByRep**  
- Visibility: Visible  
- Purpose: summary pivot output of revenue by sales rep/year grouping.  
- Data shape: ~17 data rows in `A3:B19`.

4) **Q1 Revenue**  
- Visibility: Visible  
- Purpose: base pipeline/revenue transaction detail table (contains intentionally inconsistent formats).  
- Data shape: ~43 rows x 9 columns (`A1:I43`), tabular.  
- Key columns: `Region`, `Sales Rep`, `Product`, `Customer`, `Date`, `Amount`, `Status`, `Commission %` (+ one extra column I).

5) **Q1 Expenses**  
- Visibility: Visible  
- Purpose: departmental expense transactions and approvals.  
- Data shape: ~26 rows x 7 columns (`A1:G26`), tabular.  
- Key columns: `Department`, `Category`, `Vendor`, `Invoice Date`, `Amount`, `Approved By`, `PO Number`.

6) **Q1 Revenue v2**  
- Visibility: Visible  
- Purpose: revised revenue table version for comparison/diff demos.  
- Data shape: ~44 rows x 9 columns (`A1:I44`), tabular and similar schema to Q1 Revenue.

7) **Budget Summary**  
- Visibility: Visible  
- Purpose: department budget vs actual and variance status.  
- Data shape: ~24 rows x 6 columns (`A1:F24`) with precomputed variance fields.

8) **Contact List**  
- Visibility: Visible  
- Purpose: stakeholder/contact roster with intentionally mixed phone formatting.  
- Data shape: ~16 rows x 6 columns (`A1:F16`).

9) **Archive_Q4_2025**  
- Visibility: **Hidden**  
- Purpose: archived prior-quarter snapshot with explicit “do not modify” note.  
- Data shape: ~9 rows x 5 columns (`A1:E9`).

Workbook named ranges found:
- None.

Existing VBA (best-effort inventory from embedded `vbaProject.bin` strings; module export not yet performed):
- Universal toolkit-style modules detected: `modUTL_Audit`, `modUTL_Branding`, `modUTL_ColumnOps`, `modUTL_CommandCenter`, `modUTL_Comments`, `modUTL_Compare`, `modUTL_Consolidate`, `modUTL_DataCleaning`, `modUTL_DataSanitizer`, `modUTL_ExecBrief`, `modUTL_Finance`, `modUTL_Formatting`, `modUTL_Highlights`, `modUTL_LookupBuilder`, `modUTL_PivotTools`, `modUTL_ProgressBar`, `modUTL_SheetTools`, `modUTL_SplashScreen`, `modUTL_TabOrganizer`, `modUTL_ValidationBuilder`, `modUTL_WhatIf`, `modUTL_WorkbookMgmt`.
- Non-UTL modules also detected: `modConfig`, `modDashboard`, `modDashboardAdvanced`, `modDataQuality`, `modDataSanitizer`, `modDirector`, `modExecBrief`, `modIntegrationTest`, `modLogger`, `modPerformance`, `modReconciliation`, `modSensitivity`, `modTimeSaved`, `modVarianceAnalysis`, `modVersionControl`, `modWhatIf`.
- Procedure samples detected: `modUTL_Audit.WorkbookErrorScanner`, `modUTL_DataCleaning.UnmergeAndFillDown`, `modUTL_Compare.CompareSheets`, `modUTL_Consolidate.ConsolidateSheets`, `modUTL_ExecBrief.GenerateExecBrief`, `modUTL_Finance.TrialBalanceChecker`, `modUTL_WhatIf.RunWhatIf`, `modUTL_WorkbookMgmt.WorkbookHealthCheckagentId`.

### Questions about the samples
- Should I treat the existing VBA inside both `.xlsm` files as **reference-only** (inspiration), or as legacy code we are expected to partially preserve/reuse in later stages?
- `Archive_Q4_2025` is hidden and marked “do not modify.” Should Stage 2+ tools default to skipping hidden/archive sheets unless the user explicitly opts in?
- In `ExcelDemoFile_adv.xlsm`, there are monthly summary sheets for Jan–Mar but not visible equivalents for all months. Should Prong-2 features assume full-year monthly detail exists elsewhere, or only use what is present today?
- Do you want Prong-2 outputs written into **new generated sheets only** (safer demo), or can approved features update existing report/dashboard sheets directly?

---

## 2. Proposed Folder Structure

```
vba/
  universal/
    modUTL_Core.bas                    — shared helpers (sheet discovery, safe range detection, logging hooks)
    modUTL_DataSanitizer.bas           — robust cleanup across unknown data shapes
    modUTL_CompareConsolidate.bas      — cross-sheet/range compare + consolidation workflows
    modUTL_Intelligence.bas            — materiality rules, anomaly scoring, smart flags
    modUTL_OutputPack.bas              — branded export bundles + summary sheets
    modUTL_CommandCenter.bas           — one-click launcher and action catalog
  demo/
    modDemo_Config.bas                 — workbook-specific constants + named range map
    modDemo_ReconciliationEngine.bas   — deep checks tied to P&L model logic
    modDemo_VarianceNarrative.bas      — auto commentary sentences from threshold logic
    modDemo_ExecBriefPack.bas          — CFO-ready brief sheet + PDF pack generator
    modDemo_WhatIfScenario.bas         — driver-based scenario generation and comparison
    modDemo_AuditTrail.bas             — writes execution events to VBA_AuditLog

python/
  universal/
    profile_workbook.py                — workbook profiler (tabs, headers, types, quality score)
    sanitize_dataset.py                — batch cleaner for CSV/XLSX sources before Excel load
    compare_workbooks.py               — cell/range diff report for two workbooks
    build_exec_summary.py              — plain-English KPI summary from arbitrary tabular input
  demo/
    pnl_data_extract.py                — extracts P&L source structures for demo workflows
    variance_classifier.py             — labels favorable/unfavorable/material variance patterns
    scenario_runner.py                 — generates scenario result tables for assumptions changes
    export_brief_package.py            — creates branded deliverable artifacts for leadership

sql/
  universal/
    template_gl_extract.sql            — generic GL extract template with parameter placeholders
    template_revenue_extract.sql       — generic revenue extract template
  demo/
    demo_pnl_reconciliation_view.sql   — demo view for recon cross-check logic
    demo_variance_fact.sql             — material variance fact-table style output

guides/
  copilot-prompt-guide.md              — adaptation bridge for file-specific (Prong 2) code
  universal-toolkit-user-guide.md      — install/use universal tools in any workbook
  demo-walkthrough-guide.md            — runbook for end-to-end P&L demo storyline
  brand-styling-reference.md           — operational checklist derived from BRAND.md
  troubleshooting-reference.md         — common failures + fixes across VBA/Python/Excel

videos/
  video-1-executive-hook.md            — highlight reel script
  video-2-demo-workbook-deep-dive.md   — Prong-2 walkthrough script
  video-3-universal-toolkit-in-action.md — Prong-1 coworker workflow script
  video-4-python-sql-integration.md    — cross-source automation script
  video-5-copilot-adaptation-lab.md    — optional adaptation guide walkthrough script

artifacts/
  screenshots/                         — guide/video stills captured during build stages
  sample-outputs/                      — generated PDFs/CSV/log outputs used in demos
```

---

## 3. Prong 1 — Universal Toolkit Modules

1) **modUTL_DataSanitizer**  
- Purpose: one-click cleanup for messy finance tabs without hardcoded columns.  
- Example tools: text-number normalization; date harmonization across mixed formats; floating-point tail fixer + audit diff preview.  
- Constraint threshold crossed: **performance/scale**, **portability**, **multi-step automation**.

2) **modUTL_CompareConsolidate**  
- Purpose: compare and merge data across sheets/workbooks with explicit change logs.  
- Example tools: sheet-vs-sheet delta highlighter with materiality filter; wildcard sheet consolidation; two-file variance map export.  
- Thresholds: **cross-file/cross-source**, **output generation**, **portability**.

3) **modUTL_Intelligence**  
- Purpose: inject finance judgment logic that native conditional formatting cannot perform.  
- Example tools: materiality classifier (`% and absolute $`); anomaly triage tags; automated exception commentary snippets.  
- Thresholds: **intelligence/decision logic**, **teaching leverage**, **multi-step automation**.

4) **modUTL_OutputPack**  
- Purpose: generate branded deliverables from arbitrary workbook inputs.  
- Example tools: executive one-pager sheet builder; sectioned PDF pack export with run timestamp; change-summary cover page.  
- Thresholds: **output generation**, **multi-step automation**, **teaching leverage**.

5) **modUTL_CommandCenter**  
- Purpose: single-entry UI so non-technical users run complex workflows in one click.  
- Example tools: categorized tool launcher; run history panel; preflight checks + success/failure toast.  
- Thresholds: **portability**, **multi-step automation**, **teaching leverage**.

6) **Python Universal Utilities (`python/universal/*`)**  
- Purpose: handle larger-volume profiling/diff/sanitization tasks beyond comfortable VBA scale.  
- Example tools: workbook profile report; bulk compare report; auto-generated plain-English KPI digest.  
- Thresholds: **performance/scale**, **cross-source**, **output generation**.

---

## 4. Prong 2 — Demo-Specific Features

1) **Material Variance Narrative Engine**  
- Purpose: generate executive-ready commentary for major P&L movements.  
- Operates on sheets: `P&L - Monthly Trend`, `Product Line Summary`, `Functional P&L - Monthly Trend`, `Assumptions`.  
- Output: new `Exec_Variance_Narrative` sheet + optional PDF page.  
- Why Prong 2: depends on specific line-item layout and business logic of this P&L model.

2) **Reconciliation Command Run**  
- Purpose: execute a full reconciliation suite with pass/fail audit evidence.  
- Operates on sheets: `CrossfireHiddenWorksheet`, `Checks`, `VBA_AuditLog`, plus target summary tabs.  
- Output: refreshed `Checks` table, appended audit log rows, exception detail sheet.  
- Why Prong 2: rules are tied to this workbook’s structures, source tab names, and expected totals.

3) **Executive Brief Pack Builder**  
- Purpose: build a branded CFO pack from current period results in one action.  
- Operates on sheets: `Report-->`, `Charts & Visuals`, `P&L - Monthly Trend`, `Checks`.  
- Output: new `Exec_Brief` sheet and multi-page branded PDF packet.  
- Why Prong 2: references specific dashboard/KPI placement unique to this workbook.

4) **AWS Allocation Sensitivity Runner**  
- Purpose: run assumption shocks and show impact on cost allocation and margin.  
- Operates on sheets: `Assumptions`, `AWS Allocation`, `Product Line Summary`.  
- Output: scenario comparison sheet (`Scenario_Compare`) and impact chart panel.  
- Why Prong 2: hard-linked to named assumptions and this model’s allocation mechanics.

5) **Period Close Storyboard Generator**  
- Purpose: produce a month-close “what changed and why” timeline artifact.  
- Operates on sheets: `Functional P&L Summary - Jan 25`, `...Feb 25`, `...Mar 25`, `Checks`, `VBA_AuditLog`.  
- Output: `Close_Storyboard` sheet with key deltas, check outcomes, and commentary blocks.  
- Why Prong 2: relies on specific period tabs and this workbook’s close-control narrative.

6) **Command Center (Demo Mode)**  
- Purpose: curated UI with high-impact buttons for live executive presentation.  
- Operates on sheets: creates/updates dedicated control sheet + triggers all above features.  
- Output: branded control interface + status cards + elapsed-time metric.  
- Why Prong 2: optimized around this demo workbook flow, not generic use.

---

## 5. CoPilot Prompt Guide Outline

1) **Who this guide is for**  
- Audience: non-developers who want to adapt Prong-2 macros to different workbook layouts.

2) **What you need before starting**  
- Prerequisites checklist (Excel version, macro settings, where code lives, backup steps).

3) **How to describe your workbook to CoPilot (plain-English template)**  
- Tab inventory template (sheet name, purpose, header row, key columns, named ranges).  
- Business rules template (what counts as favorable/unfavorable/material).

4) **Starter prompt (copy/paste block)**  
- “Here is my workbook map + here is the original macro + adapt for my file + keep behavior equivalent + show changed lines.”

5) **Adaptation workflow**  
- Step-by-step: paste workbook map → paste source code → request adaptation → review diff → test in copy.

6) **Worked examples (one per Prong-2 feature)**  
- Variance narrative adaptation example.  
- Reconciliation run adaptation example.  
- Exec brief pack adaptation example.  
- AWS sensitivity adaptation example.  
- Close storyboard adaptation example.  
- Demo command center adaptation example.

7) **Validation checklist after CoPilot edits**  
- Preflight checks, test data run, audit log verification, output spot-check.

8) **Troubleshooting when CoPilot gets it wrong**  
- Missing/renamed columns, wrong header row, hidden sheet assumptions, formula reference breaks, type conversion issues.

9) **Safe rollback and escalation**  
- How to restore backup and how to ask CoPilot for minimal targeted fixes.

---

## 6. Video Lineup

1) **Video 1 — “From Spreadsheet to Finance Control Tower”**  
- Target length: 7 minutes  
- Core premise: fast, executive-friendly proof that one-click automation can turn a complex workbook into insights and decision artifacts.  
- Main features: Command Center, Reconciliation Run, Executive Brief Pack, Variance Narrative.  
- Opening hook: “In the next 90 seconds, one click will run checks, write commentary, and build a CFO-ready pack.”  
- CTA: “Start with the Command Center in your own file using the universal toolkit guide.”

2) **Video 2 — “Deep Dive: P&L Demo Workbook Automation”**  
- Target length: 9 minutes  
- Core premise: detailed walkthrough of every Prong-2 feature on `ExcelDemoFile_adv.xlsm`, including before/after outputs and auditability.  
- Main features: all Prong-2 modules + audit log flow.  
- Opening hook: “This is the exact workbook flow Finance can run every close cycle.”  
- CTA: “Use the demo walkthrough guide to reproduce this run step-by-step.”

3) **Video 3 — “Universal Toolkit on a Plain Coworker File”**  
- Target length: 8 minutes  
- Core premise: demonstrate plug-and-play utility on `Sample_Quarterly_ReportV2.xlsm` with no sheet hardcoding.  
- Main features: sanitizer, compare/consolidate, intelligence flags, output pack, command center.  
- Opening hook: “No model redesign, no coding—drop in the toolkit and run.”  
- CTA: “Install the add-in and run the 15-minute quickstart from the user guide.”

4) **Video 4 — “Python + SQL: Extending Excel Beyond the Workbook”**  
- Target length: 8 minutes  
- Core premise: show cross-source ETL, profiling, and high-volume comparison feeding back into Excel outputs.  
- Main features: SQL templates, Python profile/diff scripts, output artifact handoff.  
- Opening hook: “When Excel gets heavy, this hybrid layer keeps speed and control.”  
- CTA: “Clone the script templates and point them at your own extracts.”

5) **Video 5 — “CoPilot Adaptation Lab”**  
- Target length: 6 minutes  
- Core premise: teach non-developers how to adapt file-specific macros safely with structured prompts.  
- Main features: workbook map template, starter prompt, worked adaptation, troubleshooting loop.  
- Opening hook: “If your workbook looks different, this is how you still use everything.”  
- CTA: “Copy the starter prompt and run one feature adaptation today.”

---

## 7. Training / User Guide Lineup

1) **Universal Toolkit User Guide**  
- Audience: coworker end users (non-technical).  
- Length estimate: ~12–16 pages.  
- Key sections: install, first run, tool catalog, safe mode/backups, troubleshooting, FAQ.

2) **Demo Walkthrough Guide (P&L Workbook)**  
- Audience: presenter + finance power users.  
- Length estimate: ~14–18 pages.  
- Key sections: demo setup, sequence of clicks, expected outputs, reset steps, talking points.

3) **CoPilot Prompt Guide**  
- Audience: coworkers adapting Prong-2 features to their own files.  
- Length estimate: ~16–20 pages.  
- Key sections: file mapping method, prompt templates, feature-by-feature adaptation examples, validation checklist, error recovery.

4) **Brand Styling Reference**  
- Audience: anyone creating sheets/reports/PDFs/videos in this project.  
- Length estimate: ~6–8 pages.  
- Key sections: color palette, typography, layout blocks, chart standards, common brand mistakes.

5) **Troubleshooting Reference**  
- Audience: installers/support champions.  
- Length estimate: ~8–10 pages.  
- Key sections: macro trust settings, blocked file issues, broken references, performance lag, rollback procedures.

---

## 8. Open Questions

1) **Reuse policy for existing macros:** should Stage 2+ explicitly build on existing VBA in the sample files (where useful), or should we treat those files as data-only examples and create a clean new codebase?  
   - Option A (recommended): clean new modules in repo, sample VBA used only as inspiration.  
   - Option B: selectively port/refactor sample VBA into repo modules.

2) **Distribution target for Prong 1:** do you want the universal toolkit delivered primarily as:  
   - Option A (recommended): `.xlam` add-in first, with optional Python companion scripts.  
   - Option B: workbook-based macro library first, add-in later.  
   - Option C: Python-first CLI with optional VBA wrappers.

3) **Demo interaction style:** for live executive demos, should actions be launched from:  
   - Option A (recommended): one branded “Command Center” sheet with large buttons/status cards.  
   - Option B: custom ribbon only.  
   - Option C: both command center + ribbon shortcuts.

4) **Output governance:** where should generated outputs (PDFs/logs/scenario files) be written by default?  
   - Option A (recommended): `/artifacts/` under repo-relative working folder.  
   - Option B: same folder as active workbook.  
   - Option C: user-selected folder each run.

5) **Scope of SQL in this demo:** should SQL be templates/mock scripts only, or do you want simulated end-to-end runs against local sample CSVs to mimic DB pulls?  
   - Option A (recommended): template + simulated local run (no live DB dependency).  
   - Option B: templates only.

6) **Brand compliance for videos/guides:** confirm that we should apply iPipeline styling rules to markdown guide formatting patterns and visual callouts (colors/fonts references), even though markdown itself cannot enforce Excel font rendering.

7) **Tolerance for modifying existing visible sheets in Prong 2:** should we default to writing new output sheets to avoid disturbing originals during demos?  
   - Option A (recommended): non-destructive, new output sheets by default.  
   - Option B: direct in-place updates where faster.

---

## 9. Proposed Stage 2+ Breakdown

- **Stage 2:** Build universal VBA foundation (`modUTL_Core`, `modUTL_DataSanitizer`, `modUTL_CommandCenter`) + smoke tests on both sample files.
- **Stage 3:** Build remaining universal VBA modules (`CompareConsolidate`, `Intelligence`, `OutputPack`) + usage logging + user messages.
- **Stage 4:** Build universal Python utilities (`profile_workbook.py`, `sanitize_dataset.py`, `compare_workbooks.py`, `build_exec_summary.py`) + sample run outputs.
- **Stage 5:** Build demo-specific core (`modDemo_Config`, `modDemo_AuditTrail`, `modDemo_ReconciliationEngine`, `modDemo_VarianceNarrative`).
- **Stage 6:** Build demo-specific advanced outputs (`modDemo_ExecBriefPack`, `modDemo_WhatIfScenario`, demo command center integration).
- **Stage 7:** Write `copilot-prompt-guide.md` with full worked examples for each Prong-2 feature.
- **Stage 8:** Write user guides (`universal-toolkit-user-guide.md`, `demo-walkthrough-guide.md`, `troubleshooting-reference.md`, `brand-styling-reference.md`).
- **Stage 9:** Write video scripts 1–3 with timestamps, narration, and on-screen actions.
- **Stage 10:** Write video scripts 4–5 + end-to-end polish pass (consistency, plain-English check, brand check, constraints check).

Estimated total after Stage 1: **9 additional stages**.

---

## 10. Risks & Trade-offs

1) **Risk: existing sample workbooks already contain significant VBA footprint.**  
Trade-off: building clean repo modules avoids hidden dependencies, but may duplicate some functionality already present in sample files.

2) **Risk: VBA inventory is partially inferred from `vbaProject.bin` strings in this environment.**  
Trade-off: module/procedure names are high-confidence, but full source-level behavior needs explicit export/open in Excel during build/test stages.

3) **Risk: over-building features that violate the `<5 native clicks` rule.**  
Trade-off: we will prioritize compound workflows/intelligence/output generation; any thin wrappers around native GUI features will be excluded unless part of a larger automation.

4) **Risk: balancing “world-class polish” with staged delivery speed.**  
Trade-off: more polished UI/output and strong non-technical guides increase timeline but are essential for CFO/CEO and 2,000-user audience.

5) **Risk: Python and VBA interoperability friction on coworker machines.**  
Trade-off: keeping VBA-first for core actions maximizes adoption; Python features will remain optional accelerators with clear prerequisites.

6) **Risk: hidden/archive sheets might get unintentionally altered by bulk tools.**  
Trade-off: default behavior will exclude hidden/very hidden sheets unless explicit inclusion is selected.

7) **Rule pressure area:** some requested capabilities (e.g., tab organization, validation builders) can drift toward native-feature wrappers.  
Mitigation: enforce a strict “multi-step + portability + intelligence/output” threshold gate during implementation reviews.

---

## Approval

> The user will reply with the literal word **"approved"** below when this plan is ready to execute. Do not assume approval without that word.

**Status:** PENDING USER REVIEW
**Approved by:** <<NAME + DATE>>
