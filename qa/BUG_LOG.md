# KBT P&L Toolkit — Centralized Bug Log

> **Last Updated:** 2026-03-11
> **Total Bugs Found:** 35
> **Total Bugs Fixed:** 35
> **Open Bugs:** 0

---

## Summary by Discovery Phase

| Phase | Date | Bugs Found | Severity | Commits |
|-------|------|-----------|----------|---------|
| Testing (user-found) | 2026-03-03/04 | 9 | 2 CRITICAL, 5 HIGH, 2 MODERATE | af44453, 19320db, 6f40f91, ed3276f, 3024c44, 6c17bd5, 847a982, 304743b |
| Self-Review | 2026-03-04 | 12 | 2 CRITICAL, 9 MODERATE, 1 INFO | 22ba831 |
| Code Review #1 | 2026-03-05 | 4 | 3 MODERATE, 1 HIGH | (included in later commits) |
| Pre-Delivery Code Review | 2026-03-07 | 7 | 4 HIGH, 2 MODERATE, 1 MODERATE | 6818b01, b132885 |
| 5-Pass Bug Review | 2026-03-11 | 3 | 2 HIGH, 1 LOW | (included in session commit) |
| **Total** | | **35** | | |

---

## Top Recurring Patterns

| Pattern | Occurrences | Description |
|---------|------------|-------------|
| LogAction signature | 13 | Double passed as 4th arg (status String) — corrupts audit log |
| SpecialCells rng reset | 3 | rng not set to Nothing before SpecialCells in loops |
| HDR_ROW_REPORT vs row 1 | 3 | Code scanning row 1 (company title) instead of row 4 (headers) |
| Chr() range > 255 | 4 | Chr(8212) and Chr(9472) crash VBA (only handles 0-255) |
| Hex-to-decimal BGR | 2 | VBA color constants use BGR byte order, not RGB |

---

## All Bugs — Chronological

### Phase 1: Testing Bugs (2026-03-03/04)

| # | ID | Severity | Module | Description | Fix | Commit |
|---|-----|----------|--------|-------------|-----|--------|
| 1 | BUG-T2.01 | HIGH | modConfig_v2.1.bas | 9 missing sheet-name constants (SH_GL, SH_TECH_DOC, etc.) | Added all 9 constants | af44453 |
| 2 | BUG-T2.03a | HIGH | modConfig_v2.1.bas | CLR_NAVY = 2050943 decodes to tan/brown, not navy (BGR byte order) | Changed to 7949855 | 19320db |
| 3 | BUG-T2.03b | HIGH | modConfig_v2.1.bas | CLR_ALT_ROW = 15651567 decodes to pink/lavender, not light blue | Changed to 16380653 | 19320db |
| 4 | BUG-T2.04a | MODERATE | modMonthlyTabGenerator_v2.1.bas | UpdateHeaderText declared Private — can't call from Immediate Window | Added TestUpdateHeaderText wrapper | 6f40f91 |
| 5 | BUG-T2.04b | MODERATE | modMonthlyTabGenerator_v2.1.bas | Test wrapper wrote "Mar 25" without Text format — auto-converted to date | Added NumberFormat = "@" | ed3276f |
| 6 | BUG-T4.04a | HIGH | pnl_tests.py | Windows PermissionError on temp file cleanup during pytest | Fixed temp file handling | 3024c44 |
| 7 | BUG-T5.01a | CRITICAL | modDashboard_v2.1.bas | CreateExecutiveDashboard read row 1 instead of row 4 (HDR_ROW_REPORT) | Changed to HDR_ROW_REPORT | 6c17bd5 |
| 8 | BUG-T5.01b | CRITICAL | modDashboard_v2.1.bas | Error 5 (Invalid procedure call) crash in CreateExecutiveDashboard | Fixed row/column detection | 847a982 |
| 9 | BUG-T5.02 | HIGH | modDashboard_v2.1.bas | WaterfallChart hardcoded "Total Revenue" — P&L Trend may use variants | Added multi-variant fallback search | 304743b |

### Phase 2: Self-Review Bugs (2026-03-04, commit 22ba831)

| # | ID | Severity | Module | Line | Description | Fix |
|---|-----|----------|--------|------|-------------|-----|
| 10 | SR-01 | CRITICAL | modReconciliation_v2.1.bas | 292 | dateCol = 5 (Category E) instead of COL_GL_DATE = 2 (Date B) | Changed to COL_GL_DATE |
| 11 | SR-02 | INFO | modReconciliation_v2.1.bas | 293 | amtCol = 7 hardcoded instead of COL_GL_AMOUNT | Changed to COL_GL_AMOUNT |
| 12 | SR-03 | CRITICAL | modVarianceAnalysis_v2.1.bas | 221 | GenerateCommentary read row 1 for tLastCol — FY/Budget loops never ran | Changed to HDR_ROW_REPORT |
| 13 | SR-04 | MODERATE | modDashboard_v2.1.bas | 99 | LogAction: elapsed (Double) as 4th arg | Moved into message string |
| 14 | SR-05 | MODERATE | modDashboard_v2.1.bas | 369 | Same LogAction issue — CreateExecutiveDashboard | Moved into message string |
| 15 | SR-06 | MODERATE | modDashboard_v2.1.bas | 534 | Same LogAction issue — WaterfallChart | Moved into message string |
| 16 | SR-07 | MODERATE | modDashboard_v2.1.bas | 669 | Same LogAction issue — ProductComparison | Moved into message string |
| 17 | SR-08 | MODERATE | modDashboard_v2.1.bas | 1226 | Same LogAction issue — CreateSmallMultiplesGrid | Moved into message string |
| 18 | SR-09 | MODERATE | modDemoTools_v2.1.bas | — | Same LogAction issue — CreatePrintableExecSummary | Moved into message string |
| 19 | SR-10 | MODERATE | modTrendReports_v2.1.bas | 153 | Same LogAction issue — CreateRolling12MonthView | Moved into message string |
| 20 | SR-11 | MODERATE | modMonthlyTabGenerator_v2.1.bas | 110 | Same LogAction issue — GenerateMonthlyTabs | Moved into message string |
| 21 | SR-12 | MODERATE | modMonthlyTabGenerator_v2.1.bas | 230 | Same LogAction issue — GenerateNextMonthOnly | Moved into message string |

### Phase 3: Code Review #1 (2026-03-05)

| # | ID | Severity | Module | Description | Fix |
|---|-----|----------|--------|-------------|-----|
| 22 | CRv1-01 | MODERATE | modDataQuality_v2.1.bas | LogAction 4th arg = ElapsedSeconds() (Double) | Moved into message string |
| 23 | CRv1-02 | MODERATE | modReconciliation_v2.1.bas | LogAction 4th arg = ElapsedSeconds() (Double) | Moved into message string |
| 24 | CRv1-03 | MODERATE | modPDFExport_v2.1.bas | LogAction 4th arg = ElapsedSeconds() (Double) | Moved into message string |
| 25 | CRv1-04 | HIGH | date_format_unifier.py | detect_date_columns() missing day_first parameter | Added to signature + passed through |

### Phase 4: Pre-Delivery Code Review (2026-03-07)

| # | ID | Severity | Module | Description | Fix | Commit |
|---|-----|----------|--------|-------------|-----|--------|
| 26 | CR-01 | MODERATE | modReconciliation_v2.1.bas | LogAction instance #13 — Double as 4th arg | Moved into message string | 6818b01 |
| 27 | CR-02 | HIGH | modReconciliation_v2.1.bas | ValidateCrossSheet trendLastCol scanned row 1 not HDR_ROW_REPORT | Changed to HDR_ROW_REPORT | 6818b01 |
| 28 | CR-03 | HIGH | modReconciliation_v2.1.bas | FindFYCol scanned row 1 not HDR_ROW_REPORT | Changed to HDR_ROW_REPORT | 6818b01 |
| 29 | CR-04 | HIGH | modPDFExport_v2.1.bas | GetReportSheetList hardcoded to 7 sheets | Dynamic discovery | 6818b01 |
| 30 | CR-05 | HIGH | modDataSanitizer_v2.1.bas | rng not reset before SpecialCells in 2 functions | Added Set rng = Nothing | 6818b01 |
| 31 | CR-06 | HIGH | modAuditTools_v2.1.bas | rng not reset before SpecialCells in FindExternalLinks | Added Set rng = Nothing | 6818b01 |
| 32 | CR-07 | MODERATE | modDrillDown_v2.1.bas | xlSheetVeryHidden blocks hyperlinks | Changed to xlSheetHidden | b132885 |

### Phase 5: 5-Pass Bug Review (2026-03-11)

| # | ID | Severity | Module | Description | Fix |
|---|-----|----------|--------|-------------|-----|
| 33 | BR-01 | HIGH | modSplashScreen_v2.1.bas | Chr(9472) crashes VBA — only handles 0-255 | Changed to String(50, "=") |
| 34 | BR-02 | HIGH | modUTL_SplashScreen.bas | Chr(9472) in 2 locations — same crash | Changed to String(50, "=") |
| 35 | BR-03 | LOW | modSplashScreen_v2.1.bas | Unused SPLASH_BG and SPLASH_ACCENT constants | Removed |
