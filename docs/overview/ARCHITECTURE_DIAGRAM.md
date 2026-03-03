# KBT P&L Toolkit — Architecture Diagram

---

## System Overview

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                                                                             │
│                          USER INTERFACE LAYER                                │
│                                                                             │
│    ┌──────────────────────┐  ┌────────────────┐  ┌─────────────────────┐   │
│    │  frmCommandCenter    │  │  InputBox       │  │  Keyboard Shortcuts │   │
│    │  (UserForm)          │  │  (Fallback)     │  │                     │   │
│    │                      │  │                 │  │  Ctrl+Shift+M Menu  │   │
│    │  - 14 categories     │  │  3-page menu    │  │  Ctrl+Shift+H Home  │   │
│    │  - 62 actions        │  │  62 items       │  │  Ctrl+Shift+J Jump  │   │
│    │  - Search filter     │  │  Type number    │  │  Ctrl+Shift+R Check │   │
│    │  - Double-click run  │  │  Press OK       │  │                     │   │
│    └──────────┬───────────┘  └───────┬─────────┘  └──────────┬──────────┘   │
│               │                      │                       │              │
│               └──────────────────────┼───────────────────────┘              │
│                                      ▼                                      │
│                    modFormBuilder.ExecuteAction(n)                           │
│                    modMasterMenu.ExecuteMenuAction(n)                        │
│                                                                             │
└───────────────────────────────┬─────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                                                                             │
│                          FEATURE MODULE LAYER                                │
│                        (32 VBA modules, ~8,500 LOC)                          │
│                                                                             │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────────────────┐ │
│  │ MONTHLY OPS     │  │ ANALYSIS        │  │ DATA QUALITY                │ │
│  │ #1-4            │  │ #5-6            │  │ #7-9                        │ │
│  │                 │  │                 │  │                             │ │
│  │ MonthlyTabGen   │  │ Sensitivity     │  │ DataQuality                 │ │
│  │  GenerateAll    │  │  RunAnalysis    │  │  ScanAll                    │ │
│  │  DeleteAll      │  │                 │  │  FixTextNumbers             │ │
│  │  GenerateNext   │  │ VarianceAnalysis│  │  FixDuplicates              │ │
│  │                 │  │  RunAnalysis    │  │                             │ │
│  │ Reconciliation  │  │  Commentary     │  │                             │ │
│  │  RunAllChecks   │  │                 │  │                             │ │
│  │  CrossSheet     │  │                 │  │                             │ │
│  │  ExportResults  │  │                 │  │                             │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────────────────┘ │
│                                                                             │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────────────────┐ │
│  │ REPORTING       │  │ UTILITIES       │  │ DATA & IMPORT               │ │
│  │ #10-12          │  │ #13-16          │  │ #17                         │ │
│  │                 │  │                 │  │                             │ │
│  │ PDFExport       │  │ Navigation      │  │ Import                      │ │
│  │  ReportPackage  │  │  TOC Refresh    │  │  ImportDataPipeline         │ │
│  │  SingleSheet    │  │  QuickJump      │  │                             │ │
│  │                 │  │  GoHome         │  │                             │ │
│  │ Dashboard       │  │                 │  │                             │ │
│  │  BuildCharts    │  │ AWSRecompute    │  │                             │ │
│  │  Executive      │  │  ValidateRecalc │  │                             │ │
│  │  Waterfall      │  │                 │  │                             │ │
│  │  ProductComp    │  │                 │  │                             │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────────────────┘ │
│                                                                             │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────────────────┐ │
│  │ FORECASTING     │  │ SCENARIOS       │  │ ALLOCATION                  │ │
│  │ #18-19          │  │ #20-23          │  │ #24-25                      │ │
│  │                 │  │                 │  │                             │ │
│  │ Forecast        │  │ Scenario        │  │ Allocation                  │ │
│  │  Rolling        │  │  Save/Load      │  │  RunEngine                  │ │
│  │  AppendToTrend  │  │  Compare/Delete │  │  Preview                    │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────────────────┘ │
│                                                                             │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────────────────┐ │
│  │ CONSOLIDATION   │  │ VERSION CONTROL │  │ GOVERNANCE                  │ │
│  │ #26-30          │  │ #31-35          │  │ #36-40                      │ │
│  │                 │  │                 │  │                             │ │
│  │ (Multi-entity)  │  │ VersionControl  │  │ Admin                       │ │
│  │  AddEntity      │  │  Save/Compare   │  │  AutoDocumentation          │ │
│  │  Consolidate    │  │  Restore/List   │  │  ChangeManagement           │ │
│  │  Eliminations   │  │                 │  │  CRs + Status + Summary     │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────────────────┘ │
│                                                                             │
│  ┌─────────────────┐  ┌──────────────────────────────────────────────────┐ │
│  │ ADMIN & TEST    │  │ ADVANCED  #46-50                                │ │
│  │ #41-45          │  │                                                  │ │
│  │                 │  │ VarianceCommentary | CrossSheetValidation        │ │
│  │ Logger          │  │ ExecutiveMode      | ForceRecalc | About        │ │
│  │  View/Export    │  │                                                  │ │
│  │  Clear          │  │ ConditionalFormat  | EmailSummary | Validation  │ │
│  │                 │  │ Snapshot | Search  | Formatting  | Setup        │ │
│  │ IntegrationTest │  │                                                  │ │
│  │  Full/Quick     │  │ (Supporting modules not directly on menu)       │ │
│  └─────────────────┘  └──────────────────────────────────────────────────┘ │
│                                                                             │
└───────────────────────────────┬─────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                                                                             │
│                          FOUNDATION LAYER                                    │
│                        (4 modules, ~800 LOC)                                 │
│                                                                             │
│  ┌──────────────────────────┐  ┌──────────────────────────────────────────┐ │
│  │ modConfig                │  │ modPerformance                           │ │
│  │                          │  │                                          │ │
│  │ Constants:               │  │ TurboOn()   - Disable screen refresh    │ │
│  │  - Sheet names (13)      │  │ TurboOff()  - Restore screen refresh    │ │
│  │  - Layout offsets (5)    │  │ StartTimer()- Begin timing              │ │
│  │  - Product/Dept lists    │  │ Elapsed()   - Read elapsed seconds      │ │
│  │  - Colors (10+)          │  │ ForceRecalc()                           │ │
│  │  - Thresholds (4)        │  │ StatusBar() - Show progress             │ │
│  │                          │  │                                          │ │
│  │ Helpers:                 │  ├──────────────────────────────────────────┤ │
│  │  - SheetExists()         │  │ modLogger                               │ │
│  │  - SafeDeleteSheet()     │  │                                          │ │
│  │  - StyleHeader()         │  │ LogAction() - Write audit trail entry   │ │
│  │  - AutoFitWithMax()      │  │ ViewLog()   - Navigate to log sheet     │ │
│  │  - WriteSummaryRow()     │  │ ExportLog() - Save to CSV               │ │
│  │  - CenterHeader()        │  │ ClearLog()  - Reset log                 │ │
│  │                          │  │ HideLog()   - Toggle visibility         │ │
│  └──────────────────────────┘  └──────────────────────────────────────────┘ │
│  ┌───────────────────────────────────────────────────────────────────────┐  │
│  │ ThisWorkbook (Events)                                                 │  │
│  │  Workbook_Open  → AssignShortcuts, initialize state                  │  │
│  │  BeforeClose    → Log session end                                    │  │
│  └───────────────────────────────────────────────────────────────────────┘  │
│                                                                             │
└───────────────────────────────┬─────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                                                                             │
│                          DATA LAYER                                          │
│                                                                             │
│  KeystoneBenefitTech_PL_Model.xlsx  (13 sheets)                             │
│                                                                             │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │ CrossfireHiddenWorksheet (GL)    │ 510 rows │ 7 cols │ Row 1=HDR  │    │
│  │ Assumptions                      │  33 rows │ 4 cols │ Row 5=HDR  │    │
│  │ Data Dictionary                  │  54 rows │ 5 cols │             │    │
│  │ AWS Allocation                   │  42 rows │ 6 cols │             │    │
│  │ Report-->                        │  22 rows │ 6 cols │ TOC/Home   │    │
│  │ P&L - Monthly Trend              │  44 rows │ 18 cols│ Row 4=HDR  │    │
│  │ Product Line Summary             │  80 rows │ 18 cols│ Row 4=HDR  │    │
│  │ Functional P&L - Monthly Trend   │ 147 rows │ 18 cols│ Row 4=HDR  │    │
│  │ Functional P&L Summary - Jan 25  │  37 rows │ 5 cols │ Row 4=HDR  │    │
│  │ Functional P&L Summary - Feb 25  │  37 rows │ 5 cols │ Row 4=HDR  │    │
│  │ Functional P&L Summary - Mar 25  │  37 rows │ 5 cols │ Row 4=HDR  │    │
│  │ US January 2025 Natural P&L      │  77 rows │ 5 cols │             │    │
│  │ Checks                           │  13 rows │ 5 cols │ Row 4=HDR  │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                             │
│  Generated sheets (created by toolkit commands):                            │
│  Functional P&L Summary - Apr 25 .. Dec 25  (Command 1)                    │
│  Dashboard / Executive Dashboard             (Command 12)                   │
│  Variance Analysis / Variance Commentary     (Commands 6, 46)              │
│  Data Quality Report                         (Command 7)                    │
│  Cross-Sheet Validation                      (Command 47)                   │
│  Sensitivity Analysis                        (Command 5)                    │
│  Search Results                              (Commands via modSearch)       │
│  Allocation Output                           (Command 24)                   │
│  Integration Test Report                     (Commands 44, 45)             │
│  Tech Documentation                          (Command 36)                   │
│  Change Management Log                       (Commands 37-40)              │
│  VBA_AuditLog                                (auto, all commands)           │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘


┌─────────────────────────────────────────────────────────────────────────────┐
│                                                                             │
│                    PYTHON ANALYTICS LAYER (Optional, External)                │
│                              (13 scripts, ~5,200 LOC)                        │
│                                                                             │
│  ┌──────────────────────────────────────────────────────────────────────┐   │
│  │  pnl_runner.py  ←── Unified CLI entry point                         │   │
│  │  ┌──────────┬─────────────┬───────────┬───────────┬──────────────┐  │   │
│  │  │ dashboard│ month-end   │ forecast  │ allocate  │ snapshot     │  │   │
│  │  │ match    │ test        │ config    │           │              │  │   │
│  │  └──────────┴─────────────┴───────────┴───────────┴──────────────┘  │   │
│  └──────────────────────────────────────────────────────────────────────┘   │
│                                                                             │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │ pnl_config.py     Shared constants, PnLBase class, utilities       │    │
│  │ pnl_dashboard.py  Streamlit interactive dashboard (web UI)         │    │
│  │ pnl_month_end.py  Month-end close checklist (6 check categories)   │    │
│  │ pnl_allocation_simulator.py  What-if allocation scenarios          │    │
│  │ pnl_forecast.py   Statistical forecasting (SMA, ETS, trend)       │    │
│  │ pnl_snapshot.py   Snapshot management                              │    │
│  │ pnl_ap_matcher.py Fuzzy vendor matching (thefuzz)                  │    │
│  │ pnl_cli.py        Click-based CLI interface                        │    │
│  │ pnl_tests.py      pytest suite (116 tests, 17 classes)            │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                             │
│  Dependencies: pandas, numpy, openpyxl, matplotlib, streamlit, plotly,     │
│                statsmodels, scikit-learn, thefuzz, click, pytest            │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘


┌─────────────────────────────────────────────────────────────────────────────┐
│                                                                             │
│                    SQL LAYER (Optional, Portable Data Store)                  │
│                                                                             │
│  ┌───────────────────┐  ┌──────────────────────┐  ┌─────────────────────┐  │
│  │ staging.sql        │  │ transformations.sql   │  │ validations.sql     │  │
│  │                    │  │                       │  │                     │  │
│  │ GL staging table   │  │ Allocation pivot      │  │ Referential checks  │  │
│  │ Dimension tables   │  │ Product summary view  │  │ Orphan detection    │  │
│  │ Date normalization │  │ Dept summary view     │  │ Balance validation  │  │
│  │ Dedup logic        │  │ MoM variance calc     │  │ Completeness scan   │  │
│  └───────────────────┘  └──────────────────────┘  └─────────────────────┘  │
│                                                                             │
│  pnl_enhancements.sql  (S1-S5: Budget vs Actual, Audit Trail, Rolling 12M, │
│                          Vendor Calendar, Recon Queries)                     │
│                                                                             │
│  Engine: SQLite 3.x (portable, zero-install)                                │
│  Power Query M equivalents included in each SQL file as comments            │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Data Flow

```
  Accounting System
        │
        ▼ (CSV/Excel extract)
  ┌─────────────┐     ┌──────────────┐
  │ modImport   │────▶│ GL Sheet     │
  │ (Command 17)│     │ (510+ rows)  │
  └─────────────┘     └──────┬───────┘
                             │
                ┌────────────┼────────────┐
                ▼            ▼            ▼
         modDataQuality  modReconciliation  modAllocation
         (scan + fix)    (PASS/FAIL)       (Dept × Product)
                │            │                │
                ▼            ▼                ▼
         DQ Report     Checks Sheet    Allocation Output
                             │
                ┌────────────┼────────────┐
                ▼            ▼            ▼
         modVariance     modDashboard   modForecast
         (MoM delta)     (3+ charts)   (SMA/ETS)
                │            │                │
                ▼            ▼                ▼
         Variance Sheet  Dashboard     Trend Update
         + Commentary
                │
                ▼
         modPDFExport ───▶ Report Package (PDF)
         modEmailSummary ──▶ Executive Summary
```

---

## Module Dependency Map

```
  ThisWorkbook_Events
        │
        ├──▶ modConfig (constants, helpers)
        │       ▲
        │       │ (every module depends on modConfig)
        │       │
        ├──▶ modPerformance (TurboOn/Off)
        │       ▲
        │       │ (every public sub calls TurboOn/Off)
        │       │
        ├──▶ modLogger (LogAction)
        │       ▲
        │       │ (every public sub calls LogAction)
        │       │
        └──▶ modNavigation.AssignShortcuts
                │
                ▼
        modFormBuilder ◀──▶ modMasterMenu
        (ExecuteAction)     (ExecuteMenuAction)
                │
                ├──▶ modMonthlyTabGenerator
                ├──▶ modReconciliation
                ├──▶ modDataQuality
                ├──▶ modVarianceAnalysis
                ├──▶ modSensitivity
                ├──▶ modDashboard
                ├──▶ modPDFExport
                ├──▶ modAWSRecompute
                ├──▶ modImport
                ├──▶ modForecast
                ├──▶ modScenario
                ├──▶ modAllocation
                ├──▶ modSearch
                ├──▶ modSnapshot
                ├──▶ modConditionalFormat
                ├──▶ modEmailSummary
                ├──▶ modValidation
                ├──▶ modAdmin
                ├──▶ modIntegrationTest
                ├──▶ modRefresh
                ├──▶ modFormatting
                └──▶ modSetup
```
