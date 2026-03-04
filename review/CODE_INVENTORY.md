# Complete Code Inventory

**iPipeline P&L Demo File + Universal Toolkit**
*Last updated: 2026-03-04*

---

## Part 1: Excel Demo File VBA Modules (33 modules)

All files in `vba/` — these live inside the demo `.xlsm` workbook.

---

### modConfig — Workbook Constants & Configuration
| Sub/Function | Purpose |
|---|---|
| GetProducts | Return product array (iGO, Affirm, InsureSight, DocFast) |
| GetDepartments | Return department array |
| GetMonths | Return month array (Jan-Dec) |
| GetMonthSheetNames | Return monthly summary sheet names for each month |
| GetSheet | Safe sheet reference by name |
| SheetExists | Check if a sheet exists by name |
| SafeDeleteSheet | Delete a sheet by name if it exists, suppress alerts |
| StyleHeader | Write header row with navy background + white bold text |
| LastRow | Return last non-empty row in a given column |
| LastCol | Return last non-empty column in a given row |
| FindColByHeader | Scan header row for keyword, return column number |
| FindRowByLabel | Scan column A for keyword, return row number |
| SafeNum | Convert any value to Double safely, return 0 on error |
| SafeStr | Convert any value to trimmed String safely, return "" on error |
| AddNamedRanges | Create revenue-share named ranges on Assumptions sheet |

---

### modNavigation — Sheet Navigation & Table of Contents
| Sub/Function | Purpose |
|---|---|
| RefreshTableOfContents | Rebuild hyperlinks on Report--> sheet |
| GoHome | Navigate to Report--> sheet |
| QuickJump | Show sheet list and jump to selection |
| AssignShortcuts | Bind keyboard shortcuts (Ctrl+Shift combos) |
| ToggleExecutiveMode | Hide utility/working sheets, show report sheets only |
| ClearShortcuts | Unbind all custom shortcuts on workbook close |

---

### modPerformance — TurboMode & Timer Utilities
| Sub/Function | Purpose |
|---|---|
| TurboOn | Suppress UI refreshes, disable events, manual calc |
| TurboOff | Restore all settings to pre-TurboOn state |
| ForceRecalc | Full workbook recalculation |
| ElapsedSeconds | Return seconds since TurboOn was called (handles midnight) |
| UpdateStatus | Write message to status bar with optional percentage |

---

### modFormBuilder — UserForm Command Center Builder & Launcher
| Sub/Function | Purpose |
|---|---|
| LaunchCommandCenter | Show UserForm or fallback to InputBox |
| BuildCommandCenter | Programmatically create the UserForm |
| ExecuteAction | Central router for all 62 Command Center actions |
| ShowAbout | Display toolkit info dialog |
| CreateFormManually | Instructions for manual form creation |
| GetFormCodeForManual | Print form code to Immediate Window for copy/paste |
| GetFormInstallGuide | Print complete step-by-step install guide |

---

### modMasterMenu — Central Command Panel (InputBox Fallback)
| Sub/Function | Purpose |
|---|---|
| ShowMasterMenu | 4-page InputBox menu listing all 62 actions |

---

### modLogger — Runtime Action Logger
| Sub/Function | Purpose |
|---|---|
| LogAction | Write one entry to audit log (module, procedure, message, status) |
| ClearLog | Erase all log data (keeps header row) |
| ExportLog | Copy audit log to new workbook for archiving |
| ViewLog | Make audit log visible and navigate to it |
| GetLogSheet | Return VBA_AuditLog worksheet, creating if needed |

---

### modMonthlyTabGenerator — Auto-Generate Monthly Summary Tabs
| Sub/Function | Purpose |
|---|---|
| GenerateMonthlyTabs | Create all 9 tabs (Apr-Dec) from Mar template |
| GenerateNextMonthOnly | Create just the next missing monthly tab |
| AddNextMonthToModel | Calendar-aware next-month prep (reads today's date) |
| DeleteGeneratedTabs | Remove all auto-generated tabs (Apr-Dec) |
| TestUpdateHeaderText | Public wrapper for testing header text updates |

---

### modPDFExport — Professional Batch PDF Export
| Sub/Function | Purpose |
|---|---|
| ExportReportPackage | Export all report sheets to single PDF |
| ExportSingleSheet | Export active sheet to PDF |

---

### modReconciliation — Automated Reconciliation & Validation
| Sub/Function | Purpose |
|---|---|
| RunAllChecks | Evaluate existing Checks sheet PASS/FAIL formulas |
| ExportCheckResults | Write results to timestamped text file |
| ValidateCrossSheet | Compute validations from raw data across all sheets |

---

### modVarianceAnalysis — Variance Detection, Reporting & Commentary
| Sub/Function | Purpose |
|---|---|
| RunVarianceAnalysis | Month-over-month comparison of two functional P&L sheets |
| GenerateCommentary | Auto-generate English narratives for top variances |

---

### modDataQuality — Data Cleaning Scanner & Fixer
| Sub/Function | Purpose |
|---|---|
| ScanAll | Full workbook data quality scan |
| FixTextNumbers | Convert ONLY pre-flagged text-stored numbers |
| FixDuplicates | Remove duplicate rows from GL staging data |

---

### modDataSanitizer — Numeric-Only Data Sanitizer
| Sub/Function | Purpose |
|---|---|
| RunFullSanitize | All 3 fixes in sequence (text-numbers, FP tails, integer format) |
| PreviewSanitizeChanges | Dry-run report showing what WOULD change (no edits) |
| FixFloatingPointTails | Fix floating-point noise on every visible sheet |
| ConvertTextStoredNumbers | Convert text-numbers to real numbers (all sheets) |
| NormalizeIntegerFormats | Apply 2dp display format to whole-number amounts |

---

### modDataGuards — Data Validation Safety Checks
| Sub/Function | Purpose |
|---|---|
| ValidateAssumptionsPresence | Block if any key driver cell is blank |
| CheckSumOfDrivers | Validate revenue share percentages sum to 100% |
| FindNegativeAmounts | Flag GL rows where Amount < 0 |
| FindZeroAmounts | Flag GL rows where Amount = 0 |
| FindSuspiciousRoundNumbers | Flag GL amounts that are exact multiples of 1000 |

---

### modDashboard — Dynamic Dashboard & Chart Generation
| Sub/Function | Purpose |
|---|---|
| BuildDashboard | 3 charts on Report--> (revenue, margin, mix) |
| RefreshDashboard | Recalculate and refresh existing charts |
| CreateExecutiveDashboard | KPI cards + summary table on dedicated sheet |
| WaterfallChart | Revenue-to-Net-Income waterfall bridge chart |
| ProductComparison | Side-by-side product metrics + ranking |
| ReformatChartsAndVisuals | Reflow Charts & Visuals sheet into clean grid |

---

### modDemoTools — Demo Presentation & Print Tools
| Sub/Function | Purpose |
|---|---|
| AddControlSheetButtons | Add clickable macro buttons to Report--> sheet |
| SetParameterizedPrintArea | Set print area by selected month/product |
| CreatePrintableExecSummary | Build one-page print layout for CFO |

---

### modSensitivity — What-If Sensitivity Analysis
| Sub/Function | Purpose |
|---|---|
| RunSensitivityAnalysis | Calculate sensitivity for Assumptions drivers |

---

### modForecast — Rolling Forecast & Trend Append
| Sub/Function | Purpose |
|---|---|
| RollingForecast | Generate forecast for remaining months |
| AppendToTrend | Copy a monthly summary into P&L trend sheet |

---

### modScenario — Scenario Management
| Sub/Function | Purpose |
|---|---|
| SaveScenario | Snapshot current Assumptions to Scenarios sheet |
| LoadScenario | Restore Assumptions from a saved scenario |
| CompareScenarios | Side-by-side comparison of saved scenarios |
| DeleteScenario | Remove a saved scenario |

---

### modSearch — Cross-Sheet Search Engine
| Sub/Function | Purpose |
|---|---|
| SearchAll | Search all visible sheets for a keyword |
| SearchAndNavigate | Search and jump to selected result |
| SearchCurrentSheet | Search active sheet only, highlight matches |

---

### modAllocation — Cost Allocation Engine
| Sub/Function | Purpose |
|---|---|
| RunAllocationEngine | Allocate shared costs to products |
| AllocationPreview | What-if preview with modified shares |

---

### modAWSRecompute — AWS Allocation Recalculation
| Sub/Function | Purpose |
|---|---|
| ValidateAndRecalcAWS | Validate AWS shares and recalculate allocations |

---

### modConsolidation — Multi-Entity P&L Consolidation
| Sub/Function | Purpose |
|---|---|
| ShowConsolidationMenu | Display consolidation status and options |
| AddEntity | Load P&L data from external workbook |
| GenerateConsolidated | Build consolidated P&L from loaded entities |
| ListEntities | Show all loaded entities |
| AddElimination | Record an intercompany elimination entry |

---

### modVersionControl — Workbook Version Management
| Sub/Function | Purpose |
|---|---|
| ShowVersionMenu | Display version control status |
| SaveVersion | Save current state as new version |
| CompareVersions | Show version history for comparison |
| RestoreVersion | Open previous version for manual restore |
| ListVersions | Show all saved versions |

---

### modAdmin — Auto-Documentation & Change Management
| Sub/Function | Purpose |
|---|---|
| GenerateDocumentation | Auto-document workbook structure |
| ShowChangeMenu | Display change management status |
| AddChangeRequest | Log new change request |
| UpdateChangeStatus | Update status of existing change request |
| ChangeManagementSummary | Summary report of all change requests |

---

### modIntegrationTest — Integration Testing & Health Check
| Sub/Function | Purpose |
|---|---|
| RunFullTest | Run all 18 integration tests |
| QuickHealthCheck | Verify sheets exist and key values are valid |

---

### modImport — GL Data Import Pipeline
| Sub/Function | Purpose |
|---|---|
| ImportDataPipeline | Import CSV/Excel data into GL staging sheet |

---

### modUtilities — Sheet & Workbook Utility Macros
| Sub/Function | Purpose |
|---|---|
| DeleteBlankRows | Delete completely blank rows in selection |
| UnhideAllSheets | Make every worksheet visible |
| SortSheetsAlphabetically | Reorder all tabs A-Z by name |
| ToggleFreezePanes | Toggle freeze panes on/off (freezes at B2) |
| ConvertToValues | Replace formulas in selection with static values |
| AutoFitAllColumns | AutoFit every column on active sheet |
| ProtectAllSheets | Password-protect every worksheet |
| UnprotectAllSheets | Remove password protection from every worksheet |
| FindReplaceAllSheets | Find & replace text across every worksheet |
| HighlightHardcodedNumbers | Change font color to blue for non-formula numbers |
| TogglePresentationMode | Hide/show gridlines, headings, formula bar |
| UnmergeAndFillDown | Unmerge selection and fill blanks from above |
| ClearStatusBar | Restore default Excel status bar |

---

### modDrillDown — Reconciliation Drill & Comparison Tools
| Sub/Function | Purpose |
|---|---|
| AddReconciliationDrillLinks | Hyperlinks from Checks rows to GL data |
| AutoPopulateReconciliationChecks | Recalculate + verify named ranges |
| ApplyReconciliationHeatmap | Color Checks tab by variance size |
| RunGoldenFileCompare | Compare current P&L to saved baseline |

---

### modTrendReports — Trend Views & Historical Archiving
| Sub/Function | Purpose |
|---|---|
| CreateRolling12MonthView | Build dynamic rolling 12-month P&L |
| CreateReconciliationTrendChart | Chart PASS/FAIL counts over time |
| ArchiveReconciliationResults | Save dated snapshot of Checks tab |

---

### modETLBridge — Python ETL Integration Bridge
| Sub/Function | Purpose |
|---|---|
| TriggerETLLocally | Run kbt_etl_pipeline.py via Windows Shell |
| ImportETLOutput | Load cleaned Excel output into workbook |

---

### modAuditTools — Workbook Audit & Link Management
| Sub/Function | Purpose |
|---|---|
| AppendChangeLogEntry | Add timestamped entry to change log |
| FindExternalLinks | Scan for cells referencing external workbooks |
| FixExternalLinks | Replace external links with current values |
| AuditHiddenSheets | List all hidden/very hidden sheets |
| CreateMaskedCopy | Create anonymized copy for sharing |
| ExportErrorSummaryClipboard | Copy error summary to clipboard |

---

### modSheetIndex — Home Sheet & Sheet Index Builder
| Sub/Function | Purpose |
|---|---|
| CreateHomeSheet | Build Home sheet with Command Center launch button |
| ListAllSheetsWithLinks | Build/update Sheet Index with clickable hyperlinks |

---

## Part 2: Universal Toolkit Modules (8 modules, ~85 tools)

All files in `UniversalToolsForAllFiles/vba/` — standalone tools for ANY Excel file.

---

### modUTL_Audit — Workbook Audit & Compliance (8 tools)
| Sub | Purpose |
|---|---|
| ExternalLinkFinder | Scan for cells referencing external workbooks, create report |
| CircularReferenceDetector | Find all circular references in workbook |
| WorkbookErrorScanner | List every error cell (#DIV/0!, #REF!, etc.) across all sheets |
| DataQualityScorecard | Column-by-column summary (blanks, errors, duplicates, types) |
| NamedRangeAuditor | Report all named ranges and flag broken references |
| DataValidationChecker | Scan data validation dropdowns for broken source ranges |
| InconsistentFormulasAuditor | Flag cells where formula differs from column majority |
| ExternalLinkSeveranceProtocol | Replace external links with values, save formulas as comments |

---

### modUTL_DataCleaning — Data Cleanup & Repair (12 tools)
| Sub | Purpose |
|---|---|
| UnmergeAndFillDown | Unmerge cells and fill values down |
| FillBlanksDown | Fill blank cells with value from cell above |
| ConvertTextToNumbers | Fix cells storing numbers as text |
| RemoveLeadingTrailingSpaces | Trim invisible spaces from text cells |
| DeleteBlankRows | Remove completely empty rows from active sheet |
| ReplaceErrorValues | Replace #N/A, #REF!, etc. with blank or custom value |
| HighlightDuplicateRows | Color duplicate rows yellow (no delete) |
| RemoveDuplicateRows | Delete duplicate rows based on key column |
| MultiReplaceDataCleaner | Batch find-and-replace using mapping table |
| FormulaToValueHardcoder | Convert all formulas in selection to static values |
| PhantomHyperlinkPurger | Remove all embedded hyperlinks from active sheet |
| ConvertNumbersToWords | Translate numeric values to written text |

---

### modUTL_Finance — Finance-Specific Tools (14 tools)
| Sub | Purpose |
|---|---|
| DuplicateInvoiceDetector | Flag duplicate invoices (vendor + amount + date within 3 days) |
| AutoBalancingGLValidator | Sum debit/credit, flag imbalance, optionally insert plug |
| TrialBalanceChecker | Verify total debits equal total credits |
| JournalEntryValidator | Group JEs by number, check each balances |
| FluxAnalysis | Compare two columns, flag changes exceeding threshold |
| APAgingSummaryGenerator | Bucket AP invoices by days overdue |
| ARAgingSummaryGenerator | Bucket AR invoices by days outstanding |
| AgingBucketCalculator | Add aging bucket column (0-30, 31-60, 61-90, 90+) |
| VarianceAnalysisTemplate | Add $ Variance and % Variance columns |
| QuickCorkscrewBuilder | Build standard roll-forward schedule |
| FinancialPeriodRollForward | Update month-end headers, clear input cells |
| MultiCurrencyConsolidationAggregator | Consolidate currencies using FX rate table |
| RatioAnalysisDashboard | Calculate key financial ratios (margins, ROE, ROA) |
| GeneralLedgerJournalMapper | Transform trial balance into JE upload template |

---

### modUTL_Formatting — Number & Layout Formatting (9 tools)
| Sub | Purpose |
|---|---|
| AutoFitAllColumnsRows | Auto-fit every column and row |
| FreezeTopRowAllSheets | Apply freeze panes to row 1 on every sheet |
| NumberFormatStandardizer | Apply #,##0.00 to all numeric cells |
| CurrencyFormatStandardizer | Apply $#,##0.00 to selected range |
| DateFormatStandardizer | Normalize all dates to MM/DD/YYYY |
| HighlightNegativesRed | Conditional formatting for negative numbers in red |
| FinancialNumberFormattingSuite | Menu-driven format choice (Accounting, 000s, %, etc.) |
| ConditionalFormatPurger | List and remove all conditional formatting rules |
| PrintHeaderFooterStandardizer | Apply consistent headers/footers across all sheets |

---

### modUTL_WorkbookMgmt — Workbook-Level Operations (15 tools)
| Sub | Purpose |
|---|---|
| UnhideAllSheetsRowsColumns | Make every hidden sheet, row, column visible |
| ExportAllSheetsCombinedPDF | Combine all visible sheets into single PDF |
| FindReplaceAcrossAllSheets | Global find-and-replace across every sheet |
| SearchAcrossAllSheets | Find any value across every sheet (200-result cap) |
| MultiSheetBatchRenamer | Replace text in all sheet tab names at once |
| SortWorksheetsAlphabetically | Reorder all sheet tabs A-Z |
| CreateTableOfContents | Generate clickable index sheet linking to every worksheet |
| ProtectAllSheets | Apply password protection to every sheet |
| UnprotectAllSheets | Remove protection from every sheet |
| LockAllFormulaCells | Lock formula cells, leave input cells editable |
| ExportActiveSheetPDF | Save current sheet as PDF |
| ExportAllSheetsIndividualPDFs | Save each sheet as its own PDF |
| ResetAllFilters | Clear all AutoFilter criteria across every sheet |
| BuildDistributionReadyCopy | Create clean copy (values only, metadata stripped) |
| WorkbookHealthCheck | Full diagnostic report (size, errors, links, formulas) |

---

### modUTL_Branding — iPipeline Brand Styling (2 tools)
| Sub | Purpose |
|---|---|
| ApplyiPipelineBranding | Style headers, alternating rows, totals with iPipeline brand colors |
| SetiPipelineThemeColors | Set workbook theme colors to iPipeline brand palette |

---

### modUTL_DataSanitizer — Enhanced Numeric Sanitizer (4 tools)
| Sub | Purpose |
|---|---|
| RunFullSanitize | All 3 fixes in one click (text-numbers, FP tails, integer format) |
| PreviewSanitizeChanges | Dry-run report showing what WOULD change (no edits) |
| FixFloatingPointTails | Fix floating-point noise (9412.300000001 -> 9412.30) |
| ConvertTextStoredNumbers | Convert text-stored numbers to real numbers |

---

### modUTL_SheetTools — Worksheet Management (3 tools)
| Sub | Purpose |
|---|---|
| ListAllSheetsWithLinks | Create sheet index with clickable hyperlinks and visibility status |
| TemplateCloner | Pick any sheet, specify count (1-50), get instant clones |
| GenerateUniqueCustomerIDs | Assign sequential IDs to blank cells (CUST-00001 format) |

---

## Summary

| Category | Modules | Tools/Subs |
|---|---|---|
| Demo File VBA | 33 | ~130 |
| Universal Toolkit | 8 | ~85 |
| **Total** | **41** | **~215** |

---

*Generated 2026-03-04 | iPipeline P&L Reporting & Allocation Model*
