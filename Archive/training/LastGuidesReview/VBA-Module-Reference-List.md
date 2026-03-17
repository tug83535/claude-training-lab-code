# VBA Module Reference List — Demo File

## All 38 VBA Modules + 1 UserForm

Use this list to find the right module when you want to grab code from the demo file and adapt it to your own workbook using the CoPilot Prompt Guide.

---

### How to Export a Module from the Demo File

1. Open the demo file in Excel
2. Press **Alt+F11** to open the VBA Editor
3. In the left panel (Project Explorer), expand **Modules**
4. Right-click the module you want
5. Click **Export File**
6. Save the .bas file somewhere easy to find (like your Desktop)
7. Upload that .bas file to Copilot along with your own Excel file
8. Use **Prompt C1** from the CoPilot Prompt Guide to adapt the code

---

### Core System Modules

These modules run the demo file's infrastructure. You probably do not need to export these individually — they support the other modules behind the scenes.

| # | Module Name | What It Does |
|---|---|---|
| 1 | modConfig | Stores all settings: sheet names, column positions, colors, thresholds. Every other module reads from this one. |
| 2 | modFormBuilder | Builds the Command Center popup menu and routes all 62 actions to the correct macro. |
| 3 | modMasterMenu | Backup menu system — if the Command Center UserForm does not load, this InputBox version takes over. |
| 4 | modPerformance | TurboMode (turns off screen updating and calculations for speed) and timer utilities. |
| 5 | modLogger | Logs every action you run to a hidden audit trail sheet (VBA_AuditLog). |
| 6 | modNavigation | Table of Contents builder, Go Home button, keyboard shortcuts (Ctrl+Shift combos). |

---

### Data Quality and Cleaning

Great candidates to adapt for your own files — these scan and fix common Excel data problems.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 7 | modDataQuality | Scans for 6 common problems: blanks, text-stored numbers, formatting issues, duplicates, outliers, formula errors. Assigns a letter grade (A through F). | Any file — find out what is wrong with your data |
| 8 | modDataSanitizer | Fixes numeric-only problems: converts text-stored numbers, removes floating-point tails, fixes integer formats. Smart enough to skip dates, names, and IDs. | Files with numbers that do not calculate correctly |
| 9 | modDataGuards | Safety checks: validates assumptions are present, checks if drivers sum correctly, finds negative amounts, zero amounts, and suspiciously round numbers. | Financial models and budget files |

---

### Analysis and Reporting

These modules analyze data and produce reports. Very useful to adapt for your own P&L or financial files.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 10 | modVarianceAnalysis | Calculates month-over-month variance, flags items above 15% threshold, auto-generates plain-English commentary. | Budget vs actuals, monthly P&L reviews |
| 11 | modReconciliation | Runs 4 cross-sheet validation checks (column totals, date ranges, duplicate entries, missing records) and produces a PASS/FAIL report. | Reconciling GL data across sheets |
| 12 | modSensitivity | What-if sensitivity analysis on assumption drivers — shows how changes in one input affect outputs. | Financial models with driver-based assumptions |
| 13 | modForecast | Rolling forecast engine — extends your actuals with trend-based projections. | Monthly forecasting and planning |
| 14 | modTrendReports | Creates rolling 12-month views, reconciliation trend charts, and archives historical results. | Trend analysis and historical comparisons |

---

### Dashboards and Visualization

These build charts, dashboards, and visual reports automatically.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 15 | modDashboard | Creates dynamic charts, links chart titles to cells, and builds base dashboard components. | Any file that needs auto-generated charts |
| 16 | modDashboardAdvanced | Executive Dashboard with KPI scorecards, waterfall charts, small multiples grids, and product comparison views. | Executive-level reporting and presentations |
| 17 | modExecBrief | One-click executive brief — scans Revenue, Reconciliation, Assumptions, Products, and Workbook Health, then builds a styled one-page summary. | Quick executive summaries |

---

### Data Management and Import/Export

These handle bringing data in, exporting it out, and managing file structure.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 18 | modImport | Import pipeline for CSV and Excel files — maps columns, validates data, and loads into the workbook. | Importing GL data or other source files |
| 19 | modPDFExport | Batch exports multiple sheets to a professional PDF with headers, footers, and page numbers. | Creating polished PDF reports |
| 20 | modETLBridge | Connects Excel to Python scripts — triggers Python ETL jobs and imports the output back into Excel. | Teams using both Excel and Python |
| 21 | modMonthlyTabGenerator | Auto-creates monthly summary tabs (clones a template for each new month). Marks the next month on trend sheets. | Monthly financial reporting packages |

---

### Workbook Management and Utilities

Tools for organizing, searching, and managing your workbook.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 22 | modSearch | Cross-sheet search engine — searches every sheet for a keyword, highlights matches, caps at 200 results. | Finding data across a large workbook |
| 23 | modUtilities | 12 general utility macros: freeze panes, toggle gridlines, reset formatting, clear filters, and more. | Everyday Excel quality-of-life improvements |
| 24 | modSheetIndex | Creates a Home sheet with a clickable sheet index (hyperlinks to every tab). | Large workbooks with many tabs |
| 25 | modAuditTools | Workbook audit tools: change log, find/fix external links, audit hidden sheets, create masked copies (remove sensitive data). | Governance, compliance, and cleanup |
| 26 | modDrillDown | Reconciliation drill-down with hyperlinks, auto-populate checks, heatmaps, and golden file comparison. | Detailed reconciliation analysis |

---

### Scenario and Version Management

Save, compare, and restore different versions and scenarios of your data.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 27 | modScenario | Save, load, compare, and delete named scenarios (e.g., Best Case, Worst Case, Budget). | Budget scenarios and what-if planning |
| 28 | modVersionControl | Save workbook versions with timestamps, compare versions side-by-side, restore previous versions. | Tracking changes over time |
| 29 | modWhatIf | Live what-if scenario demo with 7 presets (Revenue +/-15%, AWS +25%, etc.) plus custom and restore. | Live demos and scenario presentations |

---

### Multi-Entity and Allocation

For companies with multiple business units or cost centers.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 30 | modAllocation | Cost allocation engine — distributes shared costs across departments or entities using configurable drivers. | Shared cost allocation |
| 31 | modConsolidation | Combines P&L data from multiple entities into a consolidated view, with intercompany elimination support. | Multi-entity financial consolidation |
| 32 | modAWSRecompute | AWS cloud cost allocation — validates and recalculates cloud spending distribution. | AWS cost management |

---

### Testing and Administration

Quality assurance and documentation tools.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 33 | modIntegrationTest | 18-test automated test suite + quick health check — validates that all modules work correctly together. | QA before sharing your file with others |
| 34 | modAdmin | Auto-generates workbook documentation and change management records. | Compliance and handoff documentation |

---

### Optional Add-Ins (Demo Enhancements)

These add polish and presentation value — great for demos and executive audiences.

| # | Module Name | What It Does | Best For |
|---|---|---|---|
| 35 | modTimeSaved | Calculates time saved — shows manual vs automated time for all 62 Command Center actions with annual savings. | ROI justification and demo impact |
| 36 | modSplashScreen | Branded welcome screen that appears when the workbook opens. | Professional first impressions |
| 37 | modProgressBar | Animated progress bar with percentage, ETA, and elapsed time for long-running macros. | Any macro that takes more than a few seconds |
| 38 | modDemoTools | Demo presentation tools: control sheet buttons, parameterized print areas, printable executive summaries. | Live demos and presentations |

---

### UserForm (Not a .bas File)

| # | Component | What It Does |
|---|---|---|
| 39 | frmCommandCenter | The popup Command Center menu — 62 buttons organized across 4 pages. This is a UserForm, not a module, so it cannot be exported as a .bas file. It is built automatically by modFormBuilder. |

---

## Which Modules Are Easiest to Adapt?

If you are new to this, start with these — they are the most universal and need the fewest changes to work on a different file:

1. **modDataQuality** — Works on almost any file. Just point it at your sheets.
2. **modSearch** — Searches any workbook. No customization needed.
3. **modUtilities** — General tools that work everywhere.
4. **modSheetIndex** — Builds a clickable index for any workbook.
5. **modPDFExport** — Exports any sheets to PDF.
6. **modDataSanitizer** — Fixes number formatting issues on any file.

For these, use **Prompt C1** from the CoPilot guide. Upload the .bas file and your Excel file, and Copilot will map everything for you.
