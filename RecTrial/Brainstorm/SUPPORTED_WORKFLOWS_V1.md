# Supported Workflows — V1

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Date:** 2026-04-28
**Purpose:** The official first-recommendation list for coworkers adopting the Finance Automation Toolkit. These are the 7 workflows to lead with — in documentation, in the launch message, in the SharePoint Quick Reference Card, and in the Start Here guide. Everything outside this list is "advanced / discoverable / not the first recommended path."

---

## How this list works

When coworkers ask "where do I start?" the answer is one of these 7 workflows. Each one:
- Solves a real Finance & Accounting problem in under 15 minutes
- Uses tools that are already built and tested
- Has a sample file they can practice on before touching real data
- Has clear adoption guidance so they can use it on their own files, not just the demo

The full toolkit contains ~140 VBA tools and 28 Python scripts. All of it stays inside `Sample_Quarterly_ReportV2.xlsm` and the Command Center, discoverable for power users. But none of that is the doorway for V1 adoption.

---

## VBA workflows — runs inside Excel, no Python required

These workflows work in any Excel file. Import one VBA module and run from Alt+F8 (or the Command Center if using the demo workbook). No Python installation needed.

---

### Workflow 1 — Clean a messy Excel export

**What it does:** Fixes the three most common problems that appear when data comes from other systems — numbers stored as text (so formulas won't add them up), floating-point noise (numbers like 9412.300000000001 instead of 9412.30), and integer display issues. Also handles unmerged cells, blank rows, leading/trailing spaces, and other common export garbage.

**Why it matters:** Almost every exported report from a billing system, ERP, or database has at least one of these problems. Before you can analyze or summarize anything, the data has to be clean.

**Primary module:** `RecTrial\UniversalToolkit\vba\modUTL_DataSanitizer.bas`

**Key tools to run (in order):**
1. `PreviewSanitizeChanges` — dry run, no changes made. Shows exactly what will be fixed. **Always run this first** to understand the scope before committing.
2. `RunFullSanitize` — applies all three fixes in one click.
3. `FixFloatingPointTails` — standalone if you only want to fix FP noise.
4. `ConvertTextStoredNumbers` — standalone if you only want to fix text-stored numbers.

**Companion module for deeper cleaning:** `RecTrial\UniversalToolkit\vba\modUTL_DataCleaning.bas`
- `UnmergeAndFillDown` — unmerge cells and fill the value down all rows (a very common export problem)
- `FillBlanksDown` — fill blank cells with the value from the cell above
- `DeleteBlankRows` — remove all blank rows from a selected range
- `RemoveLeadingTrailingSpaces` — strip invisible spaces from a column
- `RemoveDuplicateRows` — remove exact duplicate rows
- `ReplaceErrorValues` — replace #N/A, #REF!, #DIV/0! etc. with a chosen value

**Python alternative (optional):** `RecTrial\UniversalToolkit\python\ZeroInstall\sanitize_dataset.py` — CSV-based version for very large exports or scripted batch runs. Stdlib only, no pip required.

**Demo sample file:** `RecTrial\SampleFile\SampleFileV2\` (V3 universal tools demo file)

**Time-to-value:** Under 2 minutes from import to clean data.

**Prerequisites:** Import `modUTL_DataSanitizer.bas` into your workbook (Alt+F11 → File → Import File). No other dependencies.

**Adoption guidance — using this on your own file:**
1. Open your own messy exported file in Excel.
2. Import `modUTL_DataSanitizer.bas` (one-time setup per workbook, or install permanently in Personal.xlsb so it works across all files).
3. Run `PreviewSanitizeChanges` first (Alt+F8 → select → run). Read the preview report — it tells you exactly what will change.
4. If the preview looks right, run `RunFullSanitize` to apply all fixes.
5. If the preview shows something you don't want changed, use the standalone tools (`FixFloatingPointTails` or `ConvertTextStoredNumbers`) to apply only what you need.
6. The smart keyword detection automatically skips columns that look like IDs, dates, names, or codes — you don't need to set anything up manually.

---

### Workflow 2 — Compare two versions of a file

**What it does:** Picks two sheets (in the same workbook) and compares them cell by cell. Highlights every difference in red and matching cells in green. Builds a structured report showing what changed, where, and the before/after values.

**Why it matters:** Month-end close, budget revisions, file handoffs — anytime two versions of a file should match and you need to prove they do (or find where they don't). Visual color-coding makes the differences impossible to miss.

**Primary module:** `RecTrial\UniversalToolkit\vba\modUTL_Compare.bas`

**Key tools:**
1. `CompareSheets` — pick two sheets from a dropdown, get a color-coded diff report on a new sheet named `UTL_CompareReport`.
2. `CompareRanges` — compare two specific ranges instead of whole sheets (useful when sheets have different structures).
3. `ClearCompareHighlights` — remove all comparison highlighting when you're done.

**Python alternative (optional):** `RecTrial\UniversalToolkit\python\ZeroInstall\compare_workbooks.py` — compares two separate workbook files (not just sheets within one file). Produces a plain-text difference report. Stdlib only, no pip required.

**Demo sample file:** Any two versions of a monthly report (copy a sheet, change a few numbers, then run).

**Time-to-value:** Under 3 minutes for a full visual comparison.

**Prerequisites:** Import `modUTL_Compare.bas` into your workbook.

**Adoption guidance — using this on your own file:**
1. Copy your two file versions into one workbook as separate sheets (or open a workbook that already has two versions).
2. Import `modUTL_Compare.bas`.
3. Run `CompareSheets` → pick the two sheets when prompted → review the `UTL_CompareReport` sheet.
4. Look at red-highlighted cells in each sheet and the report for a clean list of all differences.
5. For comparing two completely separate files: use `compare_workbooks.py` instead — it works without importing anything into Excel.

---

### Workflow 3 — Consolidate multiple sheets or monthly files

**What it does:** Combines data from multiple sheets (or tabs that match a name pattern) into one master sheet. Adds a "Source" column so each row is clearly labeled with where it came from.

**Why it matters:** Rolling up monthly tabs, combining department data from multiple sheets, merging Q1–Q4 into an annual view — consolidation is a constant Finance task that normally requires manual copy-paste or fragile formulas. This automates it in one click.

**Primary module:** `RecTrial\UniversalToolkit\vba\modUTL_Consolidate.bas`

**Key tools:**
1. `ConsolidateSheets` — pick sheets from a numbered list. The tool combines them into a new sheet named `UTL_Consolidated` with a Source column added.
2. `ConsolidateByPattern` — combine all sheets whose names match a pattern (e.g., all sheets starting with "Jan", "Feb", etc.) without picking them one by one.

**Demo sample file:** Any workbook with multiple same-structure monthly tabs.

**Time-to-value:** Under 5 minutes for a multi-sheet consolidation.

**Prerequisites:** Import `modUTL_Consolidate.bas`. Sheets should have the same column structure (same headers, same column order) — the tool works best on identically-structured tabs.

**Adoption guidance — using this on your own file:**
1. Gather all the data you want to combine into one workbook (copy each tab in, or use sheets already there).
2. Make sure column headers match across all sheets — same column names, same order.
3. Import `modUTL_Consolidate.bas`.
4. Run `ConsolidateSheets` → pick which tabs to include → get your combined sheet.
5. The "Source" column tells you which original tab each row came from — keep it; it's your audit trail.
6. If you consolidate every month, keep the module imported. Re-run `ConsolidateSheets` whenever you add new tabs and it will rebuild the combined sheet.

---

### Workflow 4 — Find workbook issues and external links

**What it does:** Scans the workbook for common problems — formulas that reference other files (external links), circular references, error cells (#N/A, #REF!, #DIV/0!), inconsistent formulas within columns, hidden sheets, broken named ranges. Creates a report sheet listing every issue found.

**Why it matters:** Before sharing a workbook, submitting to leadership, or handing off to another team — you need to know if it has any hidden problems. This catches them fast before they cause embarrassment or bad numbers.

**Primary module:** `RecTrial\UniversalToolkit\vba\modUTL_Audit.bas`

**Key tools (Tier 1 — run these first):**
1. `ExternalLinkFinder` — lists every cell referencing an external file, with the file path and cell address. Creates `UTL_ExternalLinks` report.
2. `CircularReferenceDetector` — finds all circular references in the workbook. Creates a report with each offending cell.
3. `WorkbookErrorScanner` — finds all error cells (#N/A, #REF!, #DIV/0!, etc.) across every sheet. Creates a report with each error, sheet, and cell.

**Additional tools (Tier 2 — for deeper review):**
4. `DataQualityScorecard` — scores the workbook 0–100 for data quality based on blanks and errors in data ranges.
5. `InconsistentFormulasAuditor` — finds columns where most rows have a formula but some rows are hardcoded values (a very common cause of wrong totals).
6. `ExternalLinkSeveranceProtocol` — after reviewing external links, converts them to static values so the workbook no longer depends on other files.

**Companion module:** `RecTrial\UniversalToolkit\vba\modUTL_SheetTools.bas`
- `ListAllSheetsWithLinks` — creates a quick inventory of every sheet in the workbook with a clickable hyperlink to each, plus visibility status (visible / hidden / very hidden). Good starting overview.

**Python alternative (optional):** `RecTrial\UniversalToolkit\python\ZeroInstall\profile_workbook.py` — produces a structural overview report of any .xlsx workbook without opening Excel. Shows sheet names, row/column counts, formula density. Stdlib only, no pip required.

**Demo sample file:** Any workbook with a few external links or mixed formula/value cells.

**Time-to-value:** Under 2 minutes for the Tier 1 scan.

**Prerequisites:** Import `modUTL_Audit.bas` (and optionally `modUTL_SheetTools.bas`).

**Adoption guidance — using this on your own file:**
1. Open the workbook you want to audit.
2. Import `modUTL_Audit.bas`.
3. Run `ExternalLinkFinder` first — this is the most common issue in shared Finance files.
4. Run `WorkbookErrorScanner` next — catches hidden #N/A cells that can corrupt totals.
5. Review the report sheets created. Each one has the sheet name, cell address, and exact problem.
6. If external links are safe to remove, run `ExternalLinkSeveranceProtocol` to lock them as static values.
7. Run `DataQualityScorecard` for a summary score — useful to include when handing off a workbook to leadership.

---

### Workflow 5 — Generate a workbook or executive summary

**What it does:** Scans any workbook automatically and produces a plain-English summary — what sheets exist, how much data, how many formulas vs. hardcoded values, any potential issues found, key statistics. Output is a formatted summary sheet ready to paste into an email or print for a meeting.

**Why it matters:** When someone hands you a workbook and asks "what's in here?" or when you need to brief a manager without walking through the whole file — this does the scan and writes the summary for you.

**Primary module:** `RecTrial\UniversalToolkit\vba\modUTL_ExecBrief.bas`

**Key tool:**
1. `GenerateExecBrief` — one click. Scans the entire workbook (sheet inventory, data volume, formula vs. value analysis, hidden sheets, potential data issues) and builds an "Executive Brief" sheet in plain English. No setup required.

**Intelligence module (for data-rich sheets):** `RecTrial\UniversalToolkit\vba\modUTL_Intelligence.bas`

Runs on the active sheet rather than the whole workbook. Three tools:
1. `MaterialityClassifierActiveSheet` — tags each data row as "Material increase," "Material decrease," "Watch," or "Normal" based on $ and % thresholds. Writes two new columns (Materiality Status, Variance %) automatically finding Current and Prior columns by header name.
2. `GenerateExceptionNarrativesActiveSheet` — writes a plain-English sentence for each row explaining the variance ("Revenue increased $45,000 (+12%) vs. prior period — Material increase").
3. `DataQualityScorecardActiveSheet` — scores the active sheet 0–100 for blank cells and error cells.

**Python alternative (optional):** `RecTrial\UniversalToolkit\python\ZeroInstall\build_exec_summary.py` — generates a text-based executive summary from any .xlsx file without needing Excel open. Stdlib only, no pip required.

**Demo sample file:** `RecTrial\SampleFile\SampleFileV2\` or `Sample_Quarterly_ReportV2.xlsm`.

**Time-to-value:** Under 1 minute for the executive brief.

**Prerequisites:** Import `modUTL_ExecBrief.bas` (and optionally `modUTL_Intelligence.bas` + its dependency `modUTL_Core.bas`).

**Adoption guidance — using this on your own file:**
1. Open any workbook — no setup needed for `GenerateExecBrief`.
2. Import `modUTL_ExecBrief.bas`.
3. Run `GenerateExecBrief`. A formatted "Executive Brief" sheet appears.
4. Copy the text to an email or print it — it's written to be readable by someone who hasn't seen the workbook.
5. For a sheet with Current/Prior/Budget columns: import `modUTL_Intelligence.bas` + `modUTL_Core.bas`, navigate to that sheet, and run `MaterialityClassifierActiveSheet` to tag every row automatically. Then run `GenerateExceptionNarrativesActiveSheet` for the plain-English sentence per row.
6. The Intelligence tools work on the active (selected) sheet — make sure the right sheet is active before running.

---

## Python workflows — runs from Command Prompt, no Excel installation required

These two workflows use new Python scripts that are part of the V4 build. They are **not yet built** — this doc describes what they will do and how they will work once they're ready. They will live at `RecTrial\UniversalToolkit\python\ZeroInstall\`. No pip install required — Python standard library only.

---

### Workflow 6 — Revenue Leakage Finder

**What it does:** Compares what customers were expected to be billed (from a contracts file) against what they were actually invoiced (from a billing export). Flags potential billing problems: underbilling, overbilling, missing invoices, duplicate invoices, inactive customers still being billed, product mismatches, stale contracts. Produces a ranked exception report — who to review first and why — plus an executive HTML summary showing total expected vs. total billed and the net variance.

**Why it matters:** Billing errors in SaaS and subscription businesses can quietly accumulate. A customer on a $12,000/year contract being invoiced $10,500 for 6 months isn't obviously broken — but it's $1,500 in potential revenue leakage that the tool will flag. This is the kind of analysis that normally takes a Finance analyst hours to do manually in Excel. Python does it in seconds.

**Primary script:** `revenue_leakage_finder.py` *(to be built at `RecTrial\UniversalToolkit\python\ZeroInstall\`)*

**Inputs:**
- `contracts.csv` — one row per customer/product contract with expected MRR, billing frequency, start/end dates, status
- `billing.csv` — one row per invoice with customer ID, product, billing period, amount billed, and invoice status
- `customer_map.csv` (optional) — maps customer IDs to customer names if the two files use different identifiers

**Outputs (all go to a timestamped folder under `/outputs/`):**
- `revenue_leakage_summary.html` — executive summary with totals, net variance, exception breakdown by type
- `revenue_leakage_exceptions.csv` — full ranked list of every exception found
- `top_10_action_list.csv` — the top 10 exceptions Connor or a coworker should review first
- `run_log.json` + `run_summary.txt` — what the script did, row counts, any warnings

**Sample data provided:** Yes — sample contracts and billing files ship with the tool. Run `--sample` to use them without any setup.

**Time-to-value:**
- With sample data: under 5 minutes from Command Prompt open to HTML report
- With your own real data: allow 15–30 minutes the first time to prepare clean input CSV files

**Prerequisites:** Python installed (version 3.8 or later). No pip install required — all standard library. See `PYTHON_SAFETY.md` for full safety rules.

**Adoption guidance — using this on your own data:**
1. Start with sample mode first: `python revenue_leakage_finder.py --sample` — understand the output before touching real data.
2. Review the HTML report and the exception CSV. Learn what each exception type means.
3. Prepare your own `contracts.csv` from your contract management system or billing platform — it needs: customer ID, product name, expected MRR, billing frequency, contract start/end dates, active/inactive status.
4. Prepare your own `billing.csv` from your invoicing system — it needs: invoice ID, customer ID, product, billing period (month/year), amount billed, invoice status.
5. Run: `python revenue_leakage_finder.py --contracts your_contracts.csv --billing your_billing.csv`
6. Review the exceptions. Prioritize the top_10_action_list. Bring the HTML summary to your manager or the Revenue team.
7. **Do not run on production-sensitive data** until you have run it on sample data and understood every output field.

---

### Workflow 7 — Data Contract Checker

**What it does:** Validates a CSV or data export file against a defined schema before you run analysis on it. Checks for missing required columns, wrong data types, blank required fields, and broken business rules (like negative amounts in a billing file). Reports PASS or FAIL with a clear list of exactly what's wrong and where.

**Why it matters:** Bad inputs produce bad analysis. If a billing export is missing the "amount_billed" column because someone renamed it "Billed Amount" this month, every downstream calculation is wrong. Running the Data Contract Checker first catches this before it corrupts your report — and gives you a clear error message instead of a broken spreadsheet.

**Primary script:** `data_contract_checker.py` *(to be built at `RecTrial\UniversalToolkit\python\ZeroInstall\`)*

**Inputs:**
- Any CSV file to validate
- A schema JSON file defining what columns are required, what types they should be, and any rules (e.g., "amount_billed must not be negative")

**Outputs (all go to a timestamped folder under `/outputs/`):**
- `data_contract_report.html` — clear PASS or FAIL with details for each check
- `data_contract_results.csv` — row-by-row results for sharing or escalation
- `data_contract_summary.json` — machine-readable results (useful for later automation)
- `run_log.json` + `run_summary.txt`

**Demo pattern:** Sample bad-file (red FAIL) → fix one column → re-run → green PASS. This is the video demo sequence in Chapter 4.

**Sample data provided:** Yes — a good-file (PASS) and a bad-file (FAIL) ship with the tool. Run `--sample` to see both.

**Time-to-value:** Under 2 minutes from Command Prompt open to PASS/FAIL result.

**Prerequisites:** Python installed (3.8 or later). No pip required.

**Adoption guidance — using this on your own data:**
1. Start with the sample bad-file demo: `python data_contract_checker.py --sample` — see what a failure report looks like.
2. Create a schema JSON file for your own data file. The schema is a plain text file you can write in Notepad — it lists which columns are required and what type they should be (string, number, date).
3. Connor can provide a starter schema template for the most common Finance file types (billing exports, GL extracts, AR aging files).
4. Run: `python data_contract_checker.py --file your_data.csv --schema your_schema.json`
5. If it fails, read the report — it tells you exactly which columns are missing or wrong.
6. Fix the input file and re-run. When it passes, proceed with your analysis knowing the data structure is clean.

---

## Everything else — advanced / discoverable, not the first recommended path

The full toolkit contains ~140 VBA tools and 28 Python scripts covering dozens of additional use cases: pivot table refresh, tab color coding, bulk column operations, scenario modeling, workbook dependency scanning, control evidence packs, exception triage ranking, variance decomposition, fuzzy matching, PDF extraction, and more.

All of it is inside `Sample_Quarterly_ReportV2.xlsm` (accessible from the Command Center) and in `RecTrial\UniversalToolkit\`. It is NOT hidden — power users who want to explore further will find everything there.

For V1 adoption, these are **not the first-recommended path**. They are available for coworkers who have completed at least one of the 7 starter workflows, are comfortable with the basics, and want to go deeper.

If a coworker asks about a specific advanced tool, Connor can point them to the right module. But the default answer to "where do I start?" is always one of the 7 workflows above.

---

## Quick-reference grid

| # | Workflow | Primary Module / Script | Python option | Effort to start |
|---|---|---|---|---|
| 1 | Clean a messy export | `modUTL_DataSanitizer.bas` | `sanitize_dataset.py` | < 5 min |
| 2 | Compare two files | `modUTL_Compare.bas` | `compare_workbooks.py` | < 5 min |
| 3 | Consolidate sheets/files | `modUTL_Consolidate.bas` | — | < 5 min |
| 4 | Find workbook issues | `modUTL_Audit.bas` | `profile_workbook.py` | < 5 min |
| 5 | Generate a workbook summary | `modUTL_ExecBrief.bas` | `build_exec_summary.py` | < 5 min |
| 6 | Revenue Leakage Finder | `revenue_leakage_finder.py` *(to build)* | — | < 30 min with real data |
| 7 | Data Contract Checker | `data_contract_checker.py` *(to build)* | — | < 5 min |

**Module file location:** `RecTrial\UniversalToolkit\vba\` (VBA) and `RecTrial\UniversalToolkit\python\ZeroInstall\` (Python)

**Sample file for VBA demo:** `RecTrial\SampleFile\SampleFileV2\`

**Sample files for Python V4 demo:** Will be at `RecTrial\UniversalToolkit\python\ZeroInstall\samples\` once built.

---

**End of supported workflows doc.**
