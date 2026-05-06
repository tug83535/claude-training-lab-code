# V4 Finish Line + Phase 2 Tool Specs

**Context:** Videos 1-3 shipped. V4 plan locked. All 6 V4 Python scripts built and smoke-tested 5/5 PASS. This doc covers two things: (1) the punch list to ship V4, and (2) the 7 tools to build after V4 ships.

**Rule:** Nothing in Part 2 gets built until V4 is recorded, shipped, and in the hands of 10-20 pilot users.

---

# PART 1: SHIP V4

## BLOCKER — Python Dependency Problem

V4 scripts import pandas and openpyxl. Bundled Python 3.11 embeddable does NOT include these.

**Fix (Option A — vendor the packages into the zip):**
1. On Connor's machine, locate the installed pandas, openpyxl, numpy, et_xmlfile, and dateutil packages (typically in `Lib\site-packages\`).
2. Copy those package folders into a `lib/` folder inside the zip alongside the bundled Python.
3. Add or modify `python311._pth` to include `.\lib` so bundled Python finds the packages.
4. Test by running `finance_automation_launcher.py` using ONLY the bundled Python — no system Python, no PATH.
5. If the import works, the dependency problem is solved.

**Claude Code action:** Help Connor implement this. Create a script that copies the required packages, sets up the path file, and validates imports from the bundled Python.

## 10-Step Finish Sequence

| # | Task | Est. Time | Notes |
|---|---|---|---|
| 1 | Resolve Python dependency (above) | 2-3 hrs | BLOCKER |
| 2 | Git commit all V4 Python files | 30 min | Flagged as NOT YET COMMITTED |
| 3 | Write V4 narration script — 8 chapters, 9-12 min | 3-4 hrs | Source: `RecTrial\Brainstorm\VIDEO_4_REVISED_PLAN.md`. Hero: Revenue Leakage Finder. Closing: ARR waterfall. |
| 4 | Build the one Excel launcher button in FinanceTools.xlsm | 1-2 hrs | VBA Shell() → bundled Python → finance_automation_launcher.py |
| 5 | Generate ElevenLabs audio from narration script | 1-2 hrs | Connor does this |
| 6 | Full dry-run of V4 recording | 1 hr | Catch problems before real take |
| 7 | Record Video 4 | 2-3 hrs | |
| 8 | Assemble SharePoint zip | 2-3 hrs | See zip structure below |
| 9 | Test zip on one coworker's machine | 1-2 hrs | No help from Connor — they open it alone |
| 10 | Ship to 10-20 pilot users | 1 hr | Role mix per MINIMUM_DISTRIBUTION_PLAN.md |

## SharePoint Zip Structure

```
FinanceTools_v1.0/
├── START_HERE.md              — "Open FinanceTools.xlsm and click the big blue button"
├── FinanceTools.xlsm          — The workbook with the launcher button
├── python/                    — Bundled Python 3.11 + vendored packages + all scripts
├── sample_data/               — Fake files coworkers can safely test with
└── docs/
    ├── Quick_Start.pdf
    └── Troubleshooting.pdf
```

## Fake Leakage Data Check

Before recording, verify the sample data makes the demo compelling:

- [ ] Total potential leakage is in the $300K-$500K range
- [ ] At least 3 leakage categories represented (unbilled, rate mismatch, coverage gap)
- [ ] At least 10-15 flagged contracts
- [ ] Summary screen shows the total dollar amount prominently
- [ ] Every flagged item says "REVIEW QUEUE — potential exception, not confirmed error"

## Quick Win Before Recording

- [ ] Add "REVIEW QUEUE" warning text to every exception output across all V4 scripts. 30-minute consistency pass. Makes the tools enterprise-safe.

---

# PART 2: POST-V4 TOOLS (Build AFTER V4 ships)

7 new tools ranked by coworker daily impact. These fill genuine workflow gaps not covered by the existing 140+ tools.

---

## Tool 1: Cross-File Reconciler (VBA)

**The gap:** Coworkers reconcile two files every month — billing vs. revenue ledger, GL vs. sub-ledger, bank statement vs. cash account. The toolkit has sheet-to-sheet compare (modUTL_Compare) and Python file compare (compare_files.py), but no VBA tool that opens two separate workbooks and reconciles them. Most coworkers won't run Python. This is the single highest-impact missing tool.

**What it does:**
1. User picks two open workbooks.
2. Selects the key column in each (invoice number, customer ID, GL account — whatever they match on).
3. Selects the amount column in each.
4. Tool compares every row by key, matches records, and finds differences.

**Output — new workbook with 5 tabs:**
- **Summary** — total records in each file, matched count, unmatched count, match rate %, total variance
- **Matched** — all records that matched on key AND amount
- **Amount Mismatches** — matched on key but amounts differ (shows both amounts + variance)
- **File A Only** — records in File A with no match in File B
- **File B Only** — records in File B with no match in File A

**Technical details:**
- Module name: `modUTL_CrossFileReconciler.bas`
- Folder: `UniversalToolkit/vba/`
- Technology: VBA — Dictionary-based key matching for speed
- Tolerance: configurable (e.g., match within $0.01 for rounding differences)
- Safety: read-only on both source files, output is a brand-new workbook, row counts logged

**Build estimate:** 3-4 hours

**Why it matters:** Reconciliation is the #1 time consumer in month-end close. Every Finance person in the building does this. A VBA version means everyone can use it — no Python required.

---

## Tool 2: Schedule Builder (Python)

**The gap:** Finance teams build straight-line amortization and revenue recognition schedules by hand every month. Prepaid expenses, deferred revenue, subscription allocations, lease payments — all require the same math: spread a total amount evenly across a date range, with partial-month proration. Manual errors happen constantly at the partial-month boundaries.

**What it does:**
1. User provides: start date, end date, total amount, and optionally a name/description.
2. Tool calculates the monthly amount (handling partial first and last months correctly).
3. Outputs a clean schedule.

**Output — Excel workbook with 2 tabs:**
- **Schedule** — one row per month: period, days in period, period amount, cumulative recognized, remaining balance
- **Summary** — total amount, term months, monthly amount, start/end dates, proration method

**Supports batch mode:** user provides a CSV with multiple items (10 prepaid invoices, 50 contracts, etc.) and gets a combined schedule plus individual item detail.

**Technical details:**
- Script name: `schedule_builder.py`
- Folder: `UniversalToolkit/python/ZeroInstall/` (if built stdlib-only) or `UniversalToolkit/python/` (if using pandas)
- Inputs: CSV or manual entry (start_date, end_date, total_amount, description)
- Proration: actual days in month / total days in period (not simplified 30-day months)
- Safety: output only, input file unchanged, totals must tie (sum of period amounts = total amount exactly)

**Build estimate:** 2-3 hours

**Why it matters:** Revenue accounting and FP&A teams build these schedules monthly. One tool replaces dozens of manual spreadsheets and eliminates the partial-month proration errors that cause audit findings.

---

## Tool 3: Pricing Exception Finder (Python)

**The gap:** The Revenue Leakage Finder catches contracts that weren't billed. This tool catches contracts that were billed at the WRONG AMOUNT. Together they form a complete billing accuracy control.

**What it does:**
1. For each invoice line, looks up the expected price from the price book.
2. Checks for contract-level pricing overrides (negotiated rates).
3. Checks for approved discount schedules.
4. Calculates: expected_price = override_price OR (standard_price × (1 - discount%))
5. Compares expected_price vs actual invoice line unit_price.
6. Flags lines where the variance exceeds a configurable tolerance.

**Output — Excel workbook with 6 tabs:**
- **Summary** — total over-billed amount, total under-billed amount, exception count by category
- **Under-Priced** — lines where actual < expected (revenue at risk — company losing money)
- **Over-Priced** — lines where actual > expected (customer charged too much — relationship risk)
- **Within Tolerance** — flagged but within acceptable range (e.g., rounding differences)
- **Clean Lines** — lines that passed all checks
- **Pricing Logic** — documents the expected-price calculation for each line (auditability)

**Technical details:**
- Script name: `pricing_exception_finder.py`
- Folder: `UniversalToolkit/python/ZeroInstall/` — reuse patterns from `revenue_leakage_finder.py`
- Inputs: invoice_lines.csv, price_book.csv, contracts.csv, (optional) approved_discounts.csv
- Tolerance: configurable (default: 5% or $500, whichever is greater)
- Safety: read-only on all inputs, output labeled as "REVIEW QUEUE"

**Build estimate:** 3-4 hours

**Why it matters:** Pricing errors are high-dollar. Under-billing is direct revenue leakage. Over-billing creates customer disputes and credit memos. Neither gets caught systematically without a tool like this.

---

## Tool 4: Print Ready (VBA)

**The gap:** Every time someone sends an Excel file to leadership, they spend 10+ minutes per sheet setting margins, headers, footers, orientation, print area, scaling, and repeat rows. Everyone does this. Nobody has automated it.

**What it does:**
1. One-click applies professional print settings to every sheet in the active workbook.
2. Settings: landscape orientation, narrow margins, header (sheet name left, company center, date right), footer (page X of Y), print area set to used range, scale to fit width on one page, repeat header row on every printed page, auto-fit columns.
3. Optional: user picks specific sheets instead of all.

**Output:** The same workbook with print settings applied. No new file created (this modifies print settings only, not data — safe to run on the working file).

**Technical details:**
- Module name: `modUTL_PrintReady.bas`
- Folder: `UniversalToolkit/vba/`
- One main sub: `MakePrintReady` — loops through selected or all sheets
- Safety: only changes PageSetup properties (margins, headers, footers, orientation, scaling, print area, repeat rows). Does not touch cell data, formulas, formatting, or structure.

**Build estimate:** 2 hours

**Why it matters:** Small time savings multiplied by every person, every report, every month. This is the type of tool that gets used 3x per week and makes people say "why didn't we have this before?"

---

## Tool 5: Waterfall Chart Builder (VBA)

**The gap:** The demo file builds waterfall charts via modDashboard, but that code is hardcoded to the P&L sheet structure. There's no universal tool where a coworker can select any data range and get a properly formatted waterfall chart. Waterfall charts are the most requested and most frustrating chart type in Excel.

**What it does:**
1. User selects a range with labels in column A and values in column B (or the tool detects the layout).
2. First row is treated as the starting value, last row as the ending total, everything in between as increases (positive) or decreases (negative).
3. Tool builds a stacked bar chart with invisible base segments to create the waterfall effect.
4. Color coding: green for increases, red for decreases, blue/gray for totals.
5. Value labels on each bar. Clean axis formatting.

**Output:** A new chart sheet or chart object on the active sheet with a properly formatted waterfall.

**Technical details:**
- Module name: `modUTL_WaterfallChart.bas`
- Folder: `UniversalToolkit/vba/`
- Chart type: stacked bar with invisible "base" series (the standard Excel workaround for true waterfalls)
- Colors: iPipeline brand green for positive, brand red-accent for negative, brand blue for totals
- Safety: creates a new chart, does not modify the source data range

**Build estimate:** 3-4 hours

**Why it matters:** Finance teams use waterfalls for ARR bridges, P&L walks, variance explanations, and cash flow analysis. Building one manually in Excel takes 20-30 minutes of fiddly formatting. One click vs. 30 minutes, every time.

---

## Tool 6: Journal Entry Builder (Python)

**The gap:** Close teams prepare journal entry batches from templates every month. The process is manual: type accounts, debits, credits, descriptions into a spreadsheet, check that debits = credits, format for upload. Errors are common and cause posting failures.

**What it does:**
1. User fills in a simple input template (CSV or Excel): account, debit_amount, credit_amount, description, period, entity.
2. Tool validates each entry: debits = credits per JE number, all accounts exist in the chart of accounts, required fields populated, amounts are positive numbers.
3. Produces a formatted JE batch ready for review and upload.

**Output — Excel workbook with 3 tabs:**
- **JE Batch** — formatted journal entries with JE number, line number, account, debit, credit, description, period, entity
- **Validation Results** — pass/fail per JE: balanced check, valid accounts, required fields, amount checks
- **Failed Entries** — any JE that failed validation, with the specific reason

**Technical details:**
- Script name: `journal_entry_builder.py`
- Folder: `UniversalToolkit/python/`
- Inputs: JE input template (CSV), chart of accounts (CSV)
- Validation: debits = credits per JE (exact match, not tolerance), account exists in COA, period format valid, no blank required fields, no negative amounts
- Safety: output only, input template unchanged, validation runs before output is produced

**Build estimate:** 3 hours

**Why it matters:** JE preparation errors cause posting failures, reprocessing, and audit findings. A tool that validates before submission catches problems at the cheapest possible point — before they enter the system.

---

## Tool 7: Version Snapshot (VBA)

**The gap:** The demo file has modVersionControl, but it depends on modConfig and demo-specific sheet structures. There's no universal "save what this sheet looks like right now" tool. During month-end close, coworkers often want a rollback point before making changes but don't want to save an entire new file.

**What it does:**
1. User clicks the button while on any sheet.
2. Tool copies the current sheet's data (values only, no formulas) to a new tab named `Snapshot_[SheetName]_[YYYYMMDD_HHMM]`.
3. The snapshot tab is protected (read-only) so it doesn't get accidentally edited.
4. A small note at the top of the snapshot says when it was taken and which sheet it came from.

**Output:** A new read-only tab in the same workbook containing a frozen copy of the original sheet's data at that moment.

**Technical details:**
- Module name: `modUTL_VersionSnapshot.bas`
- Folder: `UniversalToolkit/vba/`
- Copy method: paste values only (no formulas — snapshot is a frozen picture, not a live copy)
- Tab name: `Snap_[SheetName]_[MMDD_HHMM]` (truncated to stay within Excel's 31-character tab name limit)
- Safety: adds a new sheet, never modifies or deletes existing sheets, never overwrites a previous snapshot

**Build estimate:** 2 hours

**Why it matters:** "I want to save what this looks like before I change it" is one of the most common things people do during close. Right now they either Save As (clutters the folder) or just hope they remember what the numbers were. This gives them a one-click safety net inside the same workbook.

---

# PART 3: BUILD PRIORITY AFTER V4

| Priority | Tool | Type | Build Time | When |
|---|---|---|---|---|
| 1 | Cross-File Reconciler | VBA | 3-4 hrs | First — highest daily impact |
| 2 | Schedule Builder | Python | 2-3 hrs | Second — used monthly by multiple teams |
| 3 | Pricing Exception Finder | Python | 3-4 hrs | Third — completes billing controls |
| 4 | Print Ready | VBA | 2 hrs | Anytime — quick win |
| 5 | Waterfall Chart Builder | VBA | 3-4 hrs | After pilot feedback |
| 6 | Journal Entry Builder | Python | 3 hrs | Phase 2 |
| 7 | Version Snapshot | VBA | 2 hrs | Phase 2 |

**Total Phase 2 build time:** ~18-22 hours across all 7 tools.

**Rule:** Ship V4 first. Get pilot feedback. Build these based on what coworkers actually ask for — the ranking above is a starting point, not a mandate. If 5 pilot users ask for the Waterfall Chart Builder before anyone asks for the Schedule Builder, build the waterfall first.

---

*Archive this doc after Phase 2 is complete.*
