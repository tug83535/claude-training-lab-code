# Hard Constraints — Read Before Designing Features

The purpose of this project is to **show coworkers what's possible BEYOND what Excel and OneDrive already do natively.** If you build a feature that replicates something a Finance user can already accomplish with a few native clicks, you've wasted a feature slot.

Before adding any feature, ask: **"Could a coworker do this in <5 clicks using only built-in Excel or OneDrive features?"** If yes — don't build it. Replace it with something more ambitious.

---

## Banned Features (Do NOT Rebuild These)

### Native Excel features (built-in, zero-code)
- Sort and Filter (Data tab → Sort / Filter)
- AutoFilter dropdowns on headers
- Conditional Formatting with the GUI (color scales, data bars, icon sets, top/bottom rules)
- Basic PivotTables (Insert → PivotTable)
- Basic PivotCharts (Insert → PivotChart)
- Slicers on Tables/PivotTables
- Data Validation via GUI (Data → Data Validation → dropdown list)
- Remove Duplicates (Data → Remove Duplicates)
- Text-to-Columns (Data → Text to Columns)
- Flash Fill (Ctrl+E)
- Find & Replace (Ctrl+H)
- Freeze Panes
- Print to PDF (File → Export → PDF)
- Named Ranges via Name Manager
- Tables (Ctrl+T) with their native formatting and structured references
- XLOOKUP / INDEX-MATCH / VLOOKUP (these are built-in formulas — don't write a VBA wrapper around them)
- IFERROR / IFS / SWITCH / LET / LAMBDA (all native Excel functions)
- Power Query (Data → Get Data → From …) for basic ETL
- Power Pivot / Data Model for multi-table models
- Charts (Insert → Charts) with all their built-in types
- Sparklines

### OneDrive / Microsoft 365 features
- AutoSave
- Version History / Restore Previous Version
- Co-authoring / Real-time collaboration
- File sharing with edit/view permissions
- Excel for the Web
- Share link generation
- Recycle Bin / file recovery
- Mobile sync

### Microsoft 365 CoPilot features
- Ad-hoc data summarization ("summarize this sheet")
- Natural-language formula creation
- AI-generated charts
- AI commentary on data

If CoPilot already does it natively, don't build it. (Do reference CoPilot heavily in the CoPilot Prompt Guide — it's the user's ally, not your competitor.)

---

## What "Going Beyond" Looks Like

Your features should cross **at least one** of these thresholds to justify existing:

### 1. Multi-step / multi-sheet automation
A single click triggers a sequence that would take a human 10+ clicks and touch multiple sheets. Example: month-end close that reads 12 monthly tabs, runs reconciliation, generates variance commentary, builds a PDF pack, and emails it.

### 2. Cross-file / cross-source workflows
Pulls data from outside the workbook — a database, an API, a folder of CSVs, another Excel file — and brings it in with transformation logic.

### 3. Intelligence / decision logic
Makes a judgment call a user would otherwise have to make manually. Example: "Flag any variance > 15% AND > $10K AND label it material." Native conditional formatting can highlight a cell; it can't *decide* based on a compound rule and write a commentary sentence explaining it.

### 4. Output generation
Produces a polished artifact — a branded PDF, an executive brief document, a forecast model, a what-if comparison — rather than just formatting existing data.

### 5. Performance / scale
Does something that would choke native Excel. Example: sanitize 50,000 rows across 30 sheets in 3 seconds. Compare two workbooks cell-by-cell and highlight diffs. Consolidate 100 tabs.

### 6. Plug-and-play portability (Prong 1)
Works on *any* workbook regardless of sheet names, column counts, or header positions. This is the whole point of the universal toolkit — no native Excel feature transfers across workbooks the way a smart VBA module does.

### 7. Teaching leverage
The code itself is demonstrably more powerful than a coworker could reach natively — and paired with the CoPilot Prompt Guide, they can adapt it to their own situation without learning to code.

---

## Quality Non-Negotiables

### Every feature must:
- Have a **name** a non-developer understands
- Be **accessible in one click** (Command Center button, ribbon button, or keyboard shortcut)
- Handle **realistic Finance-file weirdness** (merged cells, blank rows, text-stored numbers, formulas that error, hidden sheets)
- Log what it did (so a user can audit or undo)
- Display a **clear success or failure message** at the end

### Every file must:
- Open without "This file contains macros" confusion for the user
- Not be bloated — no 100MB workbooks, no 50,000 unnecessary named ranges
- Have **every sheet named and purposeful** — no "Sheet1," "Sheet2"

### Every guide must:
- Be written for non-developers
- Include screenshots or step-by-step annotations where screenshots aren't possible in markdown
- Have a "Troubleshooting" section for common failures
- State prerequisites at the top ("You need Excel 2019+ and macros enabled")

### Every video script must:
- Have a clear hook in the first 15 seconds
- Include narration you'd actually say out loud (no "As you can see" filler — assume the viewer is watching and can see)
- Include on-screen callouts and timing cues
- End with a clear call to action

---

## Pragmatic Exceptions

You *may* re-implement a native feature if:
- You're wrapping it inside a larger multi-step workflow (e.g., "Month-End Close" that happens to include a sort step)
- You're making it **plug-and-play across any file** (Prong 1 — the universal toolkit — which native Excel features cannot do, because they live in the GUI, not in portable code)
- You're adding intelligence on top (e.g., a "smart sort" that detects data types and sorts accordingly — not just a wrapper for `.Sort`)

When in doubt, ask the user.

---

## Forbidden Techniques

- **`SendKeys` against modal dialogs.** It is unreliable. If a feature needs a dialog parameter, accept it as a sub parameter instead of popping an `InputBox`. Silent parameterized subs are far more robust than dialog-driven ones.
- **Hardcoded absolute file paths.** Use `ThisWorkbook.Path`, `Environ("USERPROFILE")`, or ask the user to pick a file via `Application.GetOpenFilename`.
- **Hardcoded sheet indices without fallbacks.** Use sheet names, and guard against missing sheets.
- **Silent failures.** Never `On Error Resume Next` without a corresponding `On Error GoTo 0` and a reason.
- **Bypassing safety checks.** If a pre-commit hook fails, fix the issue. Don't `--no-verify`.
- **Shipping untested code.** Before declaring a feature done, run it against the actual sample file. Capture the output.

---

## Summary

**The standard isn't "it works."** The standard is "it does something a Finance user couldn't already do in Excel, it does it reliably, it's branded, and it makes coworkers say 'I need to use this.'"

If you're about to ship a feature and you can't answer "what does this do that Excel doesn't already do?" — stop and redesign.
