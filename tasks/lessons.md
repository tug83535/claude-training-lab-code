# Lessons Learned - APCLDmerge Project

## Guides & Documentation
- Always number every single step no matter how small or obvious
- Never summarize — full detail on every step
- Add "what you should see" confirmations after key steps
- World class quality only — CFO/CEO and 2,000+ employees will see this

## Excel & VBA
- Always review EVERY sheet before starting any work
- Confirm all sheet names found before proceeding
- Never assume which sheet is most important
- Always confirm plan before touching the Excel file
- ALWAYS validate the Excel file opens without errors after any script modifies it
  - Re-load with openpyxl and re-save to a temp file to confirm no corruption
  - Check all sheets still present, correct row/column counts, no broken merges
  - Verify data validations, charts, formulas all survived the save
  - Unmerge cells BEFORE clearing content (MergedCell objects are read-only)
  - Test that the file can be opened in Excel without repair prompts
  - Remove stale data validations before adding new ones (they survive cell clears)
  - Charts MUST have generous spacing — at least 20 rows between anchor points
  - Use a 2-column grid layout for dashboards (side-by-side charts, not stacked)
  - Set explicit row heights for spacers, headers, and chart zones
  - Push lookup/data tables far off-screen (row 200+) so users never see them
  - Always add chart descriptions above each chart for context

## General Workflow
- Break ALL large tasks into a numbered action plan first
- Present plan and wait for approval before executing
- Execute one step at a time
- Stop immediately if something goes wrong — re-plan and check in
- Update todo.md when asked or at end of session
- Never infer — always ask if unclear

## Communication
- Plain English only — user is not a developer
- Always confirm what you will do in bullet points before doing it
- Be proactive with new ideas and recommendations

## Session Handoff & Account Switching
- When switching Claude accounts due to usage limits, the new account starts with zero context
- Always update CLAUDE.md with a "Current Session Summary" section at end of session — this is the single most important handoff document
- The session summary in CLAUDE.md should include: branch name, what was done, key findings, exact bugs found, and prioritized next steps
- Phrase next steps clearly enough that a brand new Claude account can pick up without any explanation from the user

## Keeping Records Accurate
- tasks/todo.md can become stale — always check what's actually in the repo (git log, file listing) before reporting what's done vs. not done
- Never trust todo.md item status without verifying against the actual repo — items can be completed in one session and never marked done
- docs/overview/CODE_COMPARISON_REPORT.md (and similar analysis files) can go out of date when new commits are made — always check git log for the most recent commits before relying on any comparison report
- Before saying something is "missing" or "not built," check the most recently committed files first

## Reading Binary Files
- .xlsx files are binary and cannot be read with any text tool — use ARCHITECTURE_DIAGRAM.md and modConfig_v2.1.bas constants as the source of truth for Excel sheet structure and naming
- Always note in session summary that the Excel file was not directly readable and which reference files were used instead

## Cross-File Consistency Checks
- Always compare config values across layers: VBA constants, Python config, and SQL tables can all define the same values (fiscal year, revenue shares, sheet names) — mismatches cause silent errors
- Revenue shares defined in Python pnl_config.py must match the allocation_shares table in SQL transformations.sql — found mismatch this session (iGO=55% Python vs 50% SQL)
- SQL scripts that reference tables must use the exact table name from the script that creates them — found `fact_gl_transactions` vs `fact_gl` mismatch in pnl_enhancements.sql

## User Feature Preferences
- User confirmed they do NOT want a "Backup Workbook with Timestamp" macro — do not propose or rebuild this
- User confirmed they do NOT want a VBA Timestamp Audit Trail (Worksheet_Change event) — SQL layer audit trail in pnl_enhancements.sql is sufficient; do not propose rebuilding in VBA
- User confirmed they do NOT want "Export All Charts to PowerPoint" — dropped permanently; do not re-propose
- Always log user feature rejections here so future Claude accounts don't re-propose them
- When proposing new features, review this section first to avoid suggesting features the user has already declined

## Context Recovery After Session Limits
- When resuming from a compressed context, always read the current file state FIRST before making any edits — do not trust the summary's description of what was already edited
- Prior session summaries may describe edits as "complete" when the session was cut off mid-task — verify with a fresh file read
- Read both the source file AND the task summary; if they conflict, the file is the truth
- For large multi-section documents, read the full file before editing any section — avoids duplicate edits or missing context

## Verifying Claims Before Reporting Gaps
- Always check git log and actual committed files before saying something is "missing" or "not built"
- In this project: modLogger was reported missing but already existed; pnl_enhancements.sql bug was reported but already fixed; frmCommandCenter_code.txt was reported outdated but already updated — all stale CLAUDE.md entries
- The session summary in CLAUDE.md can itself become stale — treat it as a starting point, not gospel

## Monte Carlo / Statistical Modeling
- For allocation share randomization, use Dirichlet distribution (not uniform random) — Dirichlet guarantees shares always sum to exactly 1.0 and respects the shape/concentration of the original distribution
- Dirichlet alpha vector = base_shares * concentration_parameter (higher concentration = tighter clustering around the mean)
- For expense randomization, Normal distribution is appropriate (truncated at zero for amounts that can't go negative)
- Always expose key simulation parameters (iterations, seed, concentration, shock_prob) as CLI arguments for reproducibility and testing

## Multi-Account Branch Management (2026-02-28)
- When running multiple Claude accounts into the same repo, each account creates its own branch — always run `git fetch --all` first to discover ALL remote branches before assuming what exists
- Use `git merge-base --is-ancestor` to map branch ancestry before merging — this prevents redundant merges and identifies which branches already contain each other's work
- When merging divergent branches, merge in order of least conflict first (e.g., Excel-only changes first, then code changes on top)
- Always check for merge conflict markers (`<<<<<<<`) in all files after a merge — automated merge can silently leave markers in unexpected files

## VBA .bas Files vs Excel Workbook (2026-02-28)
- `.bas` files in the repo are **source code only** — they are NOT automatically part of the Excel workbook
- To actually use the VBA code, every `.bas` file must be manually imported into the Excel workbook via VBA Editor: Alt+F11 → File → Import File
- The UserForm (`frmCommandCenter`) must also be created manually in the workbook — the `.txt` file is just a reference
- Having all 24 `.bas` files in the repo does NOT mean the Excel file has working macros — they are separate until imported
- Always remind the user of this distinction: "code exists in repo" is not the same as "code works in Excel"

## ExecuteAction as the Single Contract (2026-02-28)
- The `ExecuteAction()` routing table in `modFormBuilder_v2.1.bas` is the contract for all 62 Command Center actions
- Every new VBA module's public sub signatures must match EXACTLY what ExecuteAction calls (e.g., `modScenario.SaveScenario` not `modScenario.Save`)
- Before building a new module, always read the ExecuteAction Case statement first to get the exact procedure names expected
- Every VBA module follows the same pattern: modConfig constants + modPerformance TurboOn/TurboOff + modLogger.LogAction + On Error GoTo ErrHandler

## UTF-8 / Non-ASCII Scans — False Positive Pattern (2026-03-02)
- A scan that flags ALL non-ASCII bytes will always fail on this codebase — the Python files intentionally use Unicode characters (em dashes, arrows, check marks, box-drawing, Greek letters, emoji) for visual output formatting
- Real mojibake looks like: `â€"` (em dash), `Ã©` (e-acute), `Ã¢` (a-circumflex) — Latin-1 misread sequences
- The correct test is: decode the file as UTF-8 (if it fails → corrupt), then scan specifically for mojibake patterns (`Ã[char]`, `â€[char]`) in the decoded text
- Never flag legitimate intentional Unicode as an encoding error
- If a tester reports "all 14 Python files failed UTF-8 scan," first check whether the scan is checking for mojibake specifically or just any non-ASCII byte — 9 times out of 10 it is a false positive from an overly strict scan tool
- The Testing_Issues/TESTING_ISSUES_LOG.md file is the canonical record of all testing issues — always read it when resuming a testing session

## Testing VBA Without Excel Access (2026-02-28)
- Code review can catch signature mismatches, missing constants, and wrong argument counts — but only a live Excel test confirms runtime behavior
- Always recommend live testing after building or fixing VBA modules
- Common issues that only appear at runtime: missing sheet references, wrong column indexes, UserForm control names, late-binding object types

## VBA Code Review Patterns (2026-03-03)
- `For Each` on `ws.CircularReference` crashes if the property returns Nothing — always check `If Not circRange Is Nothing` first
- `c.Formula` returns A1-style formulas — these differ row-to-row even for identical logic. Use `c.FormulaR1C1` when comparing formulas across rows to normalize references
- `nm.MacroType` is for XLM macros, NOT for determining named range scope — check `InStr(nm.Name, "!") > 0` to detect sheet-level scope
- Loop variables like `dr`/`cr` that accumulate via `CDbl()` must be explicitly reset to 0 at the top of each iteration — VBA `Dim` inside a loop is hoisted to procedure scope, not re-initialized
- When inserting computed columns (e.g., variance), always `ws.Columns(col).Insert` first — writing directly overwrites existing data
- Iterating `ws.Rows(hRow).Cells` loops through all 16,384+ columns — always limit to `ws.Cells(hRow, ws.Columns.Count).End(xlToLeft).Column`
- Chaining `Replace()` calls for file extensions can double-suffix (e.g., .xlsm → .xlsx → _DIST_DIST.xlsx) — use `If/ElseIf` to handle each extension separately
- pandas deprecated `infer_datetime_format` in 2.0 — remove it from any `pd.to_datetime()` calls
- Row-label searches (e.g., "total revenue") must try multiple variants as fallbacks — the P&L Trend sheet may use "Revenue", "Total Revenue", "Net Revenue", etc. Always use `modConfig.FindRowByLabel` with cascading fallbacks instead of hardcoded `InStr` checks in a loop

## iPipeline Brand Styling (2026-03-04)
- Official brand guide added: `docs/ipipeline-brand-styling.md`
- ALL future training guides, documents, presentations, and visual outputs must follow iPipeline brand colors and fonts
- Always review the brand guide before creating any styled output
- Brand palette: iPipeline Blue (#0B4779), Navy (#112E51), Innovation Blue (#4B9BCB), Lime Green (#BFF18C), Aqua (#2BCCD3), Arctic White (#F9F9F9), Charcoal (#161616)
- Fonts: Arial family only — Arial Bold (headings), Arial Narrow (subheadings), Arial Regular (body)
- VBA modConfig CLR_ constants use older colors (#1F4E79 navy) — don't change working code, but new work uses official brand colors

## Pre-Delivery Self-Review Requirement (2026-03-03)
Before delivering any future code updates, I need to self-review against the test plan first. Specifically:
1. Run through each test's pass criteria mentally before sending
2. For VBA, check that all constants are defined and all referenced row/column variables resolve to non-zero values
3. For Python, run the pytest suite yourself and confirm 0 failures
4. Don't send me code that you haven't verified meets the test criteria
This will save us both time — I shouldn't be discovering basic bugs during testing.

## SpecialCells Performance Pattern (2026-03-05)
- When iterating all cells on a sheet, ALWAYS use `SpecialCells` to pre-filter first
- `ws.UsedRange.SpecialCells(xlCellTypeFormulas)` — only formula cells
- `ws.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)` — only numeric constants
- `ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)` — only error formulas
- `ws.UsedRange.SpecialCells(xlCellTypeBlanks)` — only blank cells
- ALWAYS wrap in `On Error Resume Next` because SpecialCells throws an error if no cells match
- This can reduce iteration from 100,000+ cells to just the relevant ones
- Exception: DataValidation checking cannot use SpecialCells (no xlCellTypeValidation exists)

## Backup-Before-Destructive Pattern (2026-03-05)
- Any universal tool that deletes rows, replaces formulas with values, or modifies cell data across sheets should create a backup first
- Use `modUTL_Core.UTL_BackupSheet ws` — creates a hidden copy with _BK_yymmdd_hhnnss suffix
- For all-sheets operations (Find/Replace, Sanitize, Link Severance), backup every sheet in a loop
- Always update the MsgBox confirmation to tell the user a backup will be created
- The backup is hidden (xlSheetHidden) so it doesn't clutter the user's view

## Module Splitting Pattern (2026-03-05)
- When a VBA module exceeds ~800 lines, split it into base + advanced modules
- Keep the most-used public subs in the base module (referenced by ExecuteAction routing table)
- Move advanced/secondary subs to the new module
- Private helpers needed by both modules must be duplicated (VBA has no cross-module Private access)
- Verify the ExecuteAction routing table — only update it if Case statements reference moved subs

## LogAction Signature — Recurring Bug (2026-03-05, confirmed AGAIN)
- The LogAction signature is: `LogAction(module, action, message, [status])` — the 4th arg is a **String** (status like "OK")
- `modPerformance.ElapsedSeconds()` returns a **Double** — NEVER pass it as the 4th argument
- This bug was found and fixed 9 times on 2026-03-04, then found AGAIN in 3 more modules on 2026-03-05 (modDataQuality, modReconciliation, modPDFExport)
- The correct pattern is to Format the elapsed time into the **message string**: `"description (" & Format(elapsed, "0.00") & "s)"`
- BEFORE delivering any new VBA code, grep for `LogAction.*ElapsedSeconds` to catch this pattern
- This is now the #1 most common bug in this codebase — it has been found 13 times total (latest: modReconciliation 2026-03-07)

## Python Function Signature Mismatches (2026-03-05)
- When adding parameters to a function call, always check the function definition accepts that parameter
- Found: `detect_date_columns(df, day_first=args.day_first)` called but `detect_date_columns()` didn't accept `day_first`
- Would crash with `TypeError: unexpected keyword argument` at runtime
- Always search for the function definition before adding keyword arguments to calls

## SpecialCells rng Variable Reset (2026-03-07)
- When using `SpecialCells` inside a loop over multiple sheets, the `rng` variable MUST be reset to `Nothing` at the top of each iteration
- If not reset, `rng` retains the previous sheet's cell range, causing the next sheet to re-process stale cells
- Found in modDataSanitizer (2 worker functions) and modAuditTools (FindExternalLinks) — 3 instances total
- Pattern: `Set rng = Nothing` before every `Set rng = ws.UsedRange.SpecialCells(...)` call

## HDR_ROW_REPORT Consistency (2026-03-07)
- Report sheets in this workbook have headers on row 4 (HDR_ROW_REPORT), NOT row 1
- Row 1 contains the company title, not column headers
- Any code that scans for column headers (FY, Budget, month names) must use `HDR_ROW_REPORT` not row 1
- Found in modReconciliation: ValidateCrossSheet trendLastCol + FindFYCol both scanned row 1 — FY column search always failed
- Always search for `row 1` or `Rows(1)` in any header-scanning code and verify it should be HDR_ROW_REPORT instead

## Dynamic Sheet Discovery vs Hardcoded Lists (2026-03-07)
- modPDFExport had GetReportSheetList hardcoded to 7 sheets (only Jan-Mar monthly tabs)
- If the user builds Apr-Dec tabs later, PDF export would silently skip them
- When listing sheets for batch operations, always discover dynamically (loop through all sheets matching a pattern) rather than hardcoding names
- Exception: sheets with truly fixed names (like "P&L Summary", "Assumptions") can be hardcoded

## xlSheetVeryHidden Blocks Hyperlinks (2026-03-07)
- `xlSheetVeryHidden` hides a sheet from both the tab bar AND the Unhide dialog — but it also blocks VBA `Hyperlinks.Add` navigation
- If you need a sheet hidden but still navigable via hyperlinks or drill-down links, use `xlSheetHidden` instead
- `xlSheetHidden` = hidden from tab bar, visible in Unhide dialog, hyperlinks work
- Found in modDrillDown: HideGLSheet used xlSheetVeryHidden which broke T8.19 drill link navigation

## VBA Chr() Function Range Limitation (2026-03-11)
- VBA `Chr()` only handles character codes 0-255 (ASCII/Latin-1 range)
- Using `Chr(9472)` (box-drawing), `Chr(8212)` (em dash), or any code > 255 crashes with "Invalid procedure call or argument"
- This is different from Python's `chr()` which handles full Unicode
- Use ASCII alternatives: `String(50, "=")` instead of `String(50, Chr(9472))`
- Found in 3 locations on 2026-03-11: modSplashScreen line 85, modUTL_SplashScreen lines 39 and 71
- Previously found with `Chr(8212)` on 2026-03-04 — same root cause, same fix
- BEFORE delivering any VBA code, search for `Chr(` and verify all codes are <= 255

## Keep CLAUDE.md Updated Every Session (2026-03-11)
- CLAUDE.md is the single most important handoff document between sessions and Claude accounts
- It MUST be updated at the end of EVERY session with:
  1. Current module/tool counts (demo VBA, universal VBA, Python)
  2. New session summary with what was built, fixed, or changed
  3. Updated Repo Structure if folders were added/moved/deleted
  4. Updated Sharing Plan if module counts changed
  5. Updated Current Status with latest milestones
- If CLAUDE.md falls out of date, the next session starts with stale context and wastes time re-discovering what exists
- Check: Does CLAUDE.md mention today's date? If not, it needs updating
- Also update tasks/lessons.md with any new patterns discovered during the session
