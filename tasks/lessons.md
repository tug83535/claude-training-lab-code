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
