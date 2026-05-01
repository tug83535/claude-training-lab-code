# Lessons Learned - APCLDmerge Project

## V4 Package Build + Testing — 2026-05-01

- **`Dir()` gives false positives on OneDrive Files On-Demand paths.** OneDrive's "Files On-Demand" feature makes cloud-only placeholder files appear to exist locally. `Dir("path\to\file")` returns the filename even when the file is not downloaded. This caused the VBA guard to skip the MsgBox and call `Shell()` on a non-existent python.exe. Fix: use `CreateObject("Scripting.FileSystemObject").FileExists(path)` instead of `Dir()`. FSO correctly returns False for cloud-only placeholders.

- **VBA `Shell()` command string quoting breaks on complex Windows paths.** `Shell "cmd.exe /k ""path"" ""path2""", vbNormalFocus` produces "The filename, directory name, or volume label syntax is incorrect" when paths contain spaces or OneDrive-style names. Even `WScript.Shell.Run` with quoted paths can fail. Most reliable fix: use `wsh.CurrentDirectory = ThisWorkbook.Path` to set the working directory, then use relative paths with no spaces in the `Run` command: `wsh.Run "cmd.exe /k python\python-embedded\python.exe scripts\finance_automation_launcher.py"`. No quoting needed when using relative paths.

- **`safe_io.py` resolves `_TOOLKIT_ROOT` as two levels up from `common/safe_io.py`.** In development (`ZeroInstall/common/safe_io.py`), two levels up = `ZeroInstall/`. In the deployed package (`scripts/common/safe_io.py`), two levels up = `scripts/`. This means `samples/` and `outputs/` land at `scripts/samples/` and `scripts/outputs/` — NOT at the package root. The assembly guide must put `samples\` inside `scripts\`, not at the `FinanceTools_v1.0\` root. Always verify Python path resolution logic before writing an assembly guide.

- **Zero-install bundled Python path: always open FinanceTools.xlsm from inside the package folder.** `ThisWorkbook.Path` resolves to wherever the .xlsm is opened from. If the user opens an old copy from OneDrive or another location, the relative paths to `python\` and `scripts\` will be wrong. Confirmed: opening from `RecTrial\FinanceTools_v1.0\` with all subfolders in place gives correct behavior.

## V4 Delivery Model — VBA Shell() + Bundled Python — 2026-04-28

- **VBA Shell() works on iPipeline laptops.** Confirmed during V4 planning: Excel can call `Shell()` to launch a `.py` file through a local Python executable. This eliminates the need for coworkers to open a Command Prompt. They click an Excel button and the script runs invisibly. This was not obvious at the start of V4 planning and changes the entire distribution approach.

- **Bundled Python 3.11 embeddable = zero install for coworkers.** The Python 3.11 embeddable zip (~10 MB) is a standalone folder — no installation, no admin rights, no PATH changes. Drop it next to the workbook. VBA Shell() points at `python-embedded\python.exe`. Coworkers never touch Python setup. Confirmed this is the V1 delivery model.

- **The delivery path: `ThisWorkbook.Path & "\python\python-embedded\python.exe"`**. All paths must be relative to the workbook location so the zip works wherever it's unzipped. Do not hardcode drive letters or user paths.

- **Excel button design decision: one launcher button OR individual per-tool buttons.** A single button that opens `finance_automation_launcher.py` (numbered menu) is simpler to build. Per-tool buttons (one per script) give coworkers faster access but require more VBA wiring. This decision gates the FinanceTools.xlsm build — ask Connor before building.

- **Stdlib-only Python is the right default for coworker distribution.** All 6 V4 scripts use only Python standard library — csv, json, pathlib, datetime, hashlib, zipfile, xml.etree, subprocess, re, collections, difflib. No pip required. This means the bundled embeddable Python works with no additional packages. If coworkers have pip access, pandas/openpyxl could be added later, but V1 ships with zero dependencies.

- **Analysis date anchoring prevents false positives as real time passes.** revenue_leakage_finder.py originally used `date.today()` to detect stale contracts. This caused Class 2 (stale contracts) to over-count massively once real calendar time passed beyond the sample data window. Fix: derive "today" from `max(billing_period dates) + 45 days`. The script now gives correct results regardless of when it's run, as long as the billing data is consistent.

- **word-boundary regex + known_sheets filter for cross-sheet references.** workbook_dependency_scanner.py regex alone cannot reliably distinguish Excel sheet names from formula text like "SUM(Data" or "B12-Data". Fix: regex captures candidates, then filter against the set of actual sheet names from workbook.xml. Only names that appear in both the formula AND the sheet list are true cross-sheet references.

- **Timestamped output folders prevent overwrite accidents.** Every V4 script writes to `outputs/YYYYMMDD_HHMMSS_toolname/`. Previous runs are never touched. This is the correct pattern for coworker-facing tools — they can re-run anytime without fear of losing prior results.

## Planning and Research Workflow Lessons — 2026-04-23

- **When parallel AI sessions produce overlapping research, compile first, review second.** User ran parallel claude.ai / Codex / other-AI sessions that produced 14 raw research files. Trying to read all 14 in a single Claude Code session is wasteful. The right move is: compile them into 3-6 structured synthesis docs using targeted prompts, then review the synthesis rather than the raw files. A subagent back-check of the raw files afterward (to verify nothing was missed) takes 5 min and confirmed HIGH confidence the synthesis was complete.

- **One master overview doc beats 8 scattered planning docs.** Before today, project state lived across: CLAUDE.md, memory/project_status.md, todo.md, lessons.md, CHERRY_PICK_TRACKER.md, VIDEO_4_DRAFT_IDEAS.md, VIDEO_4_CURRENT_PROPOSAL.md, FUTURE_AUTOMATION_IDEAS.md. Each served a specific audience. But when a user wants a second-opinion AI review, they need ONE file. Built `RecTrial\PROJECT_OVERVIEW.md` as a 15-section point-in-time snapshot with Section 14 explicitly listing "angles to push on" for a reviewer. Makes AI review cheap and reproducible.

- **Snapshot before a long-running subagent task.** User asked me to do a 14-file research review that could take 5+ min. Before kicking off the subagent, snapshot the current proposal/thinking to disk + update memory files. This way if context gets compressed or the session needs to restart, nothing is lost. Habit: any task over ~3 min, snapshot first.

- **"Gemini perception vs actual code correctness" principle.** Across 5 Video 3 Gemini review cycles, several flagged "bugs" (RGB 255,140,0 read as red, "Q1 Revenue v2" misread as "Q2 Revenue") were video-compression perception artifacts, not real code issues. Two responses available: (a) make the signal MORE extreme to beat compression (brighter orange, wider column, bolder font), OR (b) accept imperfect rubric scores and ship. User chose (b) at 70/4 Gemini score. Right call; perfection was not the bar.

- **Pull the plug on a video plan before recording if it doesn't excite the user.** Original Video 4 (10 CMD-run scripts with ElevenLabs audio) had all assets built — MP3 clips generated, demo files created, interactive guides written. When user said "I want a Video 4 that is actually useful," I should have asked "do you want to start over?" sooner. Sunk-cost on pre-built assets is not a reason to record a video the creator isn't excited about. All V4 assets remain on disk — repurposable for the new plan or discardable.

- **Videos for a 2,000-person non-dev audience need a SINGLE-FILE downloadable deliverable, not a CMD-based demo.** Every doc's research consensus: CFO won't tolerate seeing Command Prompt on screen, and coworkers won't run 8 scripts individually. The deliverable must be a menu-driven launcher (Combo 1) or an Excel-native button interface (Combo 2) or the hero-reveal artifact is explicit "here's the full output" (Combo 3). V4 redesign absorbs all three principles via split 4a+4b format.

## Codex Cherry-Pick Campaign Lessons — 2026-04-21

- **Prefer candidate-outer over column-outer in header-name search.** `FindColumnByHeaderText` originally iterated columns outer and matched the first substring hit. On Budget Summary (with two columns: "Status" and "Materiality Status"), the generic "Status" in Column F beat the more specific "Materiality Status" in Column G because F came first. Fix: iterate candidates outer, columns inner — then specific candidates win even when they sit further right. Commit 8eff337.

- **Command Center auto-discovery works but visibility is capped.** The LoadRegistry + AutoDiscoverTools flow picks up new `modUTL_*` modules automatically and lists them as "(Discovered)" categories. Problem: the LaunchCommandCenter input dialog only shows ~29 categories visually before cutting off, and it does NOT scroll — users navigate by typing a number. That means any Discovered category past position 29 is invisible to non-experts. **Fix pattern:** for tools coworkers should see, add them to the static registry in `LoadBuiltInTools` (creates a visible entry in the first 22ish positions). Auto-discovery still runs as a backup. Commit e05dced promoted modUTL_Intelligence to static category #6 "Intelligence (3 tools)".

- **Application.Run with Range parameters is fragile.** Worked for HighlightThreshold, broke silently for SplitColumn. Variant/Range marshaling through Application.Run is unreliable. Always prefer string parameters (sheet name + range address string) and resolve the Range inside the callee. Already logged in Video 3 lessons but worth repeating.

- **Static registry promotion strategy for new universal modules.** When porting a new `modUTL_*` module that coworkers should reach via the Command Center: (1) add a `'=== CATEGORY NAME (N tools) ===` block in `LoadBuiltInTools` in modUTL_CommandCenter.bas, (2) call `AddTool` for each public sub with a friendly name + description, (3) position it alphabetically or topically within the first 22 categories for maximum visibility. Auto-discovery continues to cover it too.

- **Gemini perception vs. actual code correctness.** Some "bugs" Gemini flags are visual compression artifacts (RGB(255,140,0) reading as red, "Q1 Revenue v2" reading as "Q2 Revenue"). If the code is objectively correct, two options: (a) change the code to be more visually unmistakable (brighter saturated colors, bolder/wider text) or (b) accept and ship. User chose (b) for Video 3 v2.4 after 4 review cycles — correct call; perfection was not the bar.

- **Video title card design automation.** `RecTrial\VideoTitleCards\generate_title_cards.py` uses Pillow to produce all 5 cards (V1-V4 + disclaimer) from scratch with iPipeline brand colors. Rerunnable for quick updates. Lesson: when the "same design, different text" pattern emerges, generate from code rather than edit images pixel-by-pixel.

## Video 3 v2.2 Gemini Review Findings — 2026-04-19
- **Always set `Application.DisplayStatusBar = True` inside StatusMsg** (not just per-RunVideo). If Excel's display setting has the status bar hidden, every `Application.StatusBar = "..."` silently succeeds but renders nothing. Single-source fix: add the DisplayStatusBar call inside the StatusMsg helper so every status update guarantees visibility.
- **Gemini may conflate a prominent cell value with the tab name.** The "Northeast" regression wasn't a file bug — cell A2 of Q1 Revenue happens to be "Northeast" (the first region), and it was the biggest visible text on screen. Fix: keep column widths modest so more of the data and the actual tab are both visible in frame.
- **Sample file column widths matter for AI review.** Q1 Revenue had Column G (Status) at width 42, which pushed Column I (Notes) off-screen, causing Gemini to flag "Notes column missing." Narrowed G to 14.
- **Silent `Application.Run` calls swallow missing-sub errors.** Every Director wrapper call is wrapped in `On Error Resume Next ... On Error GoTo 0`. If the target `Director*` sub doesn't exist in the opened workbook (because the .bas file wasn't re-imported), the call silently does nothing. Always re-import updated .bas files into the .xlsm before retesting.
- **DirectorConsolidateSheets Source Sheet column placement matters.** Gemini expects "Source Sheet" as Column A (first column), not appended at the end after data columns. Users reading a consolidated sheet scan left-to-right; source tracking belongs first.
- **Multi-line MsgBox text confuses AI review.** The v2.2 closing MsgBox was `"Video 3 recording complete!" & vbCrLf & vbCrLf & "Stop OBS recording now."` — Gemini read the combined text as "Video 3 recording now complete." Fix: single-line MsgBox text when AI review is expected to verify literal string match.
- **Compare report color-coding HELD OFF (2026-04-19)** — Gemini flagged MAJOR but user chose to defer until we see if other fixes resolve Gemini's complaints. If v2.3 review still flags, add color-coded header + diff rows to `BuildCompareReport` helper in modUTL_Compare.bas.

## Video Recording Automation (Video 3 / 4 / Director Macro) — 2026-04-15
- **NEVER use SendKeys against modal dialogs.** `Application.InputBox` with `Type:=8` (range picker) cannot be filled by SendKeys — it blocks waiting for mouse selection. Regular `InputBox` sometimes works with SendKeys but timing is unreliable. Sequential multi-dialog staging (dismiss MsgBox, then fill InputBox, then pick direction) fails constantly.
- **The Path A pattern is the only reliable approach for dialog-heavy macros:** add a `DirectorXxx` silent wrapper sub at the bottom of each UTL/module .bas file that takes parameters directly and replicates the core logic with NO InputBox/MsgBox. Example: `Sub DirectorHighlightThreshold(rng As Range, threshold As Double, direction As Long)`. Then the Director calls `Application.Run "DirectorXxx", param1, param2, ...`. Video 2 Clips 22 and 23 prove this pattern works (SaveCopyAs direct, RunWhatIfPreset, RestoreBaselineSilent).
- **Reason:** Clip 22/23 went from fragile SendKeys ("y{ENTER}", "1{ENTER}") to bulletproof silent calls. Gemini reviewed Video 3 after the SendKeys approach: 18 PASS / 50 FAIL. After switching to Path A wrappers: the clips will execute the core logic directly with no user interaction.
- **When adding a new UTL macro that shows dialogs:** also add the Director wrapper at the same time. Don't ship without both.
- **Keep UTL macros untouched for coworkers.** The DirectorXxx wrappers are ADDED at the bottom of the file — existing subs that real users call stay exactly as they were.
- **openpyxl CANNOT create Excel PivotTable objects.** For demo files that need pivots, either create them manually in Excel (2 minutes) or use a Copilot prompt. Don't try to fake it with Python.
- **mciSendString audio duration drift:** measure clip duration at runtime (`status alias length`), don't hardcode seconds. Fragile timing constants break when audio clips are regenerated.
- **Always call ResetMCI at the top of every public entry point** (RunVideo1, RunVideo2, RunVideo3, QuickTest, TestClip, RunPreflight). Interrupted runs leave the MCI device stuck open and silent-fail future audio.
- **WaitForAudioEnd beats fragile math.** Polling `status alias mode` until it's not "playing" is more reliable than `WaitSec m_ClipDurSec - X` math that breaks when scroll speeds change.
- **Keep backups before major refactors.** `VBABackup_PrePathA\` saved 10 files before the Path A refactor in case we needed to revert. Free insurance.
- **Sync VBA files when edited:** the Director has two copies at `RecTrial\VBAToImport\modDirector.bas` AND `RecTrial\DemoVBA\modDirector.bas`. After any edit, always `cp` to sync.

## Project Folder Structure (Learned 2026-04-16)
- **RecTrial is the active working folder**, NOT the repo. Contains all audio clips, VBA files being imported, sample Excel files, demo inputs, and recording output folders.
- **The repo at `C:\Users\connor.atlee\.claude\projects\claude-training-lab-code\` is source of truth for commits.** RecTrial changes need to be copied back to the repo when stable.
- **Memory folder is linked to the repo path** (Claude Code auto-links based on working directory). If the repo moves, memory breaks.
- **Never delete the repo's `.git` folder.** Once lost, commit history is gone.
- **Duplicate repo folders (Old1projects issue):** if two copies of a repo exist at different paths, git may create a "new empty repo" at one location while all history lives at the other. Always verify with `git log --oneline` before deleting anything.

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

## VBA RGB Color Constant Verification (2026-03-12)
- VBA `RGB(R,G,B)` computes as `R + G*256 + B*65536` — verify the Long constant matches
- iPipeline Blue = `RGB(11, 71, 121)` = `11 + 71*256 + 121*65536` = **7,948,043**
- Found CLR_HDR = 7,930,635 in 5 modules — decodes to RGB(11, **3**, 121), green channel 3 instead of 71 (near-black vs iPipeline Blue)
- When copy-pasting color constants to new modules, always verify the Long value against the RGB formula
- Quick check: `7948043 Mod 256 = 11` (R), `(7948043 \ 256) Mod 256 = 71` (G), `7948043 \ 65536 = 121` (B)

## Sheet Index Shifting During .Move Operations (2026-03-12)
- When reordering sheets by index, each `.Move` call changes the indices of ALL other sheets
- If user enters indices 5, 3, 7 and you move sheet 5 first, sheets 3 and 7 now have different positions
- Fix: resolve ALL user-entered indices to sheet **names** first, then `.Move` by name (names don't shift)
- Same pattern applies to any batch sheet operation using indices: delete, copy, hide

## Consolidation Source Column Consistency (2026-03-12)
- When consolidating multiple sheets with different column widths, the "Source Sheet" column must be placed consistently
- If placed at `srcLastCol + 1` per sheet, sheets with fewer columns put the source tag in column 6 while wider sheets put it in column 10 — data misaligns
- Fix: pre-scan ALL source sheets to find the maximum column width, then use `maxColWidth + 1` for every sheet

## Large Range Safety Caps (2026-03-12)
- `ReDim vals(1 To rng.Cells.Count)` will overflow if `rng` covers a full-sheet selection (16K+ columns x 1M+ rows)
- VBA `Long` max is ~2.1 billion but `ReDim` with massive arrays causes out-of-memory crashes
- Use `rng.Cells.CountLarge` (returns LongLong) for the check, and cap at a reasonable limit (e.g., 500,000 cells)
- Always add a user-friendly MsgBox explaining the limit rather than silently crashing

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
