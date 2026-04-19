# VIDEO 3 — Automated Review Prompt for Google Gemini (v2 — Post Path A)

---

## INSTRUCTIONS FOR AI REVIEWER

You are reviewing a screen-recorded demo video. This document is optimized for AI consumption, not human reading. Process these rules before watching:

1. **Watch the full video from start to finish before responding.** Do not write any response until you have seen the entire video.

2. **For every checklist item below, respond with exactly one of these three labels:**
   - `PASS` — the item is clearly visible on screen and working as specified
   - `FAIL` — the item is clearly not working, missing, or shows incorrect behavior
   - `UNSURE(reason)` — you cannot determine with confidence. Include a brief reason and a timestamp. Example: `UNSURE(dialog appeared briefly at 2:15 — could not read text)`

3. **Include a timestamp (MM:SS format) for every PASS, FAIL, and UNSURE response.**

4. **Do not skip any checkbox.** Every single item must have a response.

5. **Be literal and specific.** Do not say "looks fine" — describe exactly what you see. Do not paraphrase what the status bar says — quote it verbatim.

6. **Fill out the issue table at the bottom** with every real issue you find (not just things from the checklist — include anything that would look unprofessional on camera to 2,000+ employees and a CFO/CEO).

7. **Output format:** Your entire response must be a clean markdown document suitable for pasting directly into a bug report. No conversational filler. No "I noticed that..." or "It appears that..." — just structured findings.

8. **This is version 2 of the review after a major refactor.** Many clips used to fail. They should now work. Flag any clip that still behaves like the v1 bug report (dialogs appearing, macros not executing, screen going blank).

---

## CRITICAL CONTEXT — READ BEFORE REVIEWING

### What this video demonstrates
An Excel file loaded with universal VBA automation tools. A "Director" macro plays AI-narrated audio, navigates between sheets, runs automation tools, scrolls, and pauses — all without human interaction. The user presses one button and watches hands-free.

### The file being demonstrated
Filename: `Sample_Quarterly_ReportV2.xlsm`
This is a fictional quarterly financial report for "Keystone BenefitTech, Inc." It contains intentional data quality issues that the automation tools will find and fix. These "messy" issues are features, not bugs — they make the cleanup tools look valuable.

### Intentional mess baked into the sample file (do NOT flag these as bugs)
- Mixed date formats on the same column (01/15/2026, 2026-01-20, Jan 22 2026, 1/8/26)
- Text-stored numbers (numbers stored as text strings, may show with green triangles)
- Blank rows scattered in the middle of the data (rows 10, 12, 20, 29)
- Duplicate row somewhere in the data
- Oversized font on rows 2 and 6 (26pt and 28pt — text overflows column boundaries)
- Junk/placeholder sales rep names: "2026", "Test File Demo - Video 3", "Video 3 Test Sample File Demo", "Demo File Test Example"
- Leading/trailing spaces on some names (" Andrew Walsh", "Lisa Hernandez ")
- A #N/A error formula in one cell
- 1 negative amount (-$12,500) as a billing adjustment

### Expected sheet inventory in the file (7 visible + 1 hidden)
1. `Cover` — branded title page
2. `Q1 Revenue` — main messy sales data (~43 rows)
3. `Q1 Expenses` — clean departmental expenses
4. `Q1 Revenue v2` — similar to Q1 Revenue but with ~22 differences (for Compare demo)
5. `Budget Summary` — styled budget vs actual with variance
6. `Contact List` — employee directory (for Column Ops demo)
7. `Pivot_SalesByRegion` — pivot table: rows=Region, values=Sum of Amount
8. `Pivot_SalesByRep` — pivot table: rows=Sales Rep, values=Sum of Amount
9. `Archive_Q4_2025` — HIDDEN sheet (should not appear unless Tab Organizer unhides it)

### What changed since v1 review
The v1 recording had 50 FAIL findings mostly because SendKeys was being used to auto-fill dialog boxes, and SendKeys timing failed constantly. In v2, every dialog has been eliminated by using direct-call silent wrappers (functions named `DirectorXxx` that take parameters and skip all InputBox/MsgBox dialogs). The recording should now run with **zero dialogs visible on screen** during automation.

### Expected status bar behavior
The Excel status bar (bottom-left corner of the window) should display text messages during each clip, prefixed with `[Director]`. Examples:
- `[Director] Running: Data Sanitizer - Preview`
- `[Director] Running: Highlight by Threshold (> $100,000)`
- `[Director] Running: Compare Sheets (Q1 Revenue vs Q1 Revenue v2)`

These messages identify which tool is running. If the status bar is empty or shows generic Excel text like "Ready" during a tool's execution, that is a FAIL — the Director is not setting the status message.

### Expected runtime
8-10 minutes total

---

## CLIP-BY-CLIP CHECKLIST

Timestamps are approximate. Use them as rough guides — verify against what's actually on screen.

### Clip 27 — Opening (approx 0:00 - 0:50)
Audio: V3_S0_Opening.mp3
Expected behavior: Director starts on Q1 Revenue sheet, holds on header area for 4 seconds, scrolls down slowly, then tours Q1 Expenses → Budget Summary → Contact List → returns to Q1 Revenue.

- [ ] Video starts on a sheet named exactly "Q1 Revenue" (NOT "Cover", NOT "Q1 Revenue v2")
- [ ] Audio narration is audible (a voice describing the file)
- [ ] Screen shows a table with 9 column headers: Region | Sales Rep | Product | Customer | Date | Amount | Status | Commission % | Notes
- [ ] Messy data is visible: mixed date formats, oversized text on some rows, blank rows between data
- [ ] Director navigates through multiple sheets in sequence (Q1 Expenses, Budget Summary, Contact List visible at some point)
- [ ] Director returns to Q1 Revenue before the clip ends
- [ ] NO pre-flight dialog or "Pre-Flight Check for Video 3" popup appears (this was a v1 bug — should be fixed)
- [ ] No error dialogs appear during this clip

### Clip 28 — Data Sanitizer (approx 0:50 - 1:50)
Audio: V3_C1A_DataSanitizer.mp3
Expected behavior: Director runs silent preview wrapper, then silent full-sanitize wrapper. Both create output sheets without showing any MsgBox or InputBox.

- [ ] Status bar shows exactly: `[Director] Running: Data Sanitizer - Preview`
- [ ] A new sheet appears titled "UTL_Sanitizer_Preview" containing a report of issues found (columns: Sheet, Cell, Issue Type, Current Value, Proposed Value, Reason)
- [ ] Preview sheet shows at least 3 issues (text-stored numbers typically flagged in red)
- [ ] Status bar changes to exactly: `[Director] Running: Data Sanitizer - Full Clean`
- [ ] A data cleaning operation visibly takes place (numbers get normalized, text-stored numbers converted)
- [ ] NO "Sanitization complete!" MsgBox appears (v1 bug — should be fixed)
- [ ] NO "Run full numeric sanitizer?" Yes/No confirmation dialog appears (v1 bug — should be fixed)
- [ ] Audio plays through without getting cut off

### Clip 29 — Highlights (approx 1:50 - 2:30)
Audio: V3_C1B_Highlights.mp3
Expected behavior: Director selects range F2:F43 on Q1 Revenue, calls silent threshold highlighter (values > $100,000, direction = above), then silent duplicate highlighter, then silent clear.

- [ ] Status bar shows: `[Director] Running: Highlight by Threshold (> $100,000)`
- [ ] Cells in the Amount column with values greater than $100,000 visibly change color (green fill)
- [ ] Status bar then shows: `[Director] Running: Highlight Duplicate Values`
- [ ] Duplicate values in the Amount column get a different color (orange fill)
- [ ] Highlights are cleared at the end (cells return to normal color)
- [ ] NO "Select the range to check" Application.InputBox (range picker) dialog appears (v1 bug — should be fixed)
- [ ] NO "Clear Highlights" dialog appears (v1 bug — should be fixed)
- [ ] NO threshold value InputBox prompt appears
- [ ] NO direction choice InputBox prompt appears

### Clip 30 — Comments (approx 2:30 - 3:10)
Audio: V3_C1C_Comments.mp3
Expected behavior: Director calls silent comment extractor. Creates a "Comment Inventory" sheet. No dialogs.

- [ ] Status bar shows: `[Director] Running: Extract All Comments`
- [ ] A new sheet appears titled "Comment Inventory" containing extracted comment data
- [ ] The sheet shows exactly 5 comments (or close to that — the source file has 5 real comments)
- [ ] Columns present: # | Sheet | Cell | Cell Value | Comment Author | Comment Text
- [ ] NO "Total Comments: X" MsgBox appears
- [ ] NO "X comment(s) extracted" completion MsgBox appears

### Clip 31 — Tab Organizer (approx 3:10 - 4:00)
Audio: V3_C2A_TabOrganizer.mp3
Expected behavior: Director calls silent color-tabs-by-keyword (hardcoded to color "Revenue" tabs blue), then silent alphabetical tab reorder.

- [ ] Status bar shows: `[Director] Running: Color Tabs by Keyword (Revenue = Blue)`
- [ ] Sheet tabs at the bottom visibly change color (tabs containing "Revenue" turn blue: Q1 Revenue, Q1 Revenue v2)
- [ ] Status bar then shows: `[Director] Running: Reorder Tabs Alphabetically`
- [ ] Sheet tab order visibly changes (tabs rearrange at the bottom)
- [ ] NO keyword input prompt appears
- [ ] NO color choice InputBox appears
- [ ] NO sort-order selection dialog appears

### Clip 32 — Column Ops (approx 4:00 - 4:50)
Audio: V3_C2B_ColumnOps.mp3
Expected behavior: Director navigates to Contact List sheet, calls silent SplitColumn on range A2:A16 with space delimiter (Full Name splits into First + Last), then calls silent CombineColumns on range A2:B16 (the newly-split First and Last columns) with space separator, reconstructing the full name in a new column.

- [ ] Screen switches to "Contact List" sheet
- [ ] Status bar shows: `[Director] Running: Split Column (Full Name -> First + Last)`
- [ ] The Full Name column (A) visibly splits — column A becomes first names only (Timothy, Sarah, Raj...) and a new column B appears with last names (Regan, Mitchell, Krishnamurthy...). All other columns (Title, Department, Email, Phone, Office Location) shift right.
- [ ] Status bar then shows: `[Director] Running: Combine Columns (First + Last -> Full Name)`
- [ ] A new column appears combining First + Last back into Full Name with a space separator (e.g., "Timothy Regan", "Sarah Mitchell")
- [ ] NO range picker dialog appears
- [ ] NO delimiter choice InputBox appears
- [ ] NO separator choice InputBox appears
- [ ] NOTE: The audio narration may mention "City + State" generically — the demo instead shows reverse-combine (First + Last → Full Name) because the Contact List file does not have separate City and State columns. This is intentional; do not flag as a mismatch unless the narration explicitly contradicts what's happening on screen.

### Clip 33 — Sheet Tools (approx 4:50 - 5:45)
Audio: V3_C2C_SheetTools.mp3
Expected behavior: Director calls silent ListAllSheetsWithLinks (creates sheet index), then silent TemplateCloner to clone "Q1 Expenses" 2 times.

- [ ] Status bar shows: `[Director] Running: Create Sheet Index with Links`
- [ ] A new sheet appears containing a list of all sheet names as clickable hyperlinks
- [ ] Status bar then shows: `[Director] Running: Template Cloner (Q1 Expenses x 2)`
- [ ] Two new tabs appear at the bottom, each a copy of Q1 Expenses (likely named "Q1 Expenses (2)" and "Q1 Expenses (3)" or similar)
- [ ] NO "Which sheet do you want to clone?" input dialog appears (v1 bug — should be fixed)
- [ ] NO "How many copies?" input dialog appears (v1 bug — should be fixed)
- [ ] 0 copies created is a FAIL — 2 copies must be visible

### Clip 34 — Compare Sheets (approx 5:45 - 6:35)
Audio: V3_C3A_Compare.mp3
Expected behavior: Director calls silent CompareSheets with "Q1 Revenue" and "Q1 Revenue v2" as parameters. Creates a comparison report sheet.

- [ ] Status bar shows: `[Director] Running: Compare Sheets (Q1 Revenue vs Q1 Revenue v2)`
- [ ] A comparison report sheet appears (name likely "UTL_CompareReport")
- [ ] The report highlights differences between the two sheets (v1 had ~22 cell differences; v2 has 8 clean differences — either is acceptable as long as SOME differences are shown)
- [ ] Differences are visually distinguishable (color-coded or marked)
- [ ] NO sheet selection numbered list dialog appears

### Clip 35 — Consolidate (approx 6:35 - 7:15)
Audio: V3_C3B_Consolidate.mp3
Expected behavior: Director calls silent ConsolidateSheets with Array("Q1 Revenue", "Q1 Revenue v2"). Creates a consolidated sheet with a source-tracking column.

- [ ] Status bar shows: `[Director] Running: Consolidate Sheets (Q1 Revenue + Q1 Revenue v2)`
- [ ] A new sheet appears (name likely "UTL_Consolidated") combining data from both revenue sheets
- [ ] A "Source Sheet" column is visible showing which sheet each row came from
- [ ] Header row is styled (blue fill, white text)
- [ ] NO sheet-selection dialog appears (v1 bug — should be fixed)

### Clip 36 — Pivot Tools + Budget View (approx 7:15 - 8:15)
Audio 1: V3_C3C_PivotTools.mp3, Audio 2: V3_C3D_LookupValidation.mp3
Expected behavior: Director calls ListAllPivots. **The sample file now contains 2 real pivot tables** (Pivot_SalesByRegion and Pivot_SalesByRep), so the tool should find and report them. Then navigates to Budget Summary.

- [ ] Status bar shows: `[Director] Running: List All Pivot Tables`
- [ ] A pivot inventory sheet appears showing 2 pivot tables found (Pivot_SalesByRegion and Pivot_SalesByRep) — this is the key v2 improvement
- [ ] The inventory is NOT empty (empty = FAIL; v1 showed 0 pivots which looked broken on camera)
- [ ] Screen then navigates to "Budget Summary" sheet
- [ ] Budget Summary is visibly styled: blue header row, currency-formatted numbers, color-coded status column (green/orange/red)
- [ ] Budget Summary shows 7 departments plus a TOTAL row
- [ ] Audio plays both narration clips back-to-back without gaps

### Clip 37 — Universal Command Center (approx 8:15 - 9:05)
Audio: V3_C4_CommandCenter.mp3
Expected behavior: Director calls DirectorShowCommandCenter, which creates a styled "UTL_ToolInventory" sheet listing all available universal toolkit tools by category.

- [ ] Status bar shows: `[Director] Running: Universal Tool Inventory`
- [ ] A new sheet appears (name likely "UTL_ToolInventory") listing the universal toolkit tools
- [ ] The sheet shows categories (Data Cleaning, Columns, Sheets, Compare/Consolidate, Lookup/Validation, etc.) with tool names and descriptions
- [ ] Header row is styled
- [ ] NO InputBox command-center menu appears (v1 had an InputBox menu that often didn't trigger at all — should be replaced with a visible output sheet instead)

### Clip 38 — Closing (approx 9:05 - 10:00)
Audio: V3_Closing.mp3
Expected behavior: Director navigates to Q1 Revenue sheet, holds static, closing narration plays, completion message appears at the end.

- [ ] Screen navigates to "Q1 Revenue" sheet (NOT Cover, NOT Budget Summary, NOT any output sheet)
- [ ] Closing audio narration plays
- [ ] A "Video 3 recording complete!" MsgBox appears at the very end
- [ ] Video ends cleanly (not truncated mid-audio)

---

## OVERALL QUALITY CHECKS

- [ ] Audio narration plays throughout the entire video without gaps or clipping mid-sentence
- [ ] ZERO dialog boxes appear at any point during the recording (major improvement over v1 — every dialog = FAIL)
- [ ] Status bar messages are visible at the bottom and change with each tool execution
- [ ] Screen actions happen AFTER narration mentions them (not before)
- [ ] Scrolling is smooth and stays within the data range (no overshooting into blank cells)
- [ ] The video feels professional with no awkward pauses longer than 3 seconds
- [ ] All macro outputs are visible (new sheets created, colors applied, data modified)
- [ ] Total runtime is between 8 and 10 minutes (v1 ended early at 7:16 because macros were failing — runtime should now be full)

---

## ISSUE TABLE

Fill in every real issue found. Severity scale:
- **CRITICAL** — Something is clearly broken, blocks the demo, or shows an error on screen, OR any dialog box appears
- **MAJOR** — Looks unprofessional on camera, would confuse a viewer, or shows unintended behavior
- **MINOR** — Cosmetic only, slightly off timing, or barely noticeable

| # | Timestamp | Clip | Description | Severity |
|---|-----------|------|-------------|----------|
| 1 | | | | |
| 2 | | | | |
| 3 | | | | |
| 4 | | | | |
| 5 | | | | |
| 6 | | | | |
| 7 | | | | |
| 8 | | | | |
| 9 | | | | |
| 10 | | | | |

Add more rows if needed.

---

## FINAL ASSESSMENT

Provide all of the following in this exact format:

1. **Publish readiness:** `READY` / `NEEDS FIXES` / `NOT READY`
2. **Total PASS / FAIL / UNSURE counts** across every checklist item
3. **Top 3 strongest clips** — which ones looked most impressive visually with best audio-to-action sync
4. **Top 3 weakest clips** — which ones had the worst issues, most confusing moments, or least impressive output
5. **Specific fix recommendations** — for each FAIL or MAJOR issue, state exactly what should change in the VBA code or sample file
6. **Regression check** — Did v2 fix the v1 issues? List which v1 bugs are NOW fixed vs still present. v1 bugs to check against:
   - Pre-flight dialog at ~0:21
   - Sanitization complete MsgBox at ~1:15
   - Highlights failing to apply at ~1:41
   - Clear Highlights dialog at ~2:04
   - Tab Organizer failing to execute at ~2:40
   - Column Ops failing at ~3:12
   - Sheet Tools input dialogs at ~3:32 and ~3:43
   - Compare Sheets failing at ~4:11
   - Consolidate failing at ~4:38
   - Command Center absent at ~6:03
   - Video ending at 7:16 instead of full runtime

---

*Review document created: 2026-04-17*
*Version 2 — Post Path A silent wrapper refactor*
*For automated review by Google Gemini AI*
