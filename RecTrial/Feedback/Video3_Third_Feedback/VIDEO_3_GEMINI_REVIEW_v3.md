# VIDEO 3 — AI Review Prompt for Google Gemini (v3)

---

## INSTRUCTIONS — READ THESE FIRST, THEN WATCH THE FULL VIDEO

1. **Watch the entire video from start to finish before writing a single word.** Do not respond until you have seen all of it.

2. **For every checklist item, respond with exactly one label:**
   - `PASS` — visible on screen, working exactly as specified
   - `FAIL` — not working, missing, wrong value, wrong color, wrong text, or wrong sheet name
   - `UNSURE(reason)` — genuinely cannot determine. Include timestamp and exactly what you saw. Example: `UNSURE(status bar was partially obscured at 1:14 — could not read full text)`

3. **Include a timestamp (MM:SS) for every single response**, including PASSes.

4. **Do not skip any item.** Every checkbox must have a response.

5. **Be exact and literal.** If a checklist item says the color must be GREEN and you see RED — that is a FAIL. If a sheet name must be "UTL_Sanitizer_Preview" and you see "Sanitizer Preview" — that is a FAIL. Partial matches are FAILs.

6. **Quote exact text verbatim** when asked about status bar messages, sheet names, MsgBox text, or column headers. Do not paraphrase.

7. **Fill the issue table** with every problem found — not just checklist failures. Include anything that would look wrong, broken, or unprofessional to a CFO or CEO watching this demo.

8. **Output format:** Clean markdown only. No conversational text. No "I noticed that..." or "It appears that..." — structured findings only.

---

## CRITICAL CONTEXT — READ BEFORE REVIEWING

### What this video is
A screen recording of Excel running a "Director" macro that automates a 9-10 minute demo. It plays AI narration audio, navigates sheets, runs automation tools, and scrolls — all hands-free. The user presses one button. Zero human interaction should occur after that.

### The file being demonstrated
`Sample_Quarterly_ReportV2.xlsm` — a fictional quarterly financial report with intentional data quality issues baked in.

### DO NOT flag these as bugs — they are intentional features
- Mixed date formats on Q1 Revenue (01/15/2026 and 2026-01-20 and Jan 22 2026 and 1/8/26 — all in same column)
- Text-stored numbers with green triangle warnings
- Blank rows scattered in data
- One duplicate row
- Oversized font on rows 2 and 6 (text overflows cell boundaries)
- Junk sales rep names ("2026", "Test File Demo - Video 3", etc.)
- Leading/trailing spaces on some names
- One #N/A error cell
- One negative amount (-$12,500)

### Expected sheet tabs at start of recording
Exactly these sheets exist (order may shift after Tab Organizer runs). Actual starting order in the file:
1. `Cover`
2. `Pivot_SalesByRegion`
3. `Pivot_SalesByRep`
4. `Q1 Revenue` ← Director navigates here by name (not by tab index)
5. `Q1 Expenses`
6. `Q1 Revenue v2`
7. `Budget Summary`
8. `Contact List`
9. `Archive_Q4_2025` (hidden — should not be visible)

Note: The Director activates "Q1 Revenue" by name in Clip 27, so the tab's index in the bottom tab strip is not a PASS/FAIL criterion. Only the *name* of the currently-active tab matters.

### Zero dialogs rule
**In v2, every dialog has been eliminated.** The only MsgBox permitted in the entire recording is the closing "Video 3 recording complete!" message at the very end. Any other dialog, MsgBox, InputBox, or prompt is a CRITICAL bug.

### Status bar rule
Every tool execution should update the Excel status bar (bottom-left corner of window) with a `[Director]` prefixed message. If the status bar shows generic text like "Ready" or is blank during a tool's execution — that is a FAIL. Quote the exact text you see verbatim.

### Expected runtime
Between 8 minutes 00 seconds and 10 minutes 00 seconds total. Under 8:00 = FAIL.

### Known bugs from previous recording — watch for these specifically
These bugs were found in the last review. They may or may not be fixed. Check each one explicitly:
- Starting sheet may be named something other than "Q1 Revenue" (was "Northeast" last time)
- Status bar may show "Ready" throughout instead of `[Director]` messages (was broken last time)
- Threshold highlight color may be RED instead of required GREEN (was red last time)
- SplitColumn may target wrong column (was splitting Column H "Office Location" instead of Column A "Full Name")
- Consolidate UTL_Consolidated sheet may be missing "Source Sheet" column (was missing last time)
- Template Cloner copies may be pre-existing tabs from a dry run rather than dynamically created
- Sanitizer sheet may be named "Sanitizer Preview" instead of "UTL_Sanitizer_Preview"
- Comment Inventory may be missing "#" index as first column
- Closing MsgBox text may be wrong (was "Video 3 Done. Stop OBS recording now." instead of "Video 3 recording complete!")

---

## CLIP-BY-CLIP CHECKLIST

Timestamps are approximate. Verify against what's actually on screen.

---

### Clip 27 — Opening (approx 0:00 - 0:50)
**Audio:** V3_S0_Opening.mp3
**Expected:** Director starts on Q1 Revenue, holds on data, scrolls down, tours Q1 Expenses → Budget Summary → Contact List → returns to Q1 Revenue.

- [ ] The very first sheet visible at 0:00 is named exactly **"Q1 Revenue"** — not "Cover", not "Q1 Revenue v2", not "Northeast", not any other name. Quote the exact tab name you see.
- [ ] Audio narration is audible within the first 30 seconds
- [ ] The data table has exactly **9 column headers** in this exact order: `Region | Sales Rep | Product | Customer | Date | Amount | Status | Commission % | Notes` — if any column is missing or in wrong order, FAIL and state which column is missing
- [ ] Messy data is visible: mixed date formats in Date column, at least one row with oversized text, at least one blank row
- [ ] Director navigates to at least 3 of these sheets during the tour: Q1 Expenses, Budget Summary, Contact List
- [ ] Director returns to a sheet named exactly **"Q1 Revenue"** before 0:50
- [ ] NO dialog box, popup, or MsgBox appears during this clip (zero dialogs = v2 requirement)
- [ ] The pre-flight check popup does NOT appear (was a v1 bug at ~0:21)

---

### Clip 28 — Data Sanitizer (approx 0:50 - 1:50)
**Audio:** V3_C1A_DataSanitizer.mp3
**Expected:** Silent preview runs, creates UTL_Sanitizer_Preview sheet. Silent full clean runs. No dialogs.

- [ ] Status bar shows **exactly**: `[Director] Running: Data Sanitizer - Preview` — quote the exact text you see; if it shows "Ready" that is a FAIL
- [ ] A new sheet appears named **exactly** `UTL_Sanitizer_Preview` — if it is named "Sanitizer Preview" (missing UTL_ prefix) that is a FAIL
- [ ] The preview sheet contains at least 3 rows of issue data (text-stored numbers, formatting issues)
- [ ] Status bar then shows **exactly**: `[Director] Running: Data Sanitizer - Full Clean` — quote exact text
- [ ] Data cleaning visibly occurs on Q1 Revenue sheet (values normalize, text-stored numbers convert)
- [ ] NO "Sanitization complete!" MsgBox appears
- [ ] NO "Run full numeric sanitizer?" or any Yes/No confirmation dialog appears
- [ ] Audio plays through completely without cutoff

---

### Clip 29 — Highlights (approx 1:50 - 2:30)
**Audio:** V3_C1B_Highlights.mp3
**Expected:** Silent threshold highlighter runs on F2:F43 (Amount column). Values >$100,000 turn GREEN. Silent duplicate highlighter runs. Duplicate values turn ORANGE. Silent clear runs.

- [ ] Status bar shows **exactly**: `[Director] Running: Highlight by Threshold (> $100,000)` — quote exact text
- [ ] Cells in the Amount column with values greater than $100,000 visibly change to **BRIGHT SATURATED GREEN** fill (Excel "Status Good" green — clearly, unambiguously green, NOT pale mint, NOT yellow-green, NOT a pastel). If cells change to RED, YELLOW, or any muted pastel, that is a FAIL — state the actual color you see
- [ ] Status bar then shows **exactly**: `[Director] Running: Highlight Duplicate Values` — quote exact text
- [ ] Duplicate values in the Amount column get **BRIGHT PURE ORANGE fill** (Excel "Orange" — NOT amber, NOT gold, NOT red, NOT yellow, NOT pink). Must be clearly orange and clearly different from the green above. State actual color if not pure orange.
- [ ] All highlights are cleared at the end of this clip (cells return to white/normal fill)
- [ ] NO range picker dialog or "Select the range to check" InputBox appears
- [ ] NO threshold value InputBox prompt appears
- [ ] NO direction choice InputBox prompt appears
- [ ] NO "Clear Highlights" dialog appears

---

### Clip 30 — Comments (approx 2:30 - 3:10)
**Audio:** V3_C1C_Comments.mp3
**Expected:** Silent comment extractor runs. Creates "Comment Inventory" sheet with exactly 5 rows of comment data.

- [ ] Status bar shows **exactly**: `[Director] Running: Extract All Comments` — quote exact text
- [ ] A new sheet appears named **exactly** `Comment Inventory`
- [ ] The sheet contains data rows for **exactly 5 comments** — if it shows more or fewer, state the exact count
- [ ] The sheet has these **6 columns in this exact order**: `# | Sheet | Cell | Cell Value | Comment Author | Comment Text` — if the "#" column is missing or columns are in wrong order, that is a FAIL
- [ ] NO "Total Comments: X" MsgBox appears
- [ ] NO "X comment(s) extracted" completion MsgBox appears

---

### Clip 31 — Tab Organizer (approx 3:10 - 4:00)
**Audio:** V3_C2A_TabOrganizer.mp3
**Expected:** Silent color wrapper runs — tabs with "Revenue" in name turn blue. Silent reorder wrapper runs — tabs rearrange alphabetically.

- [ ] Status bar shows **exactly**: `[Director] Running: Color Tabs by Keyword (Revenue = Blue)` — quote exact text
- [ ] Sheet tabs containing the word "Revenue" visibly change to **blue color** at the bottom of the screen — name the tabs that changed color
- [ ] Status bar then shows **exactly**: `[Director] Running: Reorder Tabs Alphabetically` — quote exact text
- [ ] Sheet tab order visibly changes (tabs move to different positions at the bottom)
- [ ] NO keyword InputBox dialog appears
- [ ] NO color choice dialog appears
- [ ] NO sort-order selection dialog appears

---

### Clip 32 — Column Ops (approx 4:00 - 4:50)
**Audio:** V3_C2B_ColumnOps.mp3
**Expected:** Director switches to Contact List sheet. Silent SplitColumn runs on Column A (Full Name). Column A splits into two new columns showing first names and last names separately. Silent CombineColumns runs.

- [ ] Screen switches to **"Contact List"** sheet
- [ ] Status bar shows **exactly**: `[Director] Running: Split Column (Full Name -> First + Last)` — quote exact text
- [ ] **Column A ("Full Name")** is the column that gets split — if Column H ("Office Location") or any other column is split instead, that is a CRITICAL FAIL — state which column was actually split
- [ ] After the split, Column A contains first names only and a new adjacent column contains last names (or vice versa) — the full names are no longer combined in one cell
- [ ] Status bar then shows **exactly**: `[Director] Running: Combine Columns (First + Last -> Full Name)` — quote exact text
- [ ] A combine operation visibly occurs on screen (columns merge)
- [ ] NO range picker dialog appears
- [ ] NO delimiter choice InputBox appears
- [ ] NO separator choice InputBox appears

---

### Clip 33 — Sheet Tools (approx 4:50 - 5:45)
**Audio:** V3_C2C_SheetTools.mp3
**Expected:** Silent ListAllSheetsWithLinks creates an index sheet. Silent TemplateCloner creates 2 new copies of Q1 Expenses as new tabs — these tabs must appear dynamically during the recording, not be pre-existing.

- [ ] Status bar shows **exactly**: `[Director] Running: Create Sheet Index with Links` — quote exact text
- [ ] A new sheet appears containing a list of all sheet names as clickable hyperlinks
- [ ] Status bar then shows **exactly**: `[Director] Running: Template Cloner (Q1 Expenses x 2)` — quote exact text
- [ ] **Two new Q1 Expenses copy tabs appear during the recording** — if these tabs were already present before the macro ran (pre-existing from a dry run), that is a MAJOR FAIL — state whether they appeared dynamically or were already there
- [ ] The 2 new tabs are named something like "Q1 Expenses (2)" and "Q1 Expenses (3)" or similar copy names
- [ ] NO "Which sheet do you want to clone?" InputBox appears
- [ ] NO "How many copies?" InputBox appears

---

### Clip 34 — Compare Sheets (approx 5:45 - 6:35)
**Audio:** V3_C3A_Compare.mp3
**Expected:** Silent CompareSheets runs with "Q1 Revenue" vs "Q1 Revenue v2". Creates UTL_CompareReport sheet showing differences.

- [ ] Status bar shows **exactly**: `[Director] Running: Compare Sheets (Q1 Revenue vs Q1 Revenue v2)` — quote exact text
- [ ] A comparison report sheet appears — state the exact sheet name you see
- [ ] The report shows at least 5 differences between the two revenue sheets (could be up to 22 depending on sample file version)
- [ ] Differences are visually marked or color-coded
- [ ] NO sheet selection or numbered list dialog appears

---

### Clip 35 — Consolidate (approx 6:35 - 7:15)
**Audio:** V3_C3B_Consolidate.mp3
**Expected:** Silent ConsolidateSheets runs with Q1 Revenue + Q1 Revenue v2. Creates UTL_Consolidated sheet. First column must be "Source Sheet" showing which sheet each row came from.

- [ ] Status bar shows **exactly**: `[Director] Running: Consolidate Sheets (Q1 Revenue + Q1 Revenue v2)` — quote exact text
- [ ] A new consolidated sheet appears — state the exact sheet name you see
- [ ] The **first column is labeled "Source Sheet"** (or similar source-tracking label) showing "Q1 Revenue" or "Q1 Revenue v2" in each row — if Column A starts with "Region" instead, that is a CRITICAL FAIL
- [ ] Source Sheet column contains exactly two distinct values: **"Q1 Revenue"** and **"Q1 Revenue v2"** (the "v2" suffix must be read literally — it is NOT "Q2 Revenue"). If you believe you see "Q2 Revenue" anywhere, zoom in and re-read carefully — the sample file contains no Q2 Revenue sheet.
- [ ] The consolidated sheet contains data rows from both revenue sheets combined (should be 80+ rows total)
- [ ] Header row is styled with blue fill and white text
- [ ] NO sheet-selection dialog appears

---

### Clip 36 — Pivot Tools + Budget View (approx 7:15 - 8:15)
**Audio 1:** V3_C3C_PivotTools.mp3 | **Audio 2:** V3_C3D_LookupValidation.mp3
**Expected:** ListAllPivots runs. Finds 2 real pivot tables. Then navigates to Budget Summary.

- [ ] Status bar shows **exactly**: `[Director] Running: List All Pivot Tables` — quote exact text
- [ ] A pivot inventory sheet appears — state the exact sheet name
- [ ] The inventory shows **exactly 2 pivot tables** listed: `Pivot_SalesByRegion` and `Pivot_SalesByRep` — if it shows 0 or any other count, that is a FAIL — state exact count found
- [ ] **NO "X pivot table(s) listed on 'UTL_PivotReport' sheet." MsgBox appears** — this was a CRITICAL bug in v2.3 that blocked the whole recording. A silent `DirectorListAllPivots` wrapper now replaces the dialog-heavy `ListAllPivots`. Any pivot-related MsgBox anywhere in this clip = CRITICAL FAIL.
- [ ] Screen navigates to **"Budget Summary"** sheet
- [ ] Budget Summary has a **blue header row** with white text
- [ ] Numbers in Budget Summary are **currency-formatted** ($ signs, commas)
- [ ] Status column is **color-coded** (green for Under Budget, orange/red for Over Budget — or similar)
- [ ] Budget Summary shows **7 department rows plus a TOTAL row** (8 rows of data total)
- [ ] Both audio clips play back-to-back without a gap between them

---

### Clip 37 — Universal Command Center (approx 8:15 - 9:05)
**Audio:** V3_C4_CommandCenter.mp3
**Expected:** DirectorShowCommandCenter creates UTL_ToolInventory sheet listing all tools by category.

- [ ] Status bar shows **exactly**: `[Director] Running: Universal Tool Inventory` — quote exact text
- [ ] A new sheet appears named **exactly** `UTL_ToolInventory` — state the exact name you see
- [ ] The sheet shows tool categories such as: Data Cleaning, Columns, Sheets, Compare/Consolidate, Lookup/Validation
- [ ] The sheet has a styled header row
- [ ] NO InputBox or command-center popup menu appears (v1 had an InputBox that often failed — should be completely replaced by the sheet output)

---

### Clip 38 — Closing (approx 9:05 - 10:00)
**Audio:** V3_Closing.mp3
**Expected:** Director navigates to Q1 Revenue. Holds static. Closing narration plays. MsgBox appears at the very end with exact text "Video 3 recording complete!"

- [ ] Screen navigates to a sheet named **exactly "Q1 Revenue"** — not Cover, not Budget Summary, not any output sheet, not "Northeast". Quote the exact tab name you see. NOTE: The first data cell on Q1 Revenue is "Northeast" (a region value). Do not confuse the cell content with the tab name — read the tab at the bottom of the window, not cell A2.
- [ ] Closing audio narration plays fully
- [ ] Status bar shows `[Director] Almost done — keep recording until the 'Video 3 recording complete!' MsgBox appears` between the end of audio and the final MsgBox — this is an EXPECTED, deliberate cue for the user and must NOT be flagged as an unauthorized dialog. It appears in the status bar (bottom of Excel window), not as a popup.
- [ ] A MsgBox appears at the very end — state the **exact text** of the MsgBox verbatim. Required text is: `"Video 3 recording complete!"` — any other text is a FAIL. If the video cuts off before the MsgBox appears, that is a recording-side issue (OBS stopped too early), not a VBA failure — flag it as "MsgBox not visible in recording" rather than "MsgBox missing".
- [ ] Video ends cleanly (not truncated mid-audio, not cut off)

---

## OVERALL QUALITY CHECKS

- [ ] Audio narration plays throughout the entire video without gaps or mid-sentence cutoffs
- [ ] **ZERO dialog boxes** appear during the recording — the ONLY permitted MsgBox is "Video 3 recording complete!" at the very end. Any other dialog anywhere = CRITICAL
- [ ] Status bar shows `[Director]` prefixed messages during every tool execution — if it shows "Ready" throughout, that is a single MAJOR failure covering the entire video
- [ ] Screen actions happen AFTER narration mentions them (not before)
- [ ] Scrolling is smooth and stays within data range (no overshooting into blank rows)
- [ ] All macro outputs are visible: new sheets created with real data, colors applied, columns modified
- [ ] Total runtime is **between 8:00 and 10:00** — state exact runtime to the second

---

## ISSUE TABLE

For every issue found — checklist failures AND anything else that looks wrong on camera. Severity:
- **CRITICAL** — Breaks the demo, shows an error, a dialog appears, or macro targets wrong data
- **MAJOR** — Looks unprofessional, wrong color/name/text, missing key data, or timing badly off
- **MINOR** — Cosmetic, slightly off, barely noticeable

Pre-fill with at least 15 rows. Add more if needed.

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
| 11 | | | | |
| 12 | | | | |
| 13 | | | | |
| 14 | | | | |
| 15 | | | | |

---

## FINAL ASSESSMENT

Provide all of the following in this exact format:

1. **Publish readiness:** `READY` / `NEEDS FIXES` / `NOT READY`
2. **Total PASS / FAIL / UNSURE counts** across every checklist item above
3. **Top 3 strongest clips** — most impressive visually, best audio-to-action sync
4. **Top 3 weakest clips** — worst issues, most confusing, or least impressive
5. **Specific fix recommendations** — for each FAIL or MAJOR issue, state exactly what needs to change (VBA code, sample file, or pre-recording steps)
6. **Regression check** — for each item below, state FIXED or STILL PRESENT:
   - Pre-flight dialog at ~0:21
   - Sanitization complete MsgBox at ~1:15
   - Highlights failing to apply at ~1:41
   - Clear Highlights dialog at ~2:04
   - Tab Organizer failing to execute at ~2:40
   - Column Ops failing at ~3:12
   - Sheet Tools InputBox dialogs at ~3:32 and ~3:43
   - Compare Sheets failing at ~4:11
   - Consolidate failing at ~4:38
   - Command Center absent at ~6:03
   - Video ending early at 7:16
   - Starting sheet named incorrectly (was "Northeast" in v2.1)
   - Status bar showing "Ready" throughout entire video (was broken in v2.1)
   - Highlight color being RED instead of GREEN (was wrong in v2.1)
   - Column Ops splitting wrong column H instead of A (was wrong in v2.1)
   - Consolidate missing Source Sheet column (was missing in v2.1)
   - Template Cloner showing pre-existing tabs instead of dynamically creating them (was wrong in v2.1)
   - Closing MsgBox showing wrong text (was "Video 3 Done. Stop OBS recording now." in v2.1)

---

*Review document version: 3.3*
*Updated: 2026-04-20*
*Previous reviews: v1 (18 PASS / 50 FAIL), v2.1 (56 PASS / 31 FAIL), v2.2 (51 PASS / 39 FAIL), v2.3 (80 PASS / 12 FAIL)*
*For use with Google Gemini 2.5 Pro — set temperature to 0 before submitting*
