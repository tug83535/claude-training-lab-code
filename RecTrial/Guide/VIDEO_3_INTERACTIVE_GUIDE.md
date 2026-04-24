# VIDEO 3 — Interactive Recording Guide (v2 — Post Path A)

**Use this document with Claude.ai to build an interactive follow-along checklist.**

**What Video 3 Is:** The Director macro automates the entire recording. It runs universal VBA toolkit tools on a messy sample Excel file with zero dialog boxes. You press one button and watch hands-free for ~9-10 minutes.

**File:** Sample_Quarterly_ReportV2.xlsm (in RecTrial\SampleFile\SampleFileV2\)

**Major change from v1:** Every dialog has been eliminated. All tools now call silent "DirectorXxx" wrapper subs that take parameters directly — no SendKeys, no InputBox, no MsgBox. If you see ANY dialog during recording, that's a bug.

---

## PRE-RECORDING SETUP

### Step 1: Prepare Clean Sample File
- [ ] Copy fresh `Sample_Quarterly_ReportV2.xlsx` from `RecTrial\SampleFile\SampleFileBackup_nonMacroClean\` to `RecTrial\SampleFile\SampleFileV2\` (overwrite if needed)
- [ ] Open the file in Excel
- [ ] Click Enable Content if prompted

### Step 2: Add Pivot Tables (Required)
The sample file needs 2 pivot tables so Clip 36's "List All Pivot Tables" demo has real data to show.
- [ ] Click any cell in Q1 Revenue data
- [ ] Insert → PivotTable → in the range box, type `'Q1 Revenue'!$A$1:$I$43` (expand from default to full range)
- [ ] Select New Worksheet → OK
- [ ] Drag **Region** to Rows, **Amount** to Values
- [ ] If it shows "Count of Amount" in Values, click it → Value Field Settings → Sum → OK
- [ ] Rename the new sheet to exactly `Pivot_SalesByRegion`
- [ ] Click any cell in Q1 Revenue data again
- [ ] Insert → PivotTable → range `'Q1 Revenue'!$A$1:$I$43` → New Worksheet → OK
- [ ] Drag **Sales Rep** to Rows, **Amount** to Values (make sure it's Sum not Count)
- [ ] Rename the new sheet to exactly `Pivot_SalesByRep`

### Step 3: Save as Macro-Enabled
- [ ] File → Save As → change type to **Excel Macro-Enabled Workbook (*.xlsm)**
- [ ] Keep filename as `Sample_Quarterly_ReportV2`
- [ ] Save in same folder, overwrite if prompted
- [ ] If prompted about compatibility, click Yes

### Step 4: Import VBA Modules
- [ ] Press Alt+F11 to open VBA Editor
- [ ] File → Import File → navigate to `RecTrial\UniversalToolkit\vba\`
- [ ] Import ALL `modUTL_*.bas` files (23 files)
- [ ] Also import from `RecTrial\UniversalToolkit\vba\NewTools\` subfolder (4 files)
- [ ] File → Import File → navigate to `RecTrial\VBAToImport\`
- [ ] Import **modDirector.bas**
- [ ] Press Alt+Q to close VBA Editor
- [ ] Press Ctrl+S to save the .xlsm

### Step 5: Verify Setup
- [ ] Click on the **Q1 Revenue** sheet tab
- [ ] Select cell A1
- [ ] Excel is maximized, zoom 100%
- [ ] Ignore any "Compile error: Variable not defined" if Debug → Compile shows it — this is expected (Video 1/2 clips reference demo-only modules not present in sample file)
- [ ] Do NOT attempt to compile — just run directly

### Step 6: Computer Lockdown
- [ ] Close everything except Excel and OBS
- [ ] Notifications OFF (Focus Assist → Alarms Only)
- [ ] Desktop icons hidden
- [ ] Taskbar auto-hidden
- [ ] Display: 1920x1080, 100% scaling

### Step 7: OBS Setup
- [ ] Recording Path: `RecTrial\Recordings\Video3\`
- [ ] 1920x1080, 30 FPS, MP4
- [ ] Desktop Audio: ENABLED
- [ ] Mic: DISABLED

---

## RECORDING

1. [ ] Make sure you're on the **Q1 Revenue** sheet with cell A1 selected
2. [ ] Start OBS recording
3. [ ] Wait 3 seconds
4. [ ] In Excel: Alt+F8 → `RunVideo3` → Run
5. [ ] If it warns about demo file, click **Yes** to continue
6. [ ] DO NOT TOUCH ANYTHING for ~9-10 minutes
7. [ ] If you see any dialog during recording, that's a bug — note the time and clip
8. [ ] When "Video 3 recording complete!" appears → click OK → Stop OBS

---

## WHAT HAPPENS DURING RECORDING (Clip by Clip)

### Clip 27 — Opening (~50 sec)
**Audio:** V3_S0_Opening.mp3
**Status bar:** (no message during this clip)
**What you should see:**
- Starts on Q1 Revenue sheet at cell A1
- Holds on header row for 4 seconds
- Slowly scrolls down showing messy data (mixed dates, oversized text on rows 2 and 6, blank rows)
- Quick tour: Q1 Expenses → Budget Summary → Contact List → back to Q1 Revenue
**Watch for:**
- Does it start on Q1 Revenue (NOT Cover, NOT Q1 Revenue v2)?
- Does the pre-flight dialog appear? It should NOT — pre-flight is now skipped for Video 3.

### Clip 28 — Data Sanitizer (~60 sec)
**Audio:** V3_C1A_DataSanitizer.mp3
**Status bar:**
1. `[Director] Running: Data Sanitizer - Preview`
2. `[Director] Running: Data Sanitizer - Full Clean`
**What you should see:**
- New sheet "UTL_Sanitizer_Preview" appears showing a list of issues found (red text for text-stored numbers)
- Data gets cleaned on Q1 Revenue
**Watch for:**
- NO "Run full sanitizer?" Yes/No dialog (was v1 bug — fixed)
- NO "Sanitization complete!" completion dialog (was v1 bug — fixed)

### Clip 29 — Highlights (~35 sec)
**Audio:** V3_C1B_Highlights.mp3
**Status bar:**
1. `[Director] Running: Highlight by Threshold (> $100,000)`
2. `[Director] Running: Highlight Duplicate Values`
**What you should see:**
- Amount column cells > $100,000 turn green
- Duplicate Amount values turn orange
- Highlights cleared at end
**Watch for:**
- NO range picker dialog (was v1 bug — fixed with DirectorHighlightThreshold silent wrapper)
- NO threshold value InputBox
- NO direction choice InputBox
- NO "Clear Highlights" dialog

### Clip 30 — Comments (~40 sec)
**Audio:** V3_C1C_Comments.mp3
**Status bar:** `[Director] Running: Extract All Comments`
**What you should see:**
- New sheet "Comment Inventory" appears with all 5 comments listed
- Columns: # | Sheet | Cell | Cell Value | Comment Author | Comment Text
**Watch for:**
- NO count MsgBox dialog
- NO "X comment(s) extracted" completion dialog

### Clip 31 — Tab Organizer (~50 sec)
**Audio:** V3_C2A_TabOrganizer.mp3
**Status bar:**
1. `[Director] Running: Color Tabs by Keyword (Revenue = Blue)`
2. `[Director] Running: Reorder Tabs Alphabetically`
**What you should see:**
- Tabs containing "Revenue" turn blue (Q1 Revenue, Q1 Revenue v2)
- Sheet tabs rearrange alphabetically
**Watch for:**
- NO keyword InputBox dialog
- NO color choice dialog
- Tab colors actually change visibly

### Clip 32 — Column Ops (~50 sec)
**Audio:** V3_C2B_ColumnOps.mp3
**Status bar:**
1. `[Director] Running: Split Column (Full Name -> First + Last)`
2. `[Director] Running: Combine Columns (City + State)`
**What you should see:**
- Switches to Contact List sheet
- Full Name column splits into two new columns (First, Last)
- Then combines City + State into one column with comma separator
**Watch for:**
- NO range picker dialog
- NO delimiter choice dialog
- NO separator choice dialog
- New columns actually appear (was v1 bug where nothing happened)

### Clip 33 — Sheet Tools (~50 sec)
**Audio:** V3_C2C_SheetTools.mp3
**Status bar:**
1. `[Director] Running: Create Sheet Index with Links`
2. `[Director] Running: Template Cloner (Q1 Expenses x 2)`
**What you should see:**
- New sheet appears with clickable hyperlinks to every other sheet
- Then 2 new copies of Q1 Expenses appear as new tabs
**Watch for:**
- NO "Which sheet to clone?" input dialog (was v1 bug — fixed)
- NO "How many copies?" input dialog (was v1 bug — fixed)
- Must see 2 copies created (0 copies = FAIL)

### Clip 34 — Compare Sheets (~50 sec)
**Audio:** V3_C3A_Compare.mp3
**Status bar:** `[Director] Running: Compare Sheets (Q1 Revenue vs Q1 Revenue v2)`
**What you should see:**
- New sheet "UTL_CompareReport" appears
- Shows differences between the two revenue sheets
- Differences highlighted/marked
**Watch for:**
- NO sheet-number selection dialog (was v1 bug — fixed)
- Comparison actually runs (was v1 bug where nothing happened)

### Clip 35 — Consolidate (~40 sec)
**Audio:** V3_C3B_Consolidate.mp3
**Status bar:** `[Director] Running: Consolidate Sheets (Q1 Revenue + Q1 Revenue v2)`
**What you should see:**
- New sheet "UTL_Consolidated" appears with combined data from both revenue sheets
- "Source Sheet" column shows which sheet each row came from
- Header row styled blue with white text
**Watch for:**
- NO sheet selection dialog (was v1 bug — fixed)
- Consolidation actually runs

### Clip 36 — Pivot Tools + Budget View (~60 sec)
**Audio:** V3_C3C_PivotTools.mp3 then V3_C3D_LookupValidation.mp3
**Status bar:** `[Director] Running: List All Pivot Tables`
**What you should see:**
- Pivot inventory sheet shows 2 pivot tables: Pivot_SalesByRegion and Pivot_SalesByRep (key v2 improvement — must NOT be empty)
- Screen navigates to Budget Summary sheet
- Budget Summary shows blue headers, currency formatting, color-coded status (green/orange/red)
- Scrolls down to show dropdown source list
**Watch for:**
- Pivot inventory must show 2 pivots (empty = FAIL, means pivots weren't added in Step 2)
- NO InputBox menus

### Clip 37 — Universal Command Center (~50 sec)
**Audio:** V3_C4_CommandCenter.mp3
**Status bar:** `[Director] Running: Universal Tool Inventory`
**What you should see:**
- New sheet "UTL_ToolInventory" appears listing all universal toolkit tools
- Organized by category
- Styled header row
**Watch for:**
- NO InputBox menu popup (was v1 bug where it often didn't trigger at all — now replaced with a sheet output)
- Inventory sheet actually appears

### Clip 38 — Closing (~45 sec)
**Audio:** V3_Closing.mp3
**What you should see:**
- Navigates back to Q1 Revenue sheet
- Holds static while closing narration plays
- "Video 3 recording complete!" MsgBox appears at end
**Watch for:**
- Ends on Q1 Revenue (NOT Cover, NOT any output sheet)
- Full audio plays without clipping

---

## AFTER RECORDING

- [ ] Review the recording in `RecTrial\Recordings\Video3\`
- [ ] Check audio plays through speakers
- [ ] Note any clips where dialogs appeared (should be zero)
- [ ] Note any timing issues (audio vs screen action mismatch)
- [ ] Send recording + `VIDEO_3_GEMINI_REVIEW.md` to Gemini for automated review

---

## IF SOMETHING GOES WRONG

- **Ctrl+Break** to force-stop the Director
- **Escape** between clips to abort (Director checks between clips)
- **If ANY dialog appears during recording** — that's a bug, note the clip and time
- To re-test one clip: Alt+F11 → Ctrl+G → type `TestClip 28` (or whatever number) → Enter
- To start fresh: close file WITHOUT saving, recopy clean .xlsx from backup folder, redo Steps 1-5

---

## WHAT CHANGED FROM V1

If you tested this video before, here's what's different now:

| Area | v1 (old) | v2 (now) |
|------|----------|----------|
| Dialogs during recording | Many (Sanitizer, Highlights, Clear, Tab, Column, Sheet Tools, Compare, Consolidate) | ZERO |
| Pre-flight dialog | Appeared at 0:21 | Skipped for Video 3 |
| Highlights | Range picker blocked | Auto-uses pre-selection |
| Tab Organizer | Failed to execute | Runs via DirectorColorTabsByKeyword |
| Column Ops | Failed to execute | Runs via DirectorSplitColumn / DirectorCombineColumns |
| Sheet Tools | InputBox dialogs | Runs via DirectorTemplateCloner with hardcoded params |
| Compare | Failed to execute | Runs via DirectorCompareSheets with sheet names |
| Consolidate | Failed to execute | Runs via DirectorConsolidateSheets with array |
| Command Center | InputBox menu (often absent) | Creates visible UTL_ToolInventory sheet |
| Pivot Tables | 0 pivots found (looked broken) | 2 real pivots found (Pivot_SalesByRegion, Pivot_SalesByRep) |
| Status bar messages | Partially working | Every tool shows its name during execution |
| Video runtime | 7:16 (ended early) | ~9-10 minutes full |

---

*Updated: 2026-04-17*
*Version 2 — Post Path A silent wrapper refactor*
