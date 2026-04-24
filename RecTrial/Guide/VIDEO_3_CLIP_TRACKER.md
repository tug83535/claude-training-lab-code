# VIDEO 3 — Clip-by-Clip Tracker & Review Sheet

**How to use:** During or after your Video 3 test run, mark each clip as PASS or FAIL and add any comments. Give this to Claude (online or Code) to fix issues.

---

## CLIP 27 — Opening (~45 sec)
**Audio:** V3_S0_Opening.mp3
**What should happen:** Director navigates to Q1 Revenue sheet, holds still, then slowly scrolls down through the messy data while narration plays.
**What the viewer sees:** A real quarterly report with sales data — regions, reps, products, customers, amounts. Mixed date formats, blank rows, no formatting. The "before" state.

**Script excerpt:** "This is a real quarterly report... multiple people have touched it... dates are inconsistent, numbers stored as text, blank rows scattered through the data..."

| Check | Status | Comments |
|-------|--------|----------|
| Opens on Q1 Revenue (not Cover)? | ⬜ PASS / ⬜ FAIL | |
| Audio plays? | ⬜ PASS / ⬜ FAIL | |
| Scrolls through data smoothly? | ⬜ PASS / ⬜ FAIL | |
| Scroll stays within data (no blank area)? | ⬜ PASS / ⬜ FAIL | |
| Audio finishes without clipping? | ⬜ PASS / ⬜ FAIL | |

---

## CLIP 28 — Data Sanitizer (~60 sec)
**Audio:** V3_C1A_DataSanitizer.mp3
**What should happen:**
1. Runs PreviewSanitizeChanges — shows a report of what WOULD change (text-stored numbers, blank rows)
2. Then runs RunFullSanitize — actually cleans the data
**What the viewer sees:** Preview report appears showing issues found, then the data gets cleaned.

**Script excerpt:** "First, let's see what's wrong with this file... the sanitizer scans every sheet and shows you exactly what it would fix before changing anything... then one click to clean it all."

| Check | Status | Comments |
|-------|--------|----------|
| Preview runs without error? | ⬜ PASS / ⬜ FAIL | |
| Preview shows findings (text nums, blank rows)? | ⬜ PASS / ⬜ FAIL | |
| Full sanitize runs without error? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes pop up and block? | ⬜ YES / ⬜ NO | |

---

## CLIP 29 — Highlights (~35 sec)
**Audio:** V3_C1B_Highlights.mp3
**What should happen:**
1. HighlightByThreshold — highlights cells with amounts over $5,000 (SendKeys pre-stages "5000")
2. HighlightDuplicateValues — highlights duplicate values
3. ClearHighlights — removes all highlights after
**What the viewer sees:** Cells light up with color highlighting, showing which values exceed the threshold and which are duplicates.

**Script excerpt:** "Want to quickly spot values above a threshold? One click highlights everything over five thousand... duplicates? One click finds those too."

| Check | Status | Comments |
|-------|--------|----------|
| Threshold highlight runs? | ⬜ PASS / ⬜ FAIL | |
| Cells actually highlight with color? | ⬜ PASS / ⬜ FAIL | |
| Duplicate highlight runs? | ⬜ PASS / ⬜ FAIL | |
| Highlights cleared after? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 30 — Comments (~40 sec)
**Audio:** V3_C1C_Comments.mp3
**What should happen:**
1. CountComments — shows MsgBox with count (5 comments) — auto-dismissed by SendKeys
2. ExtractAllComments — creates a new sheet with all comments listed
**What the viewer sees:** Comment count appears briefly, then a new sheet shows all 5 comments extracted with cell references.

**Script excerpt:** "Every comment in the workbook, extracted into one list... who wrote it, what cell, what they said."

| Check | Status | Comments |
|-------|--------|----------|
| Comment count runs? | ⬜ PASS / ⬜ FAIL | |
| Shows correct count (5)? | ⬜ PASS / ⬜ FAIL | |
| Extract creates new sheet? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 31 — Tab Organizer (~50 sec)
**Audio:** V3_C2A_TabOrganizer.mp3
**What should happen:**
1. ColorTabsByKeyword — colors sheet tabs based on keywords in their names
2. ReorderTabs — reorders tabs alphabetically or by category
**What the viewer sees:** Tab colors change at the bottom of the screen, then tabs rearrange.

**Script excerpt:** "Color-code your tabs by keyword... revenue tabs get one color, expense tabs another... then reorder them in one click."

| Check | Status | Comments |
|-------|--------|----------|
| Tab colors change? | ⬜ PASS / ⬜ FAIL | |
| Tab reorder happens? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 32 — Column Ops (~50 sec)
**Audio:** V3_C2B_ColumnOps.mp3
**What should happen:**
1. SplitColumn — splits a column (likely Full Name on Contact List into First/Last)
2. CombineColumns — combines columns back together
**What the viewer sees:** A column splits into two, then columns merge back.

**Script excerpt:** "Split a full name column into first and last... or combine city and state into one field."

| Check | Status | Comments |
|-------|--------|----------|
| Split runs without error? | ⬜ PASS / ⬜ FAIL | |
| Column actually splits? | ⬜ PASS / ⬜ FAIL | |
| Combine runs without error? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 33 — Sheet Tools (~50 sec)
**Audio:** V3_C2C_SheetTools.mp3
**What should happen:**
1. ListAllSheetsWithLinks — creates a sheet index with hyperlinks to every tab
2. TemplateCloner — clones "Q1 Expenses" sheet (SendKeys pre-stages the name + "2" copies)
**What the viewer sees:** New index sheet appears with clickable links, then 2 copies of Q1 Expenses appear.

**Script excerpt:** "Create a clickable table of contents for your workbook... clone any sheet as a template."

| Check | Status | Comments |
|-------|--------|----------|
| Sheet index created? | ⬜ PASS / ⬜ FAIL | |
| Links are clickable? | ⬜ PASS / ⬜ FAIL | |
| Template clone works? | ⬜ PASS / ⬜ FAIL | |
| 2 copies created? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 34 — Compare Sheets (~50 sec)
**Audio:** V3_C3A_Compare.mp3
**What should happen:**
1. CompareSheets — compares Q1 Revenue vs Q1 Revenue v2 cell by cell
2. Creates a diff report showing 8 differences + 1 extra row
**What the viewer sees:** New sheet appears with a color-coded comparison showing exactly what changed between the two versions.

**Script excerpt:** "Two versions of the same report... which cells changed? Compare them side by side, cell by cell."

| Check | Status | Comments |
|-------|--------|----------|
| Compare runs without error? | ⬜ PASS / ⬜ FAIL | |
| Diff report sheet created? | ⬜ PASS / ⬜ FAIL | |
| Shows actual differences? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 35 — Consolidate (~40 sec)
**Audio:** V3_C3B_Consolidate.mp3
**What should happen:**
1. ConsolidateSheets — stacks Q1 Revenue + Q1 Revenue v2 into one sheet with a source column
**What the viewer sees:** New consolidated sheet with all rows from both sheets, plus a column showing which sheet each row came from.

**Script excerpt:** "Pull data from multiple sheets into one... with source tracking so you know where every row came from."

| Check | Status | Comments |
|-------|--------|----------|
| Consolidate runs without error? | ⬜ PASS / ⬜ FAIL | |
| New sheet created? | ⬜ PASS / ⬜ FAIL | |
| Shows source column? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 36 — Pivot Tools + Lookup/Validation (~60 sec)
**Audio:** V3_C3C_PivotTools.mp3 then V3_C3D_LookupValidation.mp3
**What should happen:**
1. ListAllPivots — lists any existing pivot tables (may find none — that's OK)
2. BuildVLOOKUP — builds a VLOOKUP formula
3. CreateDropdownList — creates a data validation dropdown from the Budget Summary status list
**What the viewer sees:** Pivot inventory appears, then a VLOOKUP formula is built, then a dropdown appears in a cell.

**Script excerpt:** "Inventory your pivot tables... build a VLOOKUP without typing the formula... create a dropdown from a list."

| Check | Status | Comments |
|-------|--------|----------|
| ListAllPivots runs? | ⬜ PASS / ⬜ FAIL | |
| BuildVLOOKUP runs? | ⬜ PASS / ⬜ FAIL | |
| CreateDropdownList runs? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully (both clips)? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 37 — Universal Command Center (~50 sec)
**Audio:** V3_C4_CommandCenter.mp3
**What should happen:**
1. LaunchUTLCommandCenter — opens the Universal Toolkit Command Center form
**What the viewer sees:** A menu form appears listing all universal toolkit tools organized by category.

**Script excerpt:** "All of these tools — organized in one Command Center... browse by category, search, click Run."

| Check | Status | Comments |
|-------|--------|----------|
| Command Center opens? | ⬜ PASS / ⬜ FAIL | |
| Shows tool categories? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| Any dialog boxes block? | ⬜ YES / ⬜ NO | |

---

## CLIP 38 — Closing (~45 sec)
**Audio:** V3_Closing.mp3
**What should happen:** Director navigates to first sheet, holds still while closing narration plays.
**What the viewer sees:** Static view of the (now cleaned up) file while narration wraps up.

**Script excerpt:** "Everything you just saw works on any Excel file... all the tools and guides are available on SharePoint."

| Check | Status | Comments |
|-------|--------|----------|
| Navigates to a sheet (not blank)? | ⬜ PASS / ⬜ FAIL | |
| Audio plays fully? | ⬜ PASS / ⬜ FAIL | |
| "Video 3 complete" message appears? | ⬜ PASS / ⬜ FAIL | |

---

## OVERALL VIDEO 3 ASSESSMENT

| Category | Status | Notes |
|----------|--------|-------|
| All audio clips played? | ⬜ YES / ⬜ NO | |
| Any audio clipped early? | ⬜ YES / ⬜ NO | Which clips? |
| Any error dialogs appeared? | ⬜ YES / ⬜ NO | Which clips? |
| Any macros failed silently? | ⬜ YES / ⬜ NO | Which clips? |
| Scrolling looked natural? | ⬜ YES / ⬜ NO | |
| Total runtime acceptable? | ⬜ YES / ⬜ NO | How long? |
| Ready for final recording? | ⬜ YES / ⬜ NO | |

**Additional comments / issues:**

(Write anything else you noticed here)

---

*Created: 2026-03-31 | Video 3 Clip Tracker for iPipeline Finance Automation Demo*
