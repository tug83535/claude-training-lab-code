# Video Script Updates — Instructions for Claude

## Context
The video scripts (Video 1, 2, and 3) were written when the project had 62 Command Center actions, 34 VBA modules, and ~100 universal tools. Since then, new features were added. The scripts need to be updated before recording.

---

## Files to Review Before Starting

Give your Claude these files so it has the full picture:

1. **`videodraft/COMPILED_VIDEO_PACKAGE.md`** — Contains all 3 video scripts (this is the file to edit)
2. **`videodraft/VIDEO_DEMO_PLAN.md`** — Strategy and planning doc (also needs number updates)
3. **`vba/modWhatIf_v2.1.bas`** — What-If Scenario module (new feature to demo)
4. **`vba/modTimeSaved_v2.1.bas`** — Time Saved Calculator (new feature to demo)
5. **`vba/modSplashScreen_v2.1.bas`** — Branded splash screen (new feature to demo)
6. **`vba/modExecBrief_v2.1.bas`** — Executive Brief generator (new feature to demo)
7. **`UniversalToolsForAllFiles/vba/`** — All 23 universal toolkit modules (Video 3 restructure)

---

## PART 1 — Global Number Updates (All Scripts)

Find and replace these everywhere across both video docs:

| Old | New |
|-----|-----|
| 62 actions | **65 actions** |
| 62 automated actions | **65 automated actions** |
| 62 Command Center actions | **65 Command Center actions** |
| 34 VBA modules | **39 VBA modules** |
| 13 VBA modules (universal) | **23 VBA modules** |
| ~78 tools (universal) | **~140+ tools** |
| ~100 universal tools | **~140+ tools** |
| 14 Python scripts (demo) | **14 Python scripts** (unchanged) |
| 22 Python scripts (universal) | **22 Python scripts** (unchanged) |

Also update the closing stat line from:
> "62 actions. 0 cost. 10 minutes vs 10 hours."

To:
> "65 actions. 140+ universal tools. Zero cost. 10 minutes vs 10 hours."

---

## PART 2 — Video 1 Script Changes ("What's Possible")

### Minor updates only:
- Change "62" to "65" in the opening hook
- Everything else stays the same — this video is still just the highlight reel

---

## PART 3 — Video 2 Script Changes ("Full Demo Walkthrough")

### 3A. Add Splash Screen to the Opening
When the file opens at the start of the video, the branded splash screen (`modSplashScreen`) fires automatically. Mention it briefly:
> "When you open the file, you're greeted with a branded welcome screen. This sets the tone — this isn't just a spreadsheet."

This costs 10 seconds and makes the first impression more polished.

### 3B. Add What-If Scenario Demo (NEW — Tier 1 Wow Moment)
This is the biggest addition. Add a new section in Chapter 5 (Enterprise Features) — or give it its own chapter. This is a CFO/CEO favorite.

**What it does:**
- Actions 63-65 in the Command Center
- Action 63: Run What-If Scenario Demo — picks from 7 presets (Revenue +15%, Revenue -15%, Revenue +10%, Revenue -10%, AWS Costs +25%, Headcount +20%, All Expenses -10%) plus Best Case, Worst Case, Custom, and Restore Baseline
- Action 64: Load Pre-Built Scenario — quick-pick a preset
- Action 65: Restore to Original Values — puts everything back

**How to demo it (suggested script):**
> "Here's something the CFO asked for — What-If scenarios. Watch this.
>
> I'll run Action 63. Let's pick 'Revenue drops 15%.'
>
> [run it — numbers change on the Assumptions sheet, impact report appears]
>
> Every driver on the Assumptions sheet just updated. And here's the impact report showing exactly what changed and by how much.
>
> Now let's restore it. Action 65 — Restore to Original Values.
>
> [run it — everything goes back]
>
> Back to baseline. That whole analysis took 10 seconds."

**Time needed:** 90 seconds. This should be one of the biggest moments in the video.

### 3C. Add Time Saved Calculator as the Closing Moment
Replace or supplement the current closing with the Time Saved Calculator (`modTimeSaved`). This is more powerful than you just saying a number — the tool calculates it and builds a report.

**How to demo it (suggested script):**
> "One last thing. Let me show you the actual time savings.
>
> [run Time Saved Calculator]
>
> This scans all 65 actions and calculates how long each one would take manually versus automated. Here's the result:
>
> [show the report — manual hours per month, automated hours per month, annual savings]
>
> That number speaks for itself."

**Time needed:** 45-60 seconds. End the video on this.

### 3D. Updated Chapter 5 Structure
**Current:** Executive Mode → Version Control → Scenario Management → Sensitivity Analysis
**New:** Executive Mode → Version Control → **What-If Scenario Demo** → Sensitivity Analysis → **Time Saved Calculator**

Drop "Scenario Management" (save/load/compare) — it overlaps with What-If and is less visual. The What-If presets are more demo-friendly for the audience.

### 3E. Optional: Executive Brief (if time allows)
`modExecBrief` generates a one-click executive brief scanning Revenue, Reconciliation, Assumptions, Products, and Workbook Health. Could be a nice 30-second addition to Chapter 4 (Reporting & Visuals) after PDF Export:
> "And if you need a quick executive summary — one click, and it scans the entire workbook and builds a brief with color-coded status indicators."

---

## PART 4 — Video 3 Script Changes ("Universal Tools")

### 4A. Complete Restructure Needed
The current Video 3 script covers "Sheet Tools" and "Data Tools" generically. But there are now 23 modules with 140+ tools. The script needs to show the best of the new stuff.

### 4B. New Chapter Structure (suggested)

**Chapter 1: Data Cleanup (2.5 min)**
- modUTL_DataSanitizer — Fix text-stored numbers, floating-point noise
- modUTL_Highlights — Threshold highlighting, top/bottom N, duplicate detection
- modUTL_Comments — Extract/clear/convert comments and notes
- Demo pattern: messy data → one click → clean data

**Chapter 2: Sheet & Column Tools (2.5 min)**
- modUTL_TabOrganizer — Sort/color/group/reorder tabs in bulk
- modUTL_ColumnOps — Column insert/delete/move/split/merge
- modUTL_SheetTools — Sheet index with hyperlinks, template cloner
- Demo pattern: disorganized workbook → one click → organized workbook

**Chapter 3: Analysis & Building Tools (2.5 min)**
- modUTL_Compare — Sheet comparison with color-coded diff report
- modUTL_Consolidate — Multi-sheet consolidation with source tracking
- modUTL_PivotTools — PivotTable create/refresh/style
- modUTL_LookupBuilder — VLOOKUP/INDEX-MATCH formula builder
- modUTL_ValidationBuilder — Data validation builder
- Demo pattern: manual task → one click → professional result

**Chapter 4: The Universal Command Center (1 min)**
- modUTL_CommandCenter — Show them the master menu for all ~140+ tools
- 13 categories, keyword search
- "You don't have to memorize anything. Just open the menu and search."

**Closing (30 sec)**
- Where to find the library on SharePoint
- Point to the training guides and CoPilot prompts

### 4C. Key Modules to Reference
Here are all 23 universal toolkit modules for accuracy:

| # | Module | Tools | What It Does |
|---|--------|-------|-------------|
| 1 | modUTL_Core | 9 | Shared utilities (StyleHeader, TurboOn/Off, etc.) |
| 2 | modUTL_Audit | ~10 | External links, hidden sheets, error scan |
| 3 | modUTL_Branding | 2 | iPipeline brand colors + theme |
| 4 | modUTL_DataQuality | ~8 | Data scans, letter grade |
| 5 | modUTL_DataSanitizer | 4 | Fix text-numbers, floating-point, sanitize |
| 6 | modUTL_Export | ~5 | PDF/CSV export |
| 7 | modUTL_Formatting | ~10 | AutoFit, borders, number formats |
| 8 | modUTL_Navigation | ~5 | TOC, sheet index, hyperlinks |
| 9 | modUTL_Search | ~3 | Cross-sheet search |
| 10 | modUTL_SheetTools | 3 | Sheet index, template cloner, customer IDs |
| 11 | modUTL_Utilities | ~12 | General-purpose tools |
| 12 | modUTL_ProgressBar | 1 | ASCII progress bar (status bar) |
| 13 | modUTL_SplashScreen | 1 | Standalone MsgBox splash |
| 14 | modUTL_ExecBrief | 1 | Universal workbook executive brief |
| 15 | modUTL_ColumnOps | 7 | Column insert/delete/move/split/merge/fill/swap |
| 16 | modUTL_Compare | 1 | Sheet comparison with diff report |
| 17 | modUTL_Consolidate | 1 | Multi-sheet consolidation |
| 18 | modUTL_Highlights | 3 | Threshold, top/bottom N, duplicates |
| 19 | modUTL_PivotTools | 4 | PivotTable create/refresh/style/drill |
| 20 | modUTL_TabOrganizer | 6 | Sort/color/group/reorder/rename tabs |
| 21 | modUTL_Comments | 3 | Extract/clear/convert comments |
| 22 | modUTL_ValidationBuilder | 5 | Data validation builder |
| 23 | modUTL_LookupBuilder | 2 | VLOOKUP/INDEX-MATCH builder |
| + | modUTL_WhatIf | 1 | Universal What-If scenarios |
| + | modUTL_CommandCenter | 1 | Master menu for all tools |

---

## PART 5 — Integration Test Count Check

The scripts mention "18/18 PASS" for the integration test. Verify this is still accurate after adding 5 new modules (modTimeSaved, modSplashScreen, modProgressBar, modWhatIf, modExecBrief). If the test count changed, update the script.

---

## Summary of All Changes

| Video | What Changes |
|-------|-------------|
| All | 62 → 65 actions, ~100 → ~140+ tools, 34 → 39 modules |
| Video 1 | Number updates only |
| Video 2 | Add splash screen opening, What-If demo (Tier 1), Time Saved closing, optional Exec Brief |
| Video 2 | Restructure Chapter 5: drop Scenario Mgmt, add What-If + Time Saved |
| Video 3 | Full chapter restructure around the 23 modules |
| Video 3 | Add Universal Command Center as its own chapter |
| Closing line | "65 actions. 140+ universal tools. Zero cost. 10 minutes vs 10 hours." |
