# VIDEO 3 — Step-by-Step Setup & Recording Guide

**What This Is:** Complete step-by-step instructions from zero to finished Video 3 recording. Follow every step in order.

---

## STEP 0: Confirm You Have Everything

Before starting, verify these files exist:

- [ ] `RecTrial\SampleFile\SampleFileV2\Sample_Quarterly_ReportV2.xlsx` — the sample file from Claude online
- [ ] `RecTrial\VBAToImport\modDirector.bas` — the Director macro
- [ ] `RecTrial\UniversalToolkit\vba\` — folder with ~23 modUTL_*.bas files
- [ ] `RecTrial\AudioClips\Video3\` — folder with 13 MP3 audio clips
- [ ] OBS Studio installed and configured (same settings as Video 1 & 2)

---

## STEP 1: Open the Sample File

1. Open File Explorer
2. Navigate to `C:\Users\connor.atlee\RecTrial\SampleFile\SampleFileV2\`
3. Double-click `Sample_Quarterly_ReportV2.xlsx`
4. If you see a yellow "Enable Content" bar, click **Enable Content**

---

## STEP 2: Save As Macro-Enabled Workbook

The file is .xlsx (no macros). You need to save it as .xlsm so VBA modules can be imported.

1. Click **File** (top left)
2. Click **Save As**
3. Click **Browse**
4. In the "Save as type" dropdown at the bottom, change it to **"Excel Macro-Enabled Workbook (*.xlsm)"**
5. Keep the same filename and location
6. Click **Save**
7. If Excel asks about compatibility, click **Yes**

---

## STEP 3: Import Universal Toolkit VBA Modules

You need to import ~23 modUTL_*.bas files. This is tedious but only done once.

1. Press **Alt+F11** to open the VBA Editor
2. Click **File** → **Import File...**
3. Navigate to `C:\Users\connor.atlee\RecTrial\UniversalToolkit\vba\`
4. Select the first .bas file (e.g., modUTL_Audit.bas)
5. Click **Open**
6. Repeat steps 2-5 for EVERY .bas file in that folder:
   - modUTL_Audit.bas
   - modUTL_Branding.bas
   - modUTL_ColumnOps.bas
   - modUTL_CommandCenter.bas
   - modUTL_Comments.bas
   - modUTL_Compare.bas
   - modUTL_Consolidate.bas
   - modUTL_Core.bas
   - modUTL_DataCleaning.bas
   - modUTL_DataSanitizer.bas
   - modUTL_ExecBrief.bas
   - modUTL_Finance.bas
   - modUTL_Formatting.bas
   - modUTL_Highlights.bas
   - modUTL_LookupBuilder.bas
   - modUTL_PivotTools.bas
   - modUTL_ProgressBar.bas
   - modUTL_SheetTools.bas
   - modUTL_SplashScreen.bas
   - modUTL_TabOrganizer.bas
   - modUTL_ValidationBuilder.bas
   - modUTL_WhatIf.bas
   - modUTL_WorkbookMgmt.bas
7. Also import the NewTools subfolder files if they exist:
   - Navigate to `RecTrial\UniversalToolkit\vba\NewTools\`
   - Import: modUTL_AuditPlus.bas, modUTL_DataCleaningPlus.bas, modUTL_DuplicateDetection.bas, modUTL_NumberFormat.bas

**Tip:** You can select multiple files at once in the Import dialog by holding Ctrl and clicking each one.

---

## STEP 4: Import modDirector

1. Still in the VBA Editor (Alt+F11)
2. Click **File** → **Import File...**
3. Navigate to `C:\Users\connor.atlee\RecTrial\VBAToImport\`
4. Select **modDirector.bas**
5. Click **Open**
6. You should see **modDirector** in the Modules list on the left

---

## STEP 5: Verify the Audio Path

1. In the VBA Editor, double-click **modDirector** in the left panel
2. Near the top (~line 79), find this line:

```
Private Const AUDIO_BASE_PATH As String = "C:\Users\connor.atlee\RecTrial\AudioClips\"
```

3. **If this path is correct, do nothing.** If you moved the AudioClips folder, update it.
4. The path MUST end with a backslash (\)

---

## STEP 6: Compile and Save

1. In the VBA Editor, click **Debug** → **Compile VBAProject**
2. If there are no errors, you're good
3. If there's an error, note which module/line and let me know
4. Press **Ctrl+S** to save
5. Press **Alt+Q** to close the VBA Editor

---

## STEP 7: Verify the File is on Q1 Revenue

1. Back in Excel, click on the **Q1 Revenue** sheet tab at the bottom
2. Select cell **A1**
3. Make sure the data looks like a sales pipeline (Region, Sales Rep, Product, etc.)
4. You should see messy data — mixed date formats, blank rows, no number formatting

---

## STEP 8: Computer Lockdown

1. Close everything except Excel and OBS
2. Turn off all notifications (Focus Assist → Alarms Only)
3. Clean desktop (hide icons)
4. Auto-hide taskbar
5. Plug in laptop
6. Phone on silent

---

## STEP 9: Set Up OBS

1. Open OBS Studio
2. Verify settings are same as Video 1 & 2:
   - Recording Path: `RecTrial\Recordings\Video3\`
   - 1920x1080, 30 FPS, MP4
   - Desktop Audio: ENABLED
   - Mic: DISABLED
3. Verify Display Capture source is active

---

## STEP 10: Quick Test

1. In Excel, press **Alt+F8**
2. Type `QuickTest`
3. Click **Run**
4. Verify: audio plays, screen scrolls, pre-flight passes
5. If audio doesn't play, check AUDIO_BASE_PATH

---

## STEP 11: Record Video 3

1. Make sure you are in the **sample file** (not the demo file)
2. Make sure you are on the **Q1 Revenue** sheet
3. Excel is maximized, zoom 100%
4. **Start OBS recording**
5. **Wait 3 seconds**
6. Press **Alt+F8** → Select **RunVideo3** → Click **Run**
7. The macro will warn you if it detects you're on the demo file — click **Yes** to continue or **No** to stop
8. **DO NOT TOUCH ANYTHING** for ~10 minutes
9. When the **"Video 3 recording complete!"** message appears:
   - Click **OK**
   - **Stop OBS recording**

---

## STEP 12: Review

1. Find the recording in `RecTrial\Recordings\Video3\`
2. Play it back
3. Check: audio plays, tools produce output, no error dialogs
4. Note any issues for the feedback document

---

## If Something Goes Wrong

- **Ctrl+Break** to force-stop the Director
- Run `CleanupAllOutputSheets` if needed (but this is the sample file, not the demo)
- To test a single clip: Alt+F11 → Ctrl+G → type `TestClip 28` → Enter
- If a macro errors, click **End** (not Debug)

---

## What the Director Does During Video 3

| Time | Clip | What Happens |
|------|------|-------------|
| 0:00 | 27 | Opens on Q1 Revenue, scrolls through messy data |
| ~0:45 | 28 | Runs DataSanitizer preview then full clean |
| ~1:45 | 29 | Highlights values >$5K, highlights duplicates |
| ~2:20 | 30 | Counts comments (5), extracts to new sheet |
| ~3:00 | 31 | Colors tabs by keyword, reorders tabs |
| ~3:50 | 32 | Splits Full Name column, combines columns |
| ~4:40 | 33 | Creates sheet index with links, clones Q1 Expenses |
| ~5:30 | 34 | Compares Q1 Revenue vs Q1 Revenue v2 (8 diffs) |
| ~6:20 | 35 | Consolidates revenue sheets into one |
| ~7:00 | 36 | Lists pivots, builds VLOOKUP, creates dropdown |
| ~8:00 | 37 | Opens Universal Command Center |
| ~8:50 | 38 | Closing — holds on first sheet |

*Created: 2026-03-31*
