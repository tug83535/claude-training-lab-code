# MASTER RECORDING GUIDE — iPipeline Finance Automation Videos

**What This Is:** Your single, complete guide for recording all 3 demo videos. The Director macro automates the screen — it plays your audio, navigates sheets, runs macros, scrolls, and pauses. You press one button and watch.

**What You Are Recording:**
- **Video 1:** "What's Possible" — 4-5 minute overview for all 2,000+ employees (Clips 1-7)
- **Video 2:** "Full Demo Walkthrough" — 15-18 minute deep dive for Finance & Accounting (Clips 8-26)
- **Video 3:** "Universal Tools" — 8-10 minute demo for anyone who uses Excel (Clips 27-38)

**Estimated Total Time:** ~2 hours (setup + recording + review)

---

## Table of Contents

1. [Your Files — Where Everything Lives](#your-files--where-everything-lives)
2. [Phase 1: Excel File Setup](#phase-1-excel-file-setup)
3. [Phase 2: Import the Director Module](#phase-2-import-the-director-module)
4. [Phase 3: Set Up OBS Studio](#phase-3-set-up-obs-studio)
5. [Phase 4: Prepare Video 3 Files](#phase-4-prepare-video-3-files)
6. [Phase 5: Computer Lockdown](#phase-5-computer-lockdown)
7. [Phase 6: Pre-Flight and Quick Test](#phase-6-pre-flight-and-quick-test)
8. [Phase 7: Record](#phase-7-record)
9. [Phase 8: If Something Goes Wrong](#phase-8-if-something-goes-wrong)
10. [Phase 9: Title Cards](#phase-9-title-cards-after-recording)
11. [Reference: What the Director Does — Clip by Clip](#reference-what-the-director-does--clip-by-clip)
12. [Reference: Action Numbers and Macro Names](#reference-action-numbers-and-macro-names)
13. [Reference: Clip Master Index](#reference-clip-master-index)

---
---

# YOUR FILES — Where Everything Lives

Everything you need is in one folder on your machine:

```
C:\Users\connor.atlee\RecTrial\
│
├── AudioClips\              ← MP3 narration files (the Director reads from here)
│   ├── Video1\  (7 MP3s)
│   ├── Video2\  (21 MP3s)
│   └── Video3\  (13 MP3s)
│
├── DemoFile\                ← Your Excel demo file
│   └── ExcelDemoFile_adv.xlsm
│
├── SampleFile\              ← Sample file for Video 3 (if you have one)
│
├── DemoVBA\                 ← All 39 demo VBA modules + frmCommandCenter code
│
├── UniversalToolkit\        ← Universal toolkit VBA + Python (for Video 3)
│   ├── vba\   (27 modules)
│   └── python\ (22 scripts)
│
├── DemoPython\              ← Demo Python scripts + SQL
│
├── VBAToImport\             ← The 2 files you need to import for recording
│   ├── modDirector.bas
│   └── modWhatIf_v2.1.bas
│
├── Guides\                  ← 15 PDF training guides
├── VideoScripts\            ← Video narration scripts (markdown)
├── Guide\                   ← THIS guide
│
└── Recordings\              ← OBS saves recordings here
    ├── Video1\
    ├── Video2\
    └── Video3\
```

**Important:** The Excel file can live ANYWHERE on your machine. It does not have to be in RecTrial. The only path that matters is `AUDIO_BASE_PATH` inside modDirector — it must point to your AudioClips folder. The copy in VBAToImport is already set to `C:\Users\connor.atlee\RecTrial\AudioClips\`.

---
---

# PHASE 1: EXCEL FILE SETUP

**This must be completed BEFORE you do anything else. Do not skip any step.**

---

## Part 1: Open the Demo File and Enable Macros

1. Find your demo file: `ExcelDemoFile_adv.xlsm` (it is in `RecTrial\DemoFile\` or wherever you saved your final copy)
2. Double-click the file to open it in Excel
3. **Look at the very top of the Excel window.** If you see a yellow bar that says **"SECURITY WARNING — Macros have been disabled"** with an "Enable Content" button:
   - Click **Enable Content**
   - This allows the VBA macros to run. Without this, nothing will work.
4. If you do NOT see that yellow bar, macros are already enabled. Move on.

### Make macros always work for this file:

5. In Excel, click **File** (top left corner)
6. Click **Options** (bottom of the left sidebar)
7. Click **Trust Center** (bottom of the left sidebar in Options)
8. Click **Trust Center Settings...** (the button on the right side)
9. Click **Trusted Locations** (left sidebar)
10. Click **Add new location...**
11. Click **Browse...** and navigate to the folder where your demo .xlsm file is saved
12. Select that folder and click **OK**
13. Check the box that says **"Subfolders of this location are also trusted"**
14. Click **OK** three times to close all dialogs
15. **Also do this:** While still in Trust Center Settings, click **Macro Settings** in the left sidebar
16. Make sure **"Enable all macros"** is selected (the last radio button)
17. Check the box at the bottom: **"Trust access to the VBA project object model"**
18. Click **OK** twice to close

**Why this matters:** If macros are not enabled, the Command Center will not open. No macros will run. Your recording will be useless.

---

## Part 2: Clean Up Leftover Output Sheets

Every time you run a macro, it creates an output sheet. If any leftover sheets exist from a previous session, they will show up on camera and confuse the viewer. You need a perfectly clean starting state.

> **Shortcut:** You can run the `CleanupAllOutputSheets` macro to do all of this automatically: Press **Alt+F8**, select **CleanupAllOutputSheets**, click **Run**. If you use the shortcut, skip to Part 3.

### If doing it manually:

19. Look at the sheet tabs at the very bottom of the Excel window
20. Right-click on each of the following sheets (if they exist) and click **Delete**, then click **Delete** again on the confirmation popup. If a sheet does not exist, skip it:
    - **Data Quality Report** — delete it
    - **Variance Analysis** — delete it
    - **Variance Commentary** — delete it
    - **Executive Dashboard** — delete it
    - **YoY Variance Analysis** — delete it
    - **Sensitivity Analysis** — delete it
    - **Time Saved Analysis** — delete it
    - **Executive Brief** — delete it
    - **Integration Test Report** — delete it
    - **What-If Impact** — delete it
    - **Charts & Visuals** — only delete if it looks like charts from a previous macro run. If it is a permanent sheet with pre-built content, leave it alone.
    - Any sheet starting with **"VER_"** — these are version control snapshots. Delete all of them.
    - Any sheet starting with **"BKP_"** — these are backup copies. Delete all of them.

21. Check for hidden sheets:
    - Right-click on ANY sheet tab at the bottom
    - Click **Unhide...**
    - If you see **"WhatIf_Baseline"**, select it, click **OK**, then right-click it and **Delete** it
    - Right-click any tab again → **Unhide...** again
    - If you see **"VBA_AuditLog"**, select it, click **OK** to unhide it
    - Click on the VBA_AuditLog sheet tab
    - Select all the data below the header row (click on row 2, then press Ctrl+Shift+End)
    - Press the **Delete** key to clear all the data (leave the headers in row 1)
    - Right-click the VBA_AuditLog tab → click **Hide**

22. Check the **Checks** sheet:
    - Click on the **Checks** tab
    - If there is any data below the headers, select it all (click row 2, Ctrl+Shift+End) and press Delete
    - The Checks sheet should be empty except for headers

---

## Part 3: Verify the Core Sheets Exist and Have Data

Your demo file should have these sheets. Click on each one to verify it exists and has data:

23. **Report-->** — This is your landing page. It should have iPipeline branding (navy headers, blue accents). This is where every video starts.
24. **P&L - Monthly Trend** — Main financial data. Should have revenue and expense line items across multiple months. Scroll right to verify you see month columns.
25. **Product Line Summary** — Should show revenue by product (iGO, Affirm, InsureSight, DocFast).
26. **Functional P&L - Monthly Trend** — Functional expense breakdown view.
27. **Functional P&L Summary** tabs (e.g., "Mar 25" or individual month tabs) — At least one should exist.
28. **Assumptions** — Should have driver values (growth rates, allocation percentages). Critical for the What-If demo.
29. **General Ledger** (may be called "CrossfireHiddenWorksheet") — Should have raw transaction data. If you don't see "General Ledger" but do see "CrossfireHiddenWorksheet" — that IS your GL sheet. It's fine.
30. **Checks** — Should exist but be empty (you cleared it in step 22).

**If any of these sheets are missing:** STOP. You cannot record until the demo file has all required sheets with data.

---

## Part 4: Test That Key Macros Actually Work

Before recording, confirm that the macros you will demo actually run without errors.

31. Press **Ctrl+Shift+M** on your keyboard
    - **What should happen:** The Command Center form pops up — a window with a list of all 65 actions organized by category, with a search box at the top
    - **If nothing happens:** Press **Alt+F8**, type `LaunchCommandCenter`, click **Run**. If that also fails, the VBA modules are not imported correctly. STOP and fix this.
32. In the Command Center, type **7** in the action number box and click **Run**
    - **What should happen:** A new sheet called "Data Quality Report" appears with a letter grade at the top
    - **If it errors:** STOP. Fix the error before proceeding.
33. Delete the "Data Quality Report" sheet (right-click tab → Delete → Delete)
34. Press **Ctrl+Shift+M** again, type **46**, click Run
    - **What should happen:** A "Variance Commentary" sheet appears with auto-generated English narratives
35. Delete the "Variance Commentary" sheet
36. Press **Ctrl+Shift+M** again, type **12**, click Run
    - **What should happen:** An "Executive Dashboard" sheet appears with branded charts
37. Delete the "Executive Dashboard" sheet
38. Press **Ctrl+Shift+M** again, type **63**, click Run
    - A menu with 9 options appears. Type **1** and click OK
    - **What should happen:** A "What-If Impact" sheet appears showing before/after comparison
39. Restore: Press **Ctrl+Shift+M**, type **65**, click Run. Click Yes on the confirmation.
    - The Assumptions sheet should revert to original values
40. Delete "What-If Impact" sheet if it still exists
41. Delete "WhatIf_Baseline" hidden sheet if created (right-click any tab → Unhide → select it → OK → right-click → Delete)

**If all 4 tests ran without errors:** Your demo file macros are working.

---

## Part 5: Set Excel to the Perfect Visual State

> **Note:** The Director macro's pre-flight check can auto-fix zoom, active sheet, and cell selection. But it's best to set these yourself first.

42. Navigate to the **Report-->** sheet tab (click on it)
43. Click on cell **A1** so no random cell is selected
44. Close the Command Center if it is open (click the X on the form)
45. Set the zoom level:
    - Click the **View** tab at the top of Excel
    - Click **Zoom** in the ribbon
    - Select **100%** and click OK
46. Make sure the ribbon is visible at the top (if collapsed, double-click any tab name like "Home" to expand it)
47. Maximize Excel so it fills the entire screen (click the square maximize button in the top-right corner)
48. Press **Ctrl+S** to save the file

**Your demo file is now in the perfect clean state for recording.**

---

### PHASE 1 CHECKLIST

Before moving on, confirm:

- [ ] Demo .xlsm file opens with macros enabled
- [ ] All leftover output sheets have been deleted (file is in clean state)
- [ ] All core sheets exist and have data
- [ ] Ctrl+Shift+M opens the Command Center
- [ ] Data Quality Scan (Action 7) ran without errors
- [ ] Variance Commentary (Action 46) ran without errors
- [ ] Build Dashboard (Action 12) ran without errors
- [ ] What-If Demo (Action 63) ran and Restore Baseline (Action 65) restored successfully
- [ ] All test output sheets have been deleted (back to clean state)
- [ ] File is saved, on Report--> sheet, cell A1 selected, zoom 100%, maximized

---
---

# PHASE 2: IMPORT THE DIRECTOR MODULE

The Director module is what automates the recording. You need to import two files into your Excel workbook.

---

## Import modDirector.bas

1. Open your demo Excel file (if not already open)
2. Press **Alt+F11** to open the VBA Editor
3. In the VBA Editor, click **File** > **Import File...**
4. Navigate to `C:\Users\connor.atlee\RecTrial\VBAToImport\`
5. Select **modDirector.bas**
6. Click **Open**
7. You should now see **modDirector** in the Modules folder in the left panel
8. Press **Alt+Q** to close the VBA Editor and return to Excel
9. Press **Ctrl+S** to save the file

### Verify it imported:

1. Press **Alt+F8** to open the Macro dialog
2. You should see these macros in the list:
   - `RunVideo1`
   - `RunVideo2`
   - `RunVideo3`
   - `RunAllVideos`
   - `TestClip`
   - `QuickTest`
   - `RunPreflight`
   - `CleanupAllOutputSheets`
3. If you see them, the import worked. Click **Cancel** to close.

---

## Import Updated modWhatIf

The Director needs an updated version of modWhatIf that has two new silent subs for the What-If demo clip. Without this, Clip 23 will error out.

1. Press **Alt+F11** to open the VBA Editor
2. In the left panel, look for **modWhatIf** in the Modules folder
3. **If it already exists:** Right-click **modWhatIf** → click **Remove modWhatIf...** → click **No** when asked "Do you want to export the file before removing it?"
4. Click **File** → **Import File...**
5. Navigate to `C:\Users\connor.atlee\RecTrial\VBAToImport\`
6. Select **modWhatIf_v2.1.bas**
7. Click **Open**
8. You should see **modWhatIf** reappear in the Modules folder
9. Press **Ctrl+S** to save
10. Press **Alt+Q** to close the VBA Editor

---

## Set Your Audio File Path

The Director needs to know where your audio clips are stored.

1. Press **Alt+F11** to open the VBA Editor
2. In the left panel, double-click **modDirector** to open it
3. Near the top of the file (~line 79), find this line:

```vba
Private Const AUDIO_BASE_PATH As String = "C:\Users\connor.atlee\RecTrial\AudioClips\"
```

4. **If you are using the RecTrial folder as-is, this path is already correct. Do not change it.**
5. If you moved the AudioClips folder somewhere else, change the path to match. The path MUST end with a backslash (`\`).
6. Press **Ctrl+S** to save
7. Press **Alt+Q** to close the VBA Editor

### Verify the path is correct:

Open File Explorer and navigate to `C:\Users\connor.atlee\RecTrial\AudioClips\`. Confirm you see:
```
AudioClips\
├── Video1\  (7 MP3 files)
├── Video2\  (21 MP3 files)
└── Video3\  (13 MP3 files)
```

---
---

# PHASE 3: SET UP OBS STUDIO

---

## OBS Settings

1. Open OBS Studio
2. Click **Settings** (bottom right)

**Output tab:**
- Recording Path: `C:\Users\connor.atlee\RecTrial\Recordings\`
- Recording Format: **mp4**
- Recording Quality: **High Quality, Medium File Size**

**Video tab:**
- Base Resolution: **1920x1080**
- Output Resolution: **1920x1080**
- FPS: **30**

**Audio tab (IMPORTANT — read carefully):**
- **Desktop Audio: ENABLED** — The Director macro plays audio through Windows. OBS captures it along with the screen. This is how the narration gets into the recording.
- **Mic/Auxiliary Audio: DISABLED** — You do not want your microphone picking up background noise.

3. Click **OK** to save settings

## Add a Display Capture Source

1. In the Sources panel at the bottom, click **+** → **Display Capture**
2. Name it "Excel Recording"
3. Select your main display
4. Click **OK**

## Test Recording

1. Click **Start Recording** in OBS
2. Wait 5 seconds
3. Click **Stop Recording**
4. Find the file in `RecTrial\Recordings\` and play it back
5. Verify: clean video, 1080p, no unwanted audio

---
---

# PHASE 4: PREPARE VIDEO 3 FILES

**Skip this phase if you are only recording Videos 1 and 2.**

Video 3 uses a completely different file — `Sample_Quarterly_Report.xlsx` — to prove that the universal tools work on ANY Excel file, not just the demo.

---

## Set Up the Sample File

49. Find or create a sample file. It should have intentional "mess":
    - Blank rows scattered in the data
    - Text-stored numbers (numbers with a tiny green triangle in the corner)
    - Merged cells
    - Extra spaces in text cells
    - Mixed date formats
    - At least one hidden sheet
    - Unstyled headers (plain default Excel look)
    - Negative numbers not formatted in red
    - At least one error value (#N/A or #REF!)

50. Import the universal toolkit VBA modules into the sample file:
    - With the sample file open, press **Alt+F11**
    - Click **File** → **Import File...**
    - Navigate to `C:\Users\connor.atlee\RecTrial\UniversalToolkit\vba\`
    - Select each .bas file and import it (you may need to import one at a time)
    - Close the VBA Editor

51. **Also import modDirector.bas** into the sample file:
    - Press **Alt+F11** → File → Import File
    - Navigate to `C:\Users\connor.atlee\RecTrial\VBAToImport\`
    - Select **modDirector.bas** → Open
    - Verify `AUDIO_BASE_PATH` is correct (same as Phase 2)
    - Close VBA Editor

52. Save the sample file as **.xlsm** (macro-enabled): File → Save As → change the file type to "Excel Macro-Enabled Workbook (*.xlsm)" → Save

53. Close the sample file. You will reopen it before Video 3.

---
---

# PHASE 5: COMPUTER LOCKDOWN

**Do this EVERY time before you start recording. Every single time. No exceptions. One Teams notification popping up during recording ruins the take.**

1. **Close EVERYTHING** except Excel and OBS:
   - Close all browser tabs and the browser itself
   - Close Outlook completely (not minimized — right-click the taskbar icon → Close Window)
   - Close Teams completely (or click your profile picture → set to **Do Not Disturb**)
   - Close Slack, OneDrive popups, Spotify, everything

2. **Turn off ALL notifications:**
   - Click the notification/clock area in the bottom-right corner of your screen
   - Click **Focus Assist** → select **Alarms Only**
   - OR: Go to Settings → System → Notifications → toggle OFF

3. **Clean your desktop:**
   - Right-click your desktop → **View** → uncheck **Show desktop icons**

4. **Auto-hide the taskbar:**
   - Right-click the taskbar → **Taskbar settings** → turn ON **Automatically hide the taskbar**

5. **Set display resolution and scaling:**
   - Right-click your desktop → **Display settings**
   - Set **Scale** to **100%** (NOT 125% or 150%)
   - Set **Display resolution** to **1920 x 1080**

6. **Plug in your laptop** (prevents battery throttling or sleep)

7. **Put your phone on silent or airplane mode**

8. **Set your desktop wallpaper to a solid dark color** (right-click desktop → Personalize → Background → Solid color → pick dark gray or black)

---
---

# PHASE 6: PRE-FLIGHT AND QUICK TEST

---

## Run the Pre-Flight Check

1. Press **Alt+F8** to open the Macro dialog
2. Select **RunPreflight**
3. Click **Run**
4. A report will appear showing:
   - Whether all audio files exist for all 3 videos
   - Whether the MCI audio subsystem works (shows a test clip duration)
   - Whether Excel is maximized, zoom is 100%, correct sheet is active
   - Whether required sheets exist
5. Everything should say **OK**, **FOUND**, or **YES**.

> **IMPORTANT NOTE:** If the pre-flight says **"Required sheet missing: General Ledger"** — this is OK! Your file uses a different internal name for that sheet ("CrossfireHiddenWorksheet"). Click **OK** and proceed. This does not affect anything.

---

## Run the Quick Test

1. Make sure you are in the demo Excel file
2. Navigate to the **Report-->** sheet
3. Press **Alt+F8** to open the Macro dialog
4. Select **QuickTest**
5. Click **Run**

### What should happen:

- The pre-flight check runs automatically first
- You should **hear audio** playing (the opening hook clip)
- The screen should **scroll down smoothly**
- A message box appears showing: audio status, measured clip duration, and pre-flight result

### If audio did NOT play:

- Check that `AUDIO_BASE_PATH` is set correctly (Phase 2)
- Check that the MP3 files exist in `RecTrial\AudioClips\Video1\`
- Try playing one of the MP3 files manually (double-click it) to verify it works
- Make sure your computer audio is not muted

### If scrolling did NOT work:

- Make sure the Report--> sheet is the active sheet
- Make sure Excel is not in Edit mode (press Escape first)

---
---

# PHASE 7: RECORD

---

## Timeline (~2 hours total)

| Time | What You Are Doing |
|------|-------------------|
| 0:00 - 0:45 | Phases 1-6 (setup, imports, OBS, lockdown, pre-flight) |
| 0:45 - 0:55 | Record Video 1 (~5 min recording + review playback) |
| 0:55 - 1:00 | Run CleanupAllOutputSheets, reset file |
| 1:00 - 1:25 | Record Video 2 (~18 min recording + review playback) |
| 1:25 - 1:35 | Switch to sample file for Video 3 |
| 1:35 - 1:50 | Record Video 3 (~10 min recording + review playback) |
| 1:50 - 2:00 | Review all recordings, re-record any problem clips |

---

## Record Video 1 — "What's Possible" (~5 minutes)

### Pre-recording checklist:

- [ ] All notifications are OFF (Windows Focus Assist → Alarms Only)
- [ ] Desktop is clean (no icons, taskbar auto-hidden)
- [ ] OBS is open and ready to record
- [ ] Command Center is CLOSED
- [ ] Demo file is on Report--> sheet, A1 selected, zoom 100%, maximized

### Step-by-step:

1. **Start OBS recording** (click Start Recording or use your hotkey)
2. **Wait 3 seconds** (so OBS captures a clean start)
3. In Excel, press **Alt+F8**
4. Select **RunVideo1**
5. Click **Run**
6. **DO NOT TOUCH the mouse or keyboard** — the macro is driving everything
7. Watch the screen — you will see:
   - Slow scroll on the landing page (with audio playing)
   - Command Center opening, categories scrolling, "variance" being typed
   - Data Quality Scan running, letter grade appearing
   - Variance Commentary generating (the jaw-drop moment)
   - Executive Dashboard building with charts
   - Return to landing page for closing
8. When the **"Video 1 recording complete!"** message appears:
   - Click **OK** to dismiss it
   - **Stop OBS recording**
9. Play back the recording to check quality

---

## Record Video 2 — "Full Demo Walkthrough" (~18 minutes)

### Before starting:

1. Run **CleanupAllOutputSheets** to reset the file (Alt+F8 → select it → Run)
2. Navigate to **Report-->** sheet
3. Select cell **A1**
4. Press **Ctrl+S** to save the file

### Step-by-step:

1. **Start OBS recording**
2. **Wait 3 seconds**
3. Press **Alt+F8** → Select **RunVideo2** → Click **Run**
4. **DO NOT TOUCH ANYTHING** — this one takes ~18 minutes
5. The macro will:
   - Tour the workbook (clicking through sheet tabs)
   - Open the Command Center and search it
   - Run Data Quality Scan, Reconciliation, Variance Analysis
   - Generate Variance Commentary (jaw-drop moment)
   - Build Dashboard Charts and Executive Dashboard
   - Run PDF Export, Executive Brief, Executive Mode toggle
   - Save a version, run What-If scenario with restore
   - Run Sensitivity Analysis, Integration Test
   - Show Audit Log and Time Saved Calculator
   - Return to Report--> for closing
6. When the **"Video 2 recording complete!"** message appears:
   - Click **OK**
   - **Stop OBS recording**

### What the Director handles automatically (no dialogs):

- **GL Import (Clip 11):** Shows existing data, skips the file dialog
- **PDF Export (Clip 19):** Exports directly to your Desktop as a PDF — no Save As dialog
- **Version Control (Clip 22):** Saves a version directly — no InputBox
- **What-If (Clip 23):** Runs "Revenue Drops 15%" directly, restores baseline directly — no InputBox, no confirmation dialog. This is the "wow moment" clip — it is bulletproof.

---

## Record Video 3 — "Universal Tools" (~10 minutes)

### Before starting:

1. Close or minimize the demo file
2. Open **Sample_Quarterly_Report.xlsm** (from `RecTrial\SampleFile\` or wherever you saved it)
3. Verify modDirector is imported in the sample file and `AUDIO_BASE_PATH` is set correctly
4. Navigate to the first data sheet in the sample file

### Step-by-step:

1. Make sure you are in the **sample file** (not the demo file)
2. **Start OBS recording**
3. **Wait 3 seconds**
4. Press **Alt+F8** → Select **RunVideo3** → Click **Run**
5. The macro will warn you if it detects you're on the demo file — click Yes to continue or No to stop and switch files
6. **DO NOT TOUCH ANYTHING** for ~10 minutes
7. When the **"Video 3 recording complete!"** message appears:
   - Click **OK**
   - **Stop OBS recording**

### Note:
Some universal toolkit macros prompt for input (thresholds, sheet names). The Director uses SendKeys to pre-stage answers. If a dialog appears and doesn't auto-close, just press **Enter** and the macro will continue.

---

## If something goes wrong during recording:

- If a macro errors out, click **End** (NOT Debug)
- Run **CleanupAllOutputSheets** to reset the file
- Navigate back to **Report-->** sheet
- Re-run the video macro

## To re-record just one clip:

1. Run **CleanupAllOutputSheets** to reset
2. Navigate to the correct starting sheet
3. Press **Alt+F11** to open VBA Editor
4. Press **Ctrl+G** to open the Immediate Window
5. Type: `TestClip 4` (replace 4 with your clip number)
6. Press **Enter**
7. Record just that segment with OBS

---

## Recording Day Checklist

### Morning — Video 1 + Video 2
- [ ] Computer lockdown (notifications off, desktop clean, taskbar hidden)
- [ ] Demo file open, macros enabled, on Report--> sheet
- [ ] Run CleanupAllOutputSheets
- [ ] Run QuickTest — confirm audio and scrolling work
- [ ] Start OBS recording
- [ ] Run RunVideo1 — watch it play through (~5 min)
- [ ] Stop OBS when completion message appears
- [ ] Run CleanupAllOutputSheets to reset
- [ ] Start OBS recording
- [ ] Run RunVideo2 — watch it play through (~18 min)
- [ ] Stop OBS when completion message appears

### Afternoon — Video 3
- [ ] Switch to Sample_Quarterly_Report.xlsm
- [ ] Verify modDirector is imported and audio path is set
- [ ] Start OBS recording
- [ ] Run RunVideo3 — watch it play through (~10 min)
- [ ] Stop OBS when completion message appears
- [ ] Review all recordings

---
---

# PHASE 8: IF SOMETHING GOES WRONG

---

## Macro errors out with a VBA error dialog
- Click **End** (NOT Debug)
- The action failed but Excel is fine
- Run **CleanupAllOutputSheets** to reset
- Re-run the video or use **TestClip** for just the failed clip

## Excel freezes or hangs
- Wait 30 seconds — some macros are just slow
- If still frozen after 60 seconds, press **Ctrl+Break** to interrupt VBA
- If totally unresponsive, open Task Manager (Ctrl+Shift+Esc), find Excel, click **End Task**, and restart

## No audio plays
- Check that `AUDIO_BASE_PATH` is set correctly in modDirector
- Check that the MP3 files actually exist in the AudioClips subfolders
- Try playing one MP3 manually (double-click it) to verify it works
- Make sure your computer audio is not muted
- The Director auto-resets the audio system at every entry point, so "stuck state" from interrupted runs is automatically cleared. If audio still fails, restart Excel.

## Command Center doesn't open
- The Director tries to show frmCommandCenter (non-blocking)
- If the UserForm doesn't exist, it skips and runs the macro directly — the demo still works, you just don't see the CC form on camera
- To fix: Run **BuildCommandCenter** once (Alt+F8 → BuildCommandCenter → Run)

## Pre-flight says "Required sheet missing: General Ledger"
- This is OK. Your file uses "CrossfireHiddenWorksheet" as the internal name for the GL sheet. Click **OK** and proceed. The macros use the correct internal name and will work fine.

## What-If won't restore baseline
1. The Director uses RestoreBaselineSilent which has no dialog — if it fails, try manually: Alt+F8 → Run **RestoreBaseline** → click Yes
2. If that fails: right-click any tab → Unhide → find "WhatIf_Baseline" → unhide it → manually copy values back to Assumptions
3. Nuclear option: close without saving and reopen from the last saved version

## PDF export fails
- Check that a PDF printer is configured (Microsoft Print to PDF)
- The Director exports directly — it doesn't use the Save As dialog

## A notification pops up during recording
- Stop recording immediately
- Dismiss the notification
- Re-do Computer Lockdown
- Re-record that video from scratch (or use TestClip for specific clips)

## OBS recording has no audio
- In OBS Settings → Audio, make sure **Desktop Audio** is ENABLED
- The Director plays audio through Windows — this counts as desktop audio
- Test: play any MP3 file manually while OBS is recording, then check the recording

## Screen looks jumpy or too fast
- You can adjust scroll speed by editing `SCROLL_STEP_DELAY` in modDirector (default 250ms — increase to 400 for slower)
- You can adjust typing speed with `TYPING_DELAY_MS` (default 90ms — increase to 120 for slower)

## I want to re-record just one clip
- Run **CleanupAllOutputSheets** first
- Get the file to the right starting state
- Start OBS recording
- Run `TestClip N` from the VBA Immediate Window (Alt+F11 → Ctrl+G → type `TestClip 4` → Enter)
- Stop OBS recording

---
---

# PHASE 9: TITLE CARDS (After Recording)

Title cards are branded slides that go at the beginning and end of each video, and between chapters in Videos 2 and 3. You add these in the video editor AFTER recording — they are NOT recorded live.

You need PNG images created in PowerPoint (16:9 widescreen slides exported as PNG):

**Video 1:**
- Opening title card: iPipeline Blue background, "Finance Automation / What's Possible"
- Closing title card: "Want to learn more?" with SharePoint link

**Video 2:**
- Opening title card: "Finance Automation — Full Demo Walkthrough"
- 7 chapter cards (Navy background): The Workbook & Command Center, Data Import & Quality, Analysis, Reporting & Visuals, Enterprise Features, Under the Hood, Next Steps
- Closing title card

**Video 3:**
- Opening title card: "Universal Tools — For Any Excel File"
- 4 chapter cards for each tool category
- Closing title card

---
---

# REFERENCE: What the Director Does — Clip by Clip

This section shows what the Director macro does during each clip. You don't need to do any of this manually — the macro handles it all. This is here so you can follow along and know what's happening on screen.

---

## VIDEO 1 — "What's Possible" (Clips 1-7)

### Clip 1 — Title Card (5 seconds)
- **Audio:** None
- **What happens:** Static pause on Report--> page. Title card added in editor.

### Clip 2 — Opening Hook (~30 sec)
- **Audio:** V1_S1_Opening_Hook.mp3
- **Narration:** "This is a single Excel file. Nothing to install, nothing to configure — you just open it and go. Inside are 62 automated actions that handle reporting, analysis, data quality checks, charts, exports, and more — each one triggered with a single click. In the next few minutes, I'm going to show you what that looks like."
- **What happens:** Slow scroll down the Report--> landing page while narration plays.

### Clip 3 — Command Center (~40 sec)
- **Audio:** V1_S2_Command_Center.mp3
- **Narration:** "Everything runs from one place — the Command Center. You can open it with Ctrl+Shift+M, or from the button on the landing page. Every action is organized by category — Monthly Operations, Analysis, Reporting, Enterprise Features, and more. You can scroll through or just search. Type what you're looking for and it filters instantly. Find the action, click Run, and it handles the rest. Let me show you a few examples."
- **What happens:** Command Center opens, categories scroll, "variance" is typed in search box, filtered results shown, CC closes.

### Clip 4 — Data Quality Scan (~40 sec)
- **Audio:** V1_S3_Data_Quality.mp3
- **Narration:** "First — data quality. Before you do anything with your numbers, you want to know if the data is clean. One click, and it scans your entire workbook across six categories — completeness, accuracy, consistency, formatting, outliers, and cross-references. It gives you a letter grade — right there at the top. You get a full breakdown underneath showing exactly where issues are, if any. Fifteen seconds, start to finish."
- **What happens:** CC opens briefly with Action 7, Data Quality Scan runs, letter grade appears, scroll through breakdown. Sheet deleted after.

### Clip 5 — Variance Commentary — JAW-DROP (~45 sec)
- **Audio:** V1_S4_Variance_Commentary.mp3
- **Narration:** "Next — one of the most useful features in the whole file. After running a variance analysis, the system can automatically generate written commentary for the top five variances. These are plain English narratives — ready to drop into an email, a report, or a presentation. It identifies the line item, the dollar and percentage change, and describes what happened. No copying numbers into a paragraph. No writing it yourself. One click."
- **What happens:** CC opens with Action 46, Variance Commentary runs, 3-second silent pause (let viewer read), slow scroll through narratives. Sheet deleted after.

### Clip 6 — Executive Dashboard (~40 sec)
- **Audio:** V1_S5_Dashboard.mp3
- **Narration:** "When it's time to present to leadership, you need visuals — not spreadsheets. One click builds a complete executive dashboard — KPI cards at the top, a waterfall chart showing how you get from revenue to net income, and a product comparison underneath. It is styled, formatted, and ready to present. You can also export the entire report package to a single PDF."
- **What happens:** CC opens with Action 12, Executive Dashboard builds, pause on KPI cards, scroll to waterfall chart, scroll to product comparison. Sheet deleted after.

### Clip 7 — Bridge + Closing (~60 sec)
- **Audio:** V1_S6_Bridge.mp3 then V1_S7_Closing.mp3
- **Narration (Bridge):** "That's a sample of what this file can do. But there's a second piece — a universal code library. The same automation tools work on any Excel file, not just this one. If you want to clean up a messy spreadsheet, compare two files, or build a dashboard from scratch — the tools are there."
- **Narration (Closing):** "Everything you just saw runs from this one Excel file. All the guides, the code library, and the file itself are available on SharePoint. Check the Finance Automation folder for everything you need. Thanks for watching."
- **What happens:** Navigate to Report--> page, static hold during both audio clips.

---

## VIDEO 2 — "Full Demo Walkthrough" (Clips 8-26)

### Clip 8 — Opening (~40 sec)
- **Audio:** V2_S0_Opening.mp3
- **What happens:** Slow scroll on Report--> landing page.

### Clip 9 — Workbook Tour (~85 sec)
- **Audio:** V2_S1a_Workbook.mp3
- **What happens:** Clicks through sheet tabs: P&L Monthly Trend, Functional P&L Summary, Product Line Summary, Assumptions, General Ledger.

### Clip 10 — Command Center (~45 sec)
- **Audio:** V2_S1b_CommandCenter.mp3
- **What happens:** Opens CC, searches "reconciliation", pauses on results, closes.

### Clip 11 — GL Import (~45 sec)
- **Audio:** V2_S2_GL_Import.mp3
- **What happens:** Shows CC with Action 17, navigates to General Ledger sheet, scrolls through data. (Skips actual file dialog.)

### Clip 12 — Data Quality Scan (~50 sec)
- **Audio:** V2_S3_Data_Quality.mp3
- **What happens:** Runs ScanAll, shows letter grade, scrolls category breakdown.

### Clip 13 — Reconciliation (~45 sec)
- **Audio:** V2_S4_Reconciliation.mp3
- **What happens:** Runs RunAllChecks, shows Checks sheet with PASS/FAIL results.

### Clip 14 — Variance Analysis (~40 sec)
- **Audio:** V2_S5_Variance_Analysis.mp3
- **What happens:** Runs RunVarianceAnalysis, scrolls through flagged items.

### Clip 15 — Variance Commentary — JAW-DROP (~45 sec)
- **Audio:** V2_S6_Variance_Commentary.mp3
- **What happens:** Runs GenerateCommentary, 3-second silent pause, scrolls narratives.

### Clip 16 — YoY Variance (~30 sec)
- **Audio:** V2_S7_YoY_Variance.mp3
- **What happens:** Runs YoY Variance Analysis, scrolls results.

### Clip 17 — Dashboard Charts (~45 sec)
- **Audio:** V2_S8_Dashboard_Charts.mp3
- **What happens:** Runs BuildDashboard, pauses on chart grid, scrolls through charts.

### Clip 18 — Executive Dashboard (~30 sec)
- **Audio:** V2_S9_Executive_Dashboard.mp3
- **What happens:** Builds Executive Dashboard, shows KPI cards, waterfall, product comparison.

### Clip 19 — PDF Export (~30 sec)
- **Audio:** V2_S10_PDF_Export.mp3
- **What happens:** Exports report sheets directly to PDF on Desktop. No dialog.

### Clip 20 — Executive Brief (~40 sec)
- **Audio:** V2_S10b_ExecBrief.mp3
- **What happens:** Runs GenerateExecBrief, scrolls through 5 sections.

### Clip 21 — Executive Mode (~20 sec)
- **Audio:** V2_S11_Executive_Mode.mp3
- **What happens:** Toggles Executive Mode ON (tabs disappear), pauses, toggles OFF (tabs return).

### Clip 22 — Version Control (~25 sec)
- **Audio:** V2_S12_Version_Control.mp3
- **What happens:** Saves a version snapshot directly as "March Draft 1". No dialog.

### Clip 23 — What-If Scenario — THE WOW MOMENT (~90 sec)
- **Audio:** V2_S13_WhatIf.mp3
- **What happens:** Runs "Revenue Drops 15%" preset directly (no dialog), shows What-If Impact sheet, scrolls impact analysis, navigates to Assumptions to show changed values, returns to impact sheet, restores baseline directly (no dialog), cleans up.

### Clip 24 — Sensitivity Analysis (~35 sec)
- **Audio:** V2_S13b_Sensitivity.mp3
- **What happens:** Runs sensitivity analysis, scrolls results.

### Clip 25 — Integration Test (~30 sec)
- **Audio:** V2_S14_Integration_Test.mp3
- **What happens:** Runs full integration test, shows 18/18 PASS result.

### Clip 26 — Audit Log + Time Saved + Closing
- **Audio:** V2_S15_Audit_Log.mp3, V2_S13c_TimeSaved.mp3, V2_S16_Closing.mp3
- **What happens:** Shows audit log, runs Time Saved Calculator, returns to Report--> for closing narration.

---

## VIDEO 3 — "Universal Tools" (Clips 27-38)

### Clip 27 — Opening (~45 sec)
- **Audio:** V3_S0_Opening.mp3
- **What happens:** Shows messy sample file, slow scroll through data.

### Clip 28 — Data Sanitizer (~60 sec)
- **Audio:** V3_C1A_DataSanitizer.mp3
- **What happens:** Runs PreviewSanitizeChanges (shows what would change), then RunFullSanitize (cleans the data).

### Clip 29 — Highlights (~35 sec)
- **Audio:** V3_C1B_Highlights.mp3
- **What happens:** Runs HighlightByThreshold (highlights values > 5000), runs HighlightDuplicateValues. Clears highlights after.

### Clip 30 — Comments (~40 sec)
- **Audio:** V3_C1C_Comments.mp3
- **What happens:** Counts comments, extracts all comments to a new sheet.

### Clip 31 — Tab Organizer (~50 sec)
- **Audio:** V3_C2A_TabOrganizer.mp3
- **What happens:** Colors tabs by keyword, reorders tabs.

### Clip 32 — Column Ops (~50 sec)
- **Audio:** V3_C2B_ColumnOps.mp3
- **What happens:** Splits a column, combines columns.

### Clip 33 — Sheet Tools (~50 sec)
- **Audio:** V3_C2C_SheetTools.mp3
- **What happens:** Creates sheet index with hyperlinks, clones a template sheet.

### Clip 34 — Compare Sheets (~50 sec)
- **Audio:** V3_C3A_Compare.mp3
- **What happens:** Compares two sheets cell by cell, shows diff report.

### Clip 35 — Consolidate (~40 sec)
- **Audio:** V3_C3B_Consolidate.mp3
- **What happens:** Consolidates multiple sheets into one with source tracking.

### Clip 36 — Pivot Tools + Lookup/Validation (~60 sec)
- **Audio:** V3_C3C_PivotTools.mp3 then V3_C3D_LookupValidation.mp3
- **What happens:** Lists all pivots, builds VLOOKUP, creates dropdown validation list.

### Clip 37 — Universal Command Center (~50 sec)
- **Audio:** V3_C4_CommandCenter.mp3
- **What happens:** Opens the Universal Toolkit Command Center.

### Clip 38 — Closing (~45 sec)
- **Audio:** V3_Closing.mp3
- **What happens:** Navigates to first sheet, holds static during closing narration.

---
---

# REFERENCE: Action Numbers and Macro Names

## Demo File Actions (Command Center)

| Action # | Action Name | VBA Module | Used In Clip(s) |
|----------|-------------|------------|-----------------|
| 3 | Reconciliation Checks | modReconciliation | Clip 13 |
| 5 | Sensitivity Analysis | modSensitivity | Clip 24 |
| 6 | Variance Analysis (MoM) | modVarianceAnalysis | Clip 14 |
| 7 | Data Quality Scan | modDataQuality | Clips 4, 12 |
| 10 | PDF Export | modPDFExport | Clip 19 |
| 12 | Build Dashboard | modDashboard | Clips 6, 17 |
| 17 | Import Data Pipeline | modImport | Clip 11 |
| 32 | Save Version | modVersionControl | Clip 22 |
| 41 | View Audit Log | modLogger | Clip 25 |
| 44 | Run Full Integration Test | modIntegrationTest | Clip 25 |
| 46 | Generate Variance Commentary | modVarianceAnalysis | Clips 5, 15 |
| 47 | YoY Variance Analysis | modVarianceAnalysis | Clip 16 |
| 48 | Toggle Executive Mode | modNavigation | Clip 21 |
| 63 | Run What-If Demo | modWhatIf | Clip 23 |
| 65 | Restore Baseline | modWhatIf | Clip 23 |
| — | GenerateExecBrief | modExecBrief | Clip 20 |
| — | ShowTimeSavedReport | modTimeSaved | Clip 26 |

## Universal Toolkit Macros (Video 3 — run via Alt+F8)

| Macro Name | Module | Used In Clip |
|------------|--------|-------------|
| PreviewSanitizeChanges | modUTL_DataSanitizer | Clip 28 |
| RunFullSanitize | modUTL_DataSanitizer | Clip 28 |
| HighlightByThreshold | modUTL_Highlights | Clip 29 |
| HighlightDuplicateValues | modUTL_Highlights | Clip 29 |
| ClearHighlights | modUTL_Highlights | Clip 29 |
| CountComments | modUTL_Comments | Clip 30 |
| ExtractAllComments | modUTL_Comments | Clip 30 |
| ColorTabsByKeyword | modUTL_TabOrganizer | Clip 31 |
| ReorderTabs | modUTL_TabOrganizer | Clip 31 |
| SplitColumn | modUTL_ColumnOps | Clip 32 |
| CombineColumns | modUTL_ColumnOps | Clip 32 |
| ListAllSheetsWithLinks | modUTL_SheetTools | Clip 33 |
| TemplateCloner | modUTL_SheetTools | Clip 33 |
| CompareSheets | modUTL_Compare | Clip 34 |
| ConsolidateSheets | modUTL_Consolidate | Clip 35 |
| BuildVLOOKUP | modUTL_LookupBuilder | Clip 36 |
| CreateDropdownList | modUTL_ValidationBuilder | Clip 36 |
| ListAllPivots | modUTL_PivotTools | Clip 36 |
| LaunchUTLCommandCenter | modUTL_CommandCenter | Clip 37 |

---
---

# REFERENCE: Clip Master Index

All clips at a glance:

| Clip | Video | Section | Duration | What Happens |
|------|-------|---------|----------|-------------|
| 1 | V1 | Title Card | 5 sec | Static branded title (add in editor) |
| 2 | V1 | Opening Hook | 30 sec | Scroll landing page |
| 3 | V1 | Command Center | 40 sec | Open CC, browse, search "variance" |
| 4 | V1 | Data Quality Scan | 40 sec | Run Action 7, show letter grade |
| 5 | V1 | Variance Commentary | 45 sec | Run Action 46, jaw-drop narratives |
| 6 | V1 | Executive Dashboard | 40 sec | Run Action 12, scroll charts |
| 7 | V1 | Bridge + Closing | 60 sec | Narration over landing page |
| 8 | V2 | Opening | 40 sec | Scroll landing page |
| 9 | V2 | Workbook Tour | 85 sec | Click through sheet tabs |
| 10 | V2 | Command Center | 45 sec | Open CC, scroll, search |
| 11 | V2 | GL Import | 45 sec | Show General Ledger data |
| 12 | V2 | Data Quality | 50 sec | Run Action 7, letter grade |
| 13 | V2 | Reconciliation | 45 sec | Run Action 3, PASS/FAIL |
| 14 | V2 | Variance Analysis | 40 sec | Run Action 6, flagged items |
| 15 | V2 | Variance Commentary | 45 sec | Run Action 46, JAW-DROP |
| 16 | V2 | YoY Variance | 30 sec | YoY variance analysis |
| 17 | V2 | Dashboard Charts | 45 sec | Run Action 12, 8 charts |
| 18 | V2 | Executive Dashboard | 30 sec | KPI cards, waterfall |
| 19 | V2 | PDF Export | 30 sec | Direct PDF to Desktop |
| 20 | V2 | Executive Brief | 40 sec | 5-section executive brief |
| 21 | V2 | Executive Mode | 20 sec | Toggle tabs on/off |
| 22 | V2 | Version Control | 30 sec | Save snapshot directly |
| 23 | V2 | What-If Scenario | 90 sec | Preset #1 + restore — THE WOW MOMENT |
| 24 | V2 | Sensitivity | 35 sec | Sensitivity analysis |
| 25 | V2 | Integration Test | 30 sec | 18/18 PASS |
| 26 | V2 | Audit + Time + Close | 90 sec | Audit log, time saved, closing |
| 27 | V3 | Data Sanitizer | 60 sec | Preview + full sanitize |
| 28 | V3 | Highlights | 35 sec | Threshold + duplicates |
| 29 | V3 | Comments | 40 sec | Count + extract |
| 30 | V3 | Tab Organizer | 50 sec | Color + reorder tabs |
| 31 | V3 | Column Ops | 50 sec | Split + combine columns |
| 32 | V3 | Sheet Tools | 50 sec | Sheet index + template cloner |
| 33 | V3 | Compare Sheets | 50 sec | Cell-by-cell diff report |
| 34 | V3 | Consolidate | 40 sec | Stack sheets with source tracking |
| 35 | V3 | Pivot + Lookup | 60 sec | VLOOKUP + Validation + Pivots |
| 36 | V3 | Universal CC | 50 sec | Full toolkit Command Center |
| 37 | V3 | Closing | 45 sec | Recap + CTA |

---

## Quick Reference: Running the Macros

| What You Want | Macro to Run | How to Run |
|---|---|---|
| Full system check | `RunPreflight` | Alt+F8 → RunPreflight → Run |
| Test audio + scrolling | `QuickTest` | Alt+F8 → QuickTest → Run |
| Record Video 1 (~5 min) | `RunVideo1` | Alt+F8 → RunVideo1 → Run |
| Record Video 2 (~18 min) | `RunVideo2` | Alt+F8 → RunVideo2 → Run |
| Record Video 3 (~10 min) | `RunVideo3` | Alt+F8 → RunVideo3 → Run (from sample file) |
| Record all 3 videos | `RunAllVideos` | Alt+F8 → RunAllVideos → Run |
| Test one specific clip | `TestClip N` | VBA Immediate Window: `TestClip 4` |
| Clean up after recording | `CleanupAllOutputSheets` | Alt+F8 → CleanupAllOutputSheets → Run |

---

*Master Recording Guide — Created 2026-03-30*
*Merges: finalguidev3CA.md + DIRECTOR_MACRO_SETUP_GUIDE.md*
*For: iPipeline Finance Automation Video Demo Project*
