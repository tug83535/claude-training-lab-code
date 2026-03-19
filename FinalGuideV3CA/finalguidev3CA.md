# ONE-DAY RECORDING PLAYBOOK — iPipeline Finance Automation Videos

**What This Is:** Your single-day, start-to-finish guide for recording all 3 demo videos in one session. Every clip is numbered sequentially (Clips 1–39), every word of narration is written out, every screen action is spelled out step by step, and every transition is explicitly called out. Nothing is assumed. Nothing is skipped.

**What You Are Recording:**
- **Video 1:** "What's Possible" — 4–5 minute overview for all 2,000+ employees (Clips 1–7)
- **Video 2:** "Full Demo Walkthrough" — 15–18 minute deep dive for Finance & Accounting (Clips 8–26)
- **Video 3:** "Universal Tools" — 8–10 minute demo for anyone who uses Excel (Clips 27–39)

**Total Clips:** 39
**Total Recording Time:** Approximately 30–33 minutes of raw footage across all 3 videos
**Estimated Session Length:** 6–8 hours including setup, practice runs, resets, and breaks

---

## Table of Contents

1. [EXCEL PRE-SETUP — Before Any Recording Starts](#excel-pre-setup--before-any-recording-starts)
2. [COMPUTER LOCKDOWN — Every Single Time](#computer-lockdown--every-single-time)
3. [SINGLE-DAY TIMELINE](#single-day-timeline)
4. [VIDEO 1 — "What's Possible" (Clips 1–7)](#video-1--whats-possible-clips-17)
5. [VIDEO 2 — "Full Demo Walkthrough" (Clips 8–26)](#video-2--full-demo-walkthrough-clips-826)
6. [VIDEO 3 — "Universal Tools" (Clips 27–39)](#video-3--universal-tools-clips-2739)
7. [EMERGENCY RECOVERY — If Something Goes Wrong](#emergency-recovery--if-something-goes-wrong)
8. [QUICK REFERENCE — Action Numbers](#quick-reference--action-numbers)

---
---

# EXCEL PRE-SETUP — Before Any Recording Starts

**This section must be completed BEFORE you record a single clip. Do not skip any step. Do not assume anything is already done. Start here.**

The goal of this section is to get your Excel demo file and your sample file into the exact correct starting state so that every macro, every click, and every visual looks perfect on camera.

---

## PART 1: Open the Demo File and Enable Macros

1. Find your demo file: `iPipeline_PnL_Demo.xlsm` (or whatever your final .xlsm file is named). It should have all 39 VBA modules already imported inside it.
2. Double-click the file to open it in Excel.
3. **Look at the very top of the Excel window.** If you see a yellow bar that says **"SECURITY WARNING — Macros have been disabled"** with an "Enable Content" button:
   - Click **Enable Content**
   - This allows the VBA macros to run. Without this, nothing will work.
4. If you do NOT see that yellow bar, macros are already enabled. Move on.

### Make macros always work for this file (so the yellow bar never appears):

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

**Why this matters:** If macros are not enabled, pressing Ctrl+Shift+M will do nothing. The Command Center will not open. No macros will run. Your recording will be useless.

---

## PART 2: Delete ALL Leftover Output Sheets

Every time you run a macro, it creates an output sheet. If any of these exist from a previous session, they will show up on camera and confuse the viewer. You need a perfectly clean starting state.

19. Look at the sheet tabs at the very bottom of the Excel window
20. Right-click on each of the following sheets (if they exist) and click **Delete**, then click **Delete** again on the confirmation popup. If a sheet does not exist, skip it and move to the next one:
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
    - **Charts & Visuals** — only delete if it looks like it has charts from a previous macro run. If it is a permanent sheet with pre-built content, leave it alone.
    - Any sheet starting with **"VER_"** — these are version control snapshots. Delete all of them.
    - Any sheet starting with **"BKP_"** — these are backup copies. Delete all of them.

21. Now check for hidden sheets that also need to be deleted:
    - Right-click on ANY sheet tab at the bottom
    - Click **Unhide...**
    - If you see a sheet called **"WhatIf_Baseline"**, select it, click **OK**, then right-click it and **Delete** it
    - Right-click any tab again → **Unhide...** again
    - If you see **"VBA_AuditLog"**, select it, click **OK** to unhide it
    - Click on the VBA_AuditLog sheet tab
    - Select all the data below the header row (click on row 2, then press Ctrl+Shift+End to select everything)
    - Press the **Delete** key on your keyboard to clear all the data (leave the headers in row 1)
    - Right-click the VBA_AuditLog tab → click **Hide** (to hide it again)

22. Also check the **Checks** sheet:
    - Click on the **Checks** tab
    - If there is any data below the headers, select it all (click row 2, Ctrl+Shift+End) and press Delete
    - The Checks sheet should be empty except for headers

---

## PART 3: Verify the Core Sheets Exist and Have Data

Your demo file should have these sheets. Click on each one to verify it exists and has data:

23. **Report-->** — This is your landing page. It should have iPipeline branding (navy headers, blue accents). This is where every video starts.
24. **P&L - Monthly Trend** — This is the main financial data sheet. It should have revenue and expense line items going across multiple months (at minimum Jan and Feb). Scroll right to verify you see month columns.
25. **Product Line Summary** — Should show revenue by product (iGO, Affirm, InsureSight, DocFast).
26. **Functional P&L - Monthly Trend** — Functional expense breakdown view.
27. **Functional P&L Summary** tabs (e.g., "Mar 25" or individual month tabs) — At least one should exist.
28. **Assumptions** — Should have driver values (growth rates, allocation percentages). This sheet is critical for the What-If demo.
29. **General Ledger** — Should have raw transaction data (Date, Account, Description, Amount columns).
30. **Checks** — Should exist but be empty (you cleared it in step 22).

**If any of these sheets are missing:** STOP. You cannot record until the demo file has all required sheets with data. Go back to the VBA module import process or find a backup copy of the file.

---

## PART 4: Test That Key Macros Actually Work

Before you record anything, you need to confirm that the macros you will demo actually run without errors. If a macro crashes during recording, you have to re-do the entire clip.

31. Press **Ctrl+Shift+M** on your keyboard
    - **What should happen:** The Command Center form pops up — a window with a list of all 65 actions organized by category, with a search box at the top
    - **If nothing happens:** The keyboard shortcut is not set up. Instead, press **Alt+F8**, type `LaunchCommandCenter` in the box, and click **Run**. If that also fails, the VBA modules are not imported correctly. STOP and fix this.
32. In the Command Center, type **7** in the action number box (or find "Data Quality Scan") and click **Run/OK**
    - **What should happen:** The macro runs for 2–5 seconds, then a new sheet called "Data Quality Report" appears with a letter grade at the top and a category breakdown below
    - **If it errors:** STOP. Fix the error before proceeding.
33. Delete the "Data Quality Report" sheet you just created (right-click tab → Delete → Delete)
34. Press **Ctrl+Shift+M** again, type **46**, click Run (this is "Generate Variance Commentary")
    - **What should happen:** A "Variance Commentary" sheet appears with auto-generated English narratives
    - **If it errors or the sheet is blank:** You may only have one month of data. You need at least 2 months in P&L Monthly Trend.
35. Delete the "Variance Commentary" sheet
36. Press **Ctrl+Shift+M** again, type **12**, click Run (this is "Build Dashboard")
    - **What should happen:** An "Executive Dashboard" sheet appears with branded charts (KPI cards, waterfall chart, product comparison)
    - **If it errors:** Chart generation failed. Check that the data sheets have real data.
37. Delete the "Executive Dashboard" sheet
38. Press **Ctrl+Shift+M** again, type **63**, click Run (this is "Run What-If Demo")
    - A menu with 9 options appears. Type **1** (Revenue drops 15%) and click OK
    - **What should happen:** A "What-If Impact" sheet appears showing before/after comparison
39. Now restore: Press **Ctrl+Shift+M**, type **65**, click Run (Restore Baseline). Click Yes on the confirmation.
    - The Assumptions sheet should revert to original values
40. Delete the "What-If Impact" sheet if it still exists
41. Delete the "WhatIf_Baseline" hidden sheet if it was created (right-click any tab → Unhide → select it → OK → right-click → Delete)

**If all 4 tests above ran without errors:** Your demo file is ready.

---

## PART 5: Set Excel to the Perfect Visual State

42. Navigate to the **Report-->** sheet tab (click on it)
43. Click on cell **A1** so no random cell is selected with a blinking cursor in a weird spot
44. Close the Command Center if it is open (click the X on the form)
45. Set the zoom level:
    - Click the **View** tab at the top of Excel
    - Click **Zoom** in the ribbon
    - Select **100%** and click OK
    - (Alternative: use the zoom slider in the bottom-right corner of Excel and drag it to exactly 100%)
46. Make sure the ribbon is visible at the top (if the ribbon is collapsed and you only see tab names, double-click any tab name like "Home" to expand it)
47. Maximize Excel so it fills the entire screen (click the square maximize button in the top-right corner, or double-click the title bar)
48. Press **Ctrl+S** to save the file

**Your demo file is now in the perfect clean state for recording.**

---

## PART 6: Prepare the Sample File for Video 3

Video 3 uses a completely different file — `Sample_Quarterly_Report.xlsx` — to prove that the universal tools work on ANY file, not just the demo.

49. Find the sample file. It should be in `FinalExport/` or you may need to create one.
50. Open the sample file in a second Excel window
51. Verify it has intentional "mess" baked in:
    - A few blank rows scattered in the data (rows with nothing in them, randomly placed)
    - Some text-stored numbers (numbers with a tiny green triangle in the top-left corner of the cell)
    - Some merged cells in Column A (department names spanning multiple rows)
    - Extra spaces in text cells (invisible trailing spaces)
    - Mixed date formats in a date column (some say 01/15/2026, others say Jan 15, 2026, others say 2026-01-15)
    - At least one hidden sheet (right-click any tab → if "Unhide" is available, there's a hidden sheet — good)
    - Unstyled headers (plain default Excel look, no colors, no bold)
    - A few negative numbers that are NOT formatted in red
    - At least one error value (#N/A or #REF!) somewhere
52. The universal toolkit VBA modules (all the `modUTL_*.bas` files) must be imported into this sample file:
    - With the sample file open, press **Alt+F11** to open the VBA Editor
    - Click **File** → **Import File...**
    - Navigate to `FinalExport/UniversalToolkit/vba/`
    - Select ALL .bas files and import them (you may need to import one at a time — repeat for each file)
    - Close the VBA Editor (click the X or press Alt+Q)
53. Save the sample file as **.xlsm** (macro-enabled): File → Save As → change the file type dropdown to "Excel Macro-Enabled Workbook (*.xlsm)" → Save
54. Close the sample file for now. You will reopen it before Video 3.

---

## PART 7: Prepare Python Environment for Video 3

Video 3 has two Python demo clips. You need Python installed and working.

55. Open **Command Prompt** (click the Windows Start button, type `cmd`, press Enter)
56. Type `python --version` and press Enter
    - **What you should see:** Something like `Python 3.9.7` or `Python 3.11.4` — any version 3.7+ is fine
    - **If you see an error:** Python is not installed. Go to python.org, download and install it. Make sure to check "Add Python to PATH" during installation.
57. Type `pip install pandas openpyxl pdfplumber scipy` and press Enter
    - Wait for the packages to install (may take 1–2 minutes)
    - These are required by the Python scripts
58. Prepare two test files for the Python file comparison demo:
    - Create two similar Excel files on your Desktop: `Budget_v1.xlsx` and `Budget_v2.xlsx`
    - They should have the same structure but 5–10 differences (changed numbers, added rows)
    - These will be compared in Clip 37
59. Prepare a sample PDF for the PDF extractor demo:
    - Find any PDF that has a visible data table in it (a financial statement, an invoice, a report)
    - The PDF must have **selectable text** — not a scanned image. You should be able to highlight text in it.
    - Save it to your Desktop
60. Close Command Prompt for now. You will reopen it before the Python clips.

---

## PART 8: Set Up Your Screen Recorder (OBS Studio)

61. Open OBS Studio
62. In the **Sources** panel at the bottom, click the **+** button → **Display Capture** → **OK** → **OK**
63. You should see your screen in the preview
64. Click **Settings** (bottom right):
    - **Output** tab: Set Recording Path to a folder you will remember (e.g., `Desktop/VideoRecordings`). Set Recording Format to **mp4**.
    - **Video** tab: Set Base Resolution to **1920x1080**. Set Output Resolution to **1920x1080**. Set FPS to **30**.
    - **Audio** tab: Set **Desktop Audio** to **Disabled**. Set **Mic/Auxiliary Audio** to **Disabled**. (You do NOT want any sound recorded. The AI narration audio gets added later in the video editor.)
65. Click **Apply**, then **OK**
66. Test: Click **Start Recording**, move your mouse for 5 seconds, click **Stop Recording**. Go to your recording folder. Open the file. Verify: clear video, 1080p, NO audio, smooth movement.

---

## PART 9: Have Your AI Audio Clips Ready

67. All 37 AI narration audio clips should already be generated in ElevenLabs and saved as MP3 files. They should be organized in folders:
    - `Audio/Video1/` — 7 clips (V1_S1 through V1_S7)
    - `Audio/Video2/` — 16 clips (V2_S0 through V2_S15)
    - `Audio/Video3/` — 14 clips (V3_S0 through V3_S13)
68. Put on your headphones and test that the audio clips play correctly. You will listen to each clip in your headphones while recording the screen actions.

**If you have not generated the audio clips yet:** You need to do this first using ElevenLabs (see the Video Production Guide for instructions). You cannot record screen clips without the matching audio to follow.

---

## PART 10: Have Your Title Card Images Ready

69. You should have PNG title card images created in PowerPoint (see the Video Production Guide for instructions). You need:
    - Video 1 opening title card and closing title card
    - Video 2 opening title card, 7 chapter cards, and closing title card
    - Video 3 opening title card, 4 chapter cards, and closing title card
70. These get added in the video editor after recording. You do NOT need them during screen recording. But confirm they exist.

---

### PRE-SETUP COMPLETE CHECKLIST

Before moving on, confirm ALL of the following:

- [ ] Demo .xlsm file opens with macros enabled
- [ ] All leftover output sheets have been deleted (file is in clean state)
- [ ] All core sheets exist and have data (Report-->, P&L Monthly Trend, Assumptions, General Ledger, etc.)
- [ ] Ctrl+Shift+M opens the Command Center successfully
- [ ] Data Quality Scan (Action 7) ran without errors
- [ ] Variance Commentary (Action 46) ran without errors
- [ ] Build Dashboard (Action 12) ran without errors
- [ ] What-If Demo (Action 63) ran and Restore Baseline (Action 65) restored successfully
- [ ] All test output sheets have been deleted (back to clean state)
- [ ] Demo file is saved, on the Report--> sheet, cell A1 selected, zoom 100%, maximized
- [ ] Sample file for Video 3 has messy data and universal toolkit modules imported
- [ ] Python is installed with required packages
- [ ] Two test Excel files ready for Python comparison demo
- [ ] Sample PDF ready for PDF extractor demo
- [ ] OBS is configured (1920x1080, 30fps, mp4, NO audio)
- [ ] OBS test recording looks clean
- [ ] AI audio clips are generated and organized in folders
- [ ] Title card images are created

**If every box is checked, you are ready to record. Move to Computer Lockdown.**

---
---

# COMPUTER LOCKDOWN — Every Single Time

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
   - (You can turn this back on after recording)
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

# SINGLE-DAY TIMELINE

This is a single-day recording session. Here is how to structure your day.

## MORNING (8:00 AM – 12:00 PM)

| Time | What You Are Doing |
|------|-------------------|
| 8:00 – 9:00 | Excel Pre-Setup (if not already done). Computer Lockdown. Test all macros one final time. |
| 9:00 – 9:15 | **Practice Run — Video 1.** Play each audio clip in headphones and practice the screen actions WITHOUT recording. Get comfortable with the timing. |
| 9:15 – 10:00 | **RECORD VIDEO 1 (Clips 1–7).** Record each clip. Play back each one immediately after recording to check quality. Re-record any bad clips. |
| 10:00 – 10:15 | **Break.** Stand up, stretch, get water. |
| 10:15 – 10:30 | **Reset the demo file for Video 2.** Delete all output sheets created during Video 1 recording. Get back to clean state. Save. Close and reopen the file (so the splash screen fires fresh for Clip 8). |
| 10:30 – 10:45 | **Practice Run — Video 2 (first few clips).** Listen to the first 3–4 audio clips and practice. |
| 10:45 – 12:00 | **RECORD VIDEO 2 (Clips 8–19).** This covers through PDF Export. Take your time. |

## LUNCH (12:00 PM – 12:45 PM)

Take a real break. Eat. Step away from the screen.

## AFTERNOON (12:45 PM – 5:00 PM)

| Time | What You Are Doing |
|------|-------------------|
| 12:45 – 1:00 | Review where you left off. Check that the demo file is in the right state for the next clip. |
| 1:00 – 2:00 | **RECORD VIDEO 2 (Clips 20–26).** This covers Executive Brief through Time Saved Calculator closing. |
| 2:00 – 2:15 | **Break.** |
| 2:15 – 2:45 | **Switch to Video 3 setup.** Close the demo file. Open the Sample_Quarterly_Report.xlsm file. Import universal toolkit VBA modules if not already done. Verify the messy data is intact. Open Command Prompt for the Python clips. Practice the first few Video 3 clips. |
| 2:45 – 4:00 | **RECORD VIDEO 3 (Clips 27–39).** |
| 4:00 – 4:30 | **Review ALL recordings.** Play back every clip quickly. Note any that need re-recording. |
| 4:30 – 5:00 | **Re-record any problem clips.** Reset the file to the right state and re-do just the clip(s) that need it. |

## EVENING (if needed)

If you finish early, great — you are done recording. The next step is editing (syncing audio to video, adding title cards) which is a separate session.

---
---

# VIDEO 1 — "What's Possible" (Clips 1–7)

**Runtime Target:** 4:00–5:00
**Audience:** All iPipeline employees (2,000+)
**Purpose:** Build awareness, show high-impact features, point to next steps
**File:** Demo .xlsm file (the P&L file with all 39 VBA modules)

---

## Video 1 File Prep

Before recording Clip 1, confirm:

- [ ] Demo file is open
- [ ] You are on the **Report-->** sheet tab
- [ ] Cell **A1** is selected (no blinking cursor in a weird spot)
- [ ] The Command Center is **CLOSED**
- [ ] No output sheets exist (no Data Quality Report, no Variance Analysis, no Dashboard, etc.)
- [ ] Zoom is set to **100%**
- [ ] Excel is **maximized** to fill the screen
- [ ] OBS is open and ready to record
- [ ] Headphones are on with the V1_S1 audio clip loaded
- [ ] Computer lockdown is complete (no notifications, clean desktop, taskbar hidden)

---

## CLIP 1 — Title Card
**Sequential Number:** Clip 1 of 39
**Duration:** 5 seconds
**Macro/Action:** None
**Audio Clip:** None (music sting added in editing)

### [START CLIP]

**What to record:** Nothing. This is a static title card that you will add in the video editor during the editing phase. You do NOT need to record this on screen.

**Alternative:** If you want to record it live, open the title card image (`V1_Title_Open.png`) full screen and record 5 seconds of it sitting still.

**On screen:**
```
iPipeline Finance Automation
What's Possible
```
iPipeline Blue (#0B4779) background, white text, Arial font.

### [END CLIP]

**What to do next:** Load audio clip V1_S1 in your headphones. Make sure Excel is showing the Report--> sheet. Move to Clip 2.

---

## CLIP 2 — Opening Hook: Landing Page
**Sequential Number:** Clip 2 of 39
**Duration:** ~30 seconds
**Macro/Action:** None (just scrolling)
**Audio Clip:** V1_S1_Opening_Hook.mp3

### Narration Script (exact words):

> "This is a single Excel file. Nothing to install, nothing to configure — you just open it and go.
>
> Inside are 62 automated actions that handle reporting, analysis, data quality checks, charts, exports, and more — each one triggered with a single click.
>
> In the next few minutes, I'm going to show you what that looks like."

### Screen Actions — Step by Step:

1. Excel is open, showing the **Report-->** landing page. Cell A1 is selected.
2. Start OBS recording (click **Start Recording** in OBS, or use the hotkey if you set one)
3. Wait 2 seconds of silence (just hold still — this makes editing easier)
4. Press Play on the audio clip V1_S1 in your headphones
5. As the audio starts talking, **slowly scroll down** through the Report--> page using your mouse scroll wheel
   - Scroll gently — one notch at a time on the scroll wheel
   - The viewer needs to see the branded headers, the layout, the real data
   - Do NOT click on anything — just scroll
6. When the audio says "62 automated actions" — pause the scroll for 1 second. Let that number land visually.
7. When the audio says "I'm going to show you what that looks like" — stop scrolling. Hold still.
8. Wait 2 seconds of silence after the audio ends
9. Stop OBS recording

### [START CLIP]

- OBS recording starts
- 2 seconds silence
- Audio plays → you slowly scroll the Report page
- Audio ends → 2 seconds silence
- OBS recording stops

### [END CLIP]

**Gotcha:** Make sure no random cell is selected with a blinking cursor in a weird spot. Click cell A1 before starting.

**What to do next:** Load audio clip V1_S2. Keep the Report--> page visible. Move to Clip 3.

---

## CLIP 3 — Command Center Introduction
**Sequential Number:** Clip 3 of 39
**Duration:** ~40 seconds
**Macro/Action:** Open Command Center (Ctrl+Shift+M)
**Audio Clip:** V1_S2_Command_Center.mp3

### Narration Script (exact words):

> "Everything runs from one place — the Command Center.
>
> You can open it with Ctrl+Shift+M, or from the button on the landing page.
>
> Every action is organized by category — Monthly Operations, Analysis, Reporting, Enterprise Features, and more. You can scroll through or just search.
>
> Type what you're looking for and it filters instantly. Find the action, click Run, and it handles the rest.
>
> Let me show you a few examples."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds of silence
3. Press Play on V1_S2 audio
4. When the audio says "the Command Center" — press **Ctrl+Shift+M** on your keyboard
5. The Command Center form pops up. **Wait 1–2 seconds** after it appears before doing anything. Let the viewer take it in.
6. When the audio mentions categories — **slowly scroll down** through the Command Center list using the scroll bar inside the form. Move smoothly. Let people read a few action names as they scroll by.
7. When the audio says "just search" — **click in the search box** at the top of the Command Center
8. Type the word **variance** (lowercase, type it slowly so the viewer can see)
9. The list filters to show only variance-related actions. **Pause 2 seconds** on the filtered results.
10. **Clear the search box** (select all text in the search box and delete it, or click the X if there is one)
11. When the audio says "Let me show you a few examples" — **close the Command Center** (click the X button on the form, or click Cancel)
12. Wait 2 seconds of silence
13. Stop OBS recording

### [START CLIP]

- 2 sec silence → audio plays → open Command Center → scroll categories → search "variance" → clear search → close Command Center → 2 sec silence

### [END CLIP]

**Gotcha:** If Ctrl+Shift+M does nothing, the keyboard shortcut is not set up. Use Alt+F8, type `LaunchCommandCenter`, click Run instead. This looks slightly worse on camera but still works.

**Gotcha #2:** If the Command Center falls back to an InputBox menu (a plain gray dialog instead of a styled form), the frmCommandCenter UserForm was not built in the workbook. The form looks much better on camera. Test this beforehand.

**What to do next:** Load V1_S3 audio. Move to Clip 4.

---

## CLIP 4 — Data Quality Scan + Letter Grade
**Sequential Number:** Clip 4 of 39
**Duration:** ~40 seconds
**Macro/Action:** Action 7 — Data Quality Scan (modDataQuality)
**Audio Clip:** V1_S3_Data_Quality.mp3

### Narration Script (exact words):

> "First — data quality. Before you do anything with your numbers, you want to know if the data is clean.
>
> One click, and it scans your entire workbook across six categories — completeness, accuracy, consistency, formatting, outliers, and cross-references.
>
> It gives you a letter grade — right there at the top. In this case, [read the actual grade that appears]. You get a full breakdown underneath showing exactly where issues are, if any.
>
> Fifteen seconds, start to finish."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds of silence
3. Press Play on V1_S3 audio
4. When the audio says "Before you do anything with your numbers" — press **Ctrl+Shift+M** to open the Command Center
5. Type **7** in the action number box (or find "Data Quality Scan" in the list)
6. When the audio says "One click" — click **Run** (or OK)
7. The macro runs for 2–5 seconds. Excel will auto-navigate to the new **"Data Quality Report"** sheet. Wait for it to finish.
8. When the sheet appears — **PAUSE 2–3 seconds** on the letter grade badge at the top. It is a large colored badge (28pt font). This is the visual anchor. Let the viewer see it.
9. When the audio says "full breakdown underneath" — **slowly scroll down** through the category breakdown (Completeness, Accuracy, Consistency, Formatting, Outliers, Cross-References)
10. When the audio says "Fifteen seconds, start to finish" — **stop scrolling and hold still**
11. Wait 2 seconds of silence
12. Stop OBS recording

### [START CLIP]

- 2 sec silence → audio plays → open Command Center → run Action 7 → Data Quality Report sheet appears → hold on letter grade → scroll breakdown → hold still → 2 sec silence

### [END CLIP]

**Gotcha:** If the demo file is perfectly clean, the letter grade will be A and all categories will show 0 issues. This is fine — it shows the tool works.

**IMPORTANT — Reset after this clip:** Right-click the **Data Quality Report** tab → **Delete** → confirm. This sheet must not exist when you start the next clip.

**What to do next:** Delete the Data Quality Report sheet. Load V1_S4 audio. Move to Clip 5.

---

## CLIP 5 — Variance Commentary (Jaw-Drop Feature)
**Sequential Number:** Clip 5 of 39
**Duration:** ~45 seconds
**Macro/Action:** Action 46 — Generate Variance Commentary (modVarianceAnalysis)
**Audio Clip:** V1_S4_Variance_Commentary.mp3

### Narration Script (exact words):

> "Next — one of the most useful features in the whole file.
>
> After running a variance analysis, the system can automatically generate written commentary for the top five variances.
>
> These are plain English narratives — ready to drop into an email, a report, or a presentation. It identifies the line item, the dollar and percentage change, and describes what happened.
>
> No copying numbers into a paragraph. No writing it yourself. One click."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds of silence
3. Press Play on V1_S4 audio
4. Press **Ctrl+Shift+M** to open the Command Center
5. Type **46** in the action number box
6. When the audio reaches "generate written commentary" — click **Run**
7. Wait 3–5 seconds for generation
8. Excel navigates to the new **"Variance Commentary"** sheet
9. **PAUSE 2–3 SECONDS.** This is the jaw-drop moment. Do NOT talk or move your mouse. Let the viewer read the auto-generated narratives.
10. After the pause, **slowly scroll down** through the narratives. Each row has a line item name, variance amount, percentage, and a plain-English paragraph.
11. **Hover your mouse cursor near** (not on top of) one particularly good narrative to draw the viewer's eye to it
12. When the audio says "One click" — hold still
13. Wait 2 seconds of silence
14. Stop OBS recording

### [START CLIP]

- 2 sec silence → audio plays → open Command Center → run Action 46 → Variance Commentary sheet appears → PAUSE to let viewer read → scroll narratives → hover near one → hold still → 2 sec silence

### [END CLIP]

**IMPORTANT — Reset after this clip:** Right-click the **Variance Commentary** tab → **Delete** → confirm.

**What to do next:** Delete the Variance Commentary sheet. Load V1_S5 audio. Move to Clip 6.

---

## CLIP 6 — Executive Dashboard
**Sequential Number:** Clip 6 of 39
**Duration:** ~40 seconds
**Macro/Action:** Action 12 — Build Dashboard (modDashboard)
**Audio Clip:** V1_S5_Dashboard.mp3

### Narration Script (exact words):

> "When it's time to present to leadership, you need visuals — not spreadsheets.
>
> One click builds a full executive dashboard — KPI summary cards at the top, a waterfall chart showing how you get from budget to actual, and a product line comparison. All branded, all formatted, all ready to present.
>
> You can also build a full set of eight charts on a separate sheet, or export everything to a clean, formatted PDF — headers, footers, page numbers included."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds of silence
3. Press Play on V1_S5 audio
4. Press **Ctrl+Shift+M** to open the Command Center
5. Type **12** in the action number box
6. When the audio says "One click builds" — click **Run**
7. Wait 5–10 seconds (this builds multiple charts — it takes longer than other macros)
8. Excel navigates to the **"Executive Dashboard"** sheet
9. When the dashboard appears — **pause 2 seconds.** The dashboard is visually rich. Let the viewer absorb it.
10. **Slowly scroll down** through the dashboard:
    - First the KPI summary cards at the top — pause 1–2 seconds
    - Then the waterfall chart — pause 2 seconds (this is the most impressive visual)
    - Then the product line comparison at the bottom — pause 1–2 seconds
11. When the audio mentions "PDF" and "eight charts" — hold still. You are NOT navigating to those. Just let the audio play.
12. Wait 2 seconds of silence
13. Stop OBS recording

### [START CLIP]

- 2 sec silence → audio plays → open Command Center → run Action 12 → dashboard appears → pause → scroll through KPI cards, waterfall, product comparison → hold still → 2 sec silence

### [END CLIP]

**IMPORTANT — Reset after this clip:** Right-click the **Executive Dashboard** tab → **Delete** → confirm.

**What to do next:** Delete the Executive Dashboard sheet. Load V1_S6/S7 audio. Move to Clip 7.

---

## CLIP 7 — Bridge to Universal Tools + Closing
**Sequential Number:** Clip 7 of 39
**Duration:** ~60 seconds (Bridge 30 sec + Closing 30 sec)
**Macro/Action:** None (narration only)
**Audio Clip:** V1_S6_Bridge.mp3 + V1_S7_Closing.mp3 (or combined into one clip)

### Narration Script — Bridge (exact words):

> "That's a sample of what this file can do for a P&L close process. But the code behind it — the VBA, the Python, the SQL — isn't locked to this one file.
>
> There's a library of universal tools that work on any Excel spreadsheet. Formatting, cleanup, sorting, searching across sheets — all reusable, all documented, all available on SharePoint.
>
> There's a separate video walking through those tools if you want to see what's available for your own work."

### Narration Script — Closing (exact words):

> "Everything you just saw runs from this one Excel file — nothing to install, no cost, no IT involvement.
>
> If you want to explore the file yourself, watch the full demo walkthrough, or grab tools from the code library — it's all on SharePoint. There are step-by-step guides for everything, and if you need help, there are pre-built Copilot prompts to walk you through it.
>
> Thanks for watching."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds of silence
3. Navigate back to the **Report-->** landing page (click the Report--> tab)
4. Press Play on the bridge audio clip
5. While the audio plays — keep the Report--> page visible. **Do NOT click anything.** Just let the audio play over the static landing page.
6. Optionally, **slowly scroll down** the landing page during the bridge section
7. When the closing audio starts ("Everything you just saw...") — hold still on the landing page
8. When the audio says "Thanks for watching" — hold still for 3 seconds
9. Wait 2 seconds of silence
10. Stop OBS recording

### [START CLIP]

- 2 sec silence → navigate to Report page → audio plays (bridge then closing) → static landing page visible → "Thanks for watching" → 3 sec hold → 2 sec silence

### [END CLIP]

**Note:** The closing title card (branded slide with SharePoint link and contact info) gets added in the video editor. You do NOT need to record it.

---

### VIDEO 1 COMPLETE

**Congratulations — Video 1 is done.** You have recorded Clips 1–7.

**Before moving to Video 2:**
1. Delete ALL output sheets that were created during Video 1 (Data Quality Report, Variance Commentary, Executive Dashboard — if any still exist)
2. Clear the Checks sheet data (if anything was written to it)
3. Clear the VBA_AuditLog (unhide → select data below headers → Delete → re-hide)
4. Delete any WhatIf_Baseline hidden sheet (if it exists)
5. Navigate back to the **Report-->** tab
6. Click cell **A1**
7. Close the Command Center if open
8. **Save the file** (Ctrl+S)
9. **Close Excel completely**
10. Take a 15-minute break

---
---

# VIDEO 2 — "Full Demo Walkthrough" (Clips 8–26)

**Runtime Target:** 15:00–18:00
**Audience:** Finance & Accounting team, interested power users
**Purpose:** Detailed tour of the P&L demo file — show how things work, step by step
**Scope:** Demo file ONLY. No mention of universal tools — that is Video 3.
**File:** Demo .xlsm file

---

## Video 2 File Prep

**CRITICAL — The file must be in a completely fresh state for Video 2.** The splash screen fires when you open the file, and that is your first clip. So:

1. The demo file should be **CLOSED** right now (you closed it at the end of Video 1)
2. Start OBS recording FIRST (before opening the file)
3. THEN double-click the .xlsm file to open it
4. The splash screen fires — that becomes Clip 8

**Before you open the file, confirm:**
- [ ] Excel is completely closed (no Excel windows open at all)
- [ ] OBS is ready to record
- [ ] Computer lockdown is still in effect (no notifications, clean desktop)
- [ ] Headphones are on with V2_S0 audio clip ready
- [ ] You know where the demo .xlsm file is (Desktop or wherever you saved it)

---

## CLIP 8 — File Opens + Splash Screen
**Sequential Number:** Clip 8 of 39
**Duration:** ~15 seconds
**Macro/Action:** Automatic — splash screen fires on file open (modSplashScreen)
**Audio Clip:** First few seconds of V2_S0_Opening.mp3 (or silence — you may add audio in editing)

### Narration Script (this clip may be silent, with audio added in editing):

This clip is primarily visual. The splash screen itself is the content. You may record this as a silent clip and overlay the opening narration in the editor.

### Screen Actions — Step by Step:

1. **Start OBS recording FIRST** (before opening Excel)
2. Wait 2 seconds of silence
3. Double-click the demo **.xlsm** file on your Desktop (or from File Explorer)
4. Excel opens → the splash screen fires automatically
5. The splash shows branded information:
   - "KEYSTONE BENEFITECH" (or iPipeline branding)
   - "P&L Reporting & Allocation Model"
   - Version info, module count, action count
   - "Press Ctrl+Shift+M to open the Command Center"
6. **Let the splash sit on screen for 3–4 seconds.** Let the viewer read it.
7. Click **OK** to dismiss the splash
8. If a second dialog asks "Would you like to open the Command Center?" — click **No** (you will open it later on camera)
9. The file opens to the **Report-->** landing page
10. Wait 2 seconds
11. Stop OBS recording

### [START CLIP]

- OBS starts recording → 2 sec silence → double-click file → splash appears → hold 3–4 sec → click OK → click No → Report page visible → 2 sec → stop

### [END CLIP]

**Gotcha:** If macros are disabled, the splash won't fire. You should have set up Trusted Locations in the Pre-Setup section.

**What to do next:** Load V2_S0 audio. Move to Clip 9.

---

## CLIP 9 — Opening + Workbook Tour
**Sequential Number:** Clip 9 of 39
**Duration:** ~85 seconds (Opening 40 sec + Tour 45 sec)
**Macro/Action:** None (just clicking sheet tabs)
**Audio Clip:** V2_S0_Opening.mp3 + V2_S1_Chapter1_Workbook.mp3

### Narration Script — Opening (exact words):

> "Welcome to the full walkthrough of the iPipeline Finance Automation file.
>
> This is a single Excel workbook that automates the monthly P&L close process — from importing data, to running quality checks, to generating analysis, building dashboards, and producing final deliverables. Sixty-two actions, all accessible from one control panel.
>
> I'm going to walk you through how it works, chapter by chapter. Everything you see is running live — no slides, no mockups. Let's get into it."

### Narration Script — Workbook Tour (exact words):

> "Let's start with what's inside this file.
>
> The landing page — Report — gives you a summary of the workbook and quick navigation to any section. Think of it as your table of contents.
>
> The file has over a dozen sheets. You've got the main P&L Monthly Trend sheet — this is your core financial data, revenue and expenses by month, with a full-year total and budget column.
>
> There are Functional P&L Summary sheets — one for each month — that break things down by department.
>
> A Product Line Summary sheet showing revenue by product — iGO, Affirm, InsureSight, DocFast.
>
> An Assumptions sheet with the key financial drivers — growth rates, allocation percentages, revenue shares.
>
> And a General Ledger sheet with the raw transaction data.
>
> You don't need to memorize any of this — because everything runs from one place."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds of silence
3. Play the Opening audio. The file is on the Report--> page. **Slowly scroll down** the landing page as the audio narrates.
4. When the Opening audio ends and the Tour audio begins saying "Let's start with what's inside this file":
5. Click the **P&L - Monthly Trend** tab at the bottom. **Pause 2 seconds.** Let the viewer see the data layout. Scroll right slightly to show month columns.
6. Click one **Functional P&L Summary** tab (e.g., "Mar 25" or "Jan"). **Pause 1 second.**
7. Click the **Product Line Summary** tab. **Pause 2 seconds.**
8. Click the **Assumptions** tab. **Pause 1–2 seconds.**
9. Click the **General Ledger** tab. **Pause 1 second.**
10. When the audio says "everything runs from one place" — navigate back to the **Report-->** tab
11. Wait 2 seconds of silence
12. Stop OBS recording

### [START CLIP]

- 2 sec silence → audio plays → scroll Report page → click through sheet tabs with pauses → back to Report → 2 sec silence

### [END CLIP]

**Note:** The Chapter 1 title card ("The Workbook & Command Center") gets added in the video editor between clips.

---

## CLIP 10 — Command Center Overview
**Sequential Number:** Clip 10 of 39
**Duration:** ~45 seconds
**Macro/Action:** Open Command Center (Ctrl+Shift+M)
**Audio Clip:** V2_S2_Command_Center.mp3

### Narration Script (exact words):

> "This is the Command Center. Every automated action in this file is listed here, organized by category. You can scroll through to browse, or use the search bar to find what you need.
>
> Monthly Operations, Analysis and Reporting, Enterprise Features, Utilities — it's all here. Pick an action, click Run, and it handles the rest.
>
> That's your home base. Every demo from here on out starts from this screen."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds of silence
3. Press **Ctrl+Shift+M** to open the Command Center
4. Wait 1–2 seconds after it opens
5. Play the audio
6. **Slowly scroll** through the full Command Center list (5–6 seconds of scrolling). Let the viewer see all 65 actions.
7. When the audio mentions categories — hover your mouse over category headers as they scroll by
8. Click in the **search box**, type **reconciliation**
9. **Pause 2 seconds** on the filtered results
10. Clear the search box (select all, delete)
11. **Close** the Command Center (click X or Cancel)
12. Wait 2 seconds of silence
13. Stop OBS recording

### [START CLIP]

- 2 sec silence → open Command Center → audio plays → scroll full list → search "reconciliation" → clear → close → 2 sec silence

### [END CLIP]

---

## CLIP 11 — GL Import
**Sequential Number:** Clip 11 of 39
**Duration:** ~45 seconds
**Macro/Action:** Action 17 — Import Data Pipeline (modImport)
**Audio Clip:** V2_S3_Data_Import.mp3

### Narration Script (exact words):

> "Before you can do any analysis, you need your data. The General Ledger Import pulls in GL data from a CSV or Excel file with format validation built in.
>
> It reads the source file, validates the format, maps the columns, and loads the transactions into the workbook. If something doesn't match the expected structure, it tells you.
>
> What used to take around 45 minutes of manual copying, pasting, and reformatting — done in about 30 seconds."

### Screen Actions — Step by Step:

1. **BEFORE RECORDING:** Have a sample GL data file ready on your Desktop (a .csv or .xlsx with GL columns: Date, Account, Description, Amount). You prepared this during pre-setup.
2. Start OBS recording
3. Wait 2 seconds of silence
4. Press **Ctrl+Shift+M**
5. Type **17** (or find "Import Data Pipeline")
6. Play the audio
7. Click **Run**
8. A file picker dialog appears — navigate to your sample GL file on the Desktop and select it
9. Click **Open**
10. Wait for the import to complete (the status bar may show progress)
11. When complete, navigate to the **General Ledger** sheet to show the loaded data
12. **Slowly scroll down** through a few rows so the viewer can see real transaction data
13. Wait 2 seconds of silence
14. Stop OBS recording

### [START CLIP]

- 2 sec silence → open Command Center → run Action 17 → file picker → select file → import runs → show General Ledger data → scroll → 2 sec silence

### [END CLIP]

**Gotcha:** If you don't have a sample import file ready, this clip will fail. Have the file prepared in advance.

**Alternative:** If the GL data is already in the file from the pre-setup, you can skip the actual import and just narrate over the General Ledger sheet showing the data that's already there. Mention that the import process loaded this data.

---

## CLIP 12 — Data Quality Scan + Letter Grade
**Sequential Number:** Clip 12 of 39
**Duration:** ~50 seconds
**Macro/Action:** Action 7 — Data Quality Scan (modDataQuality)
**Audio Clip:** V2_S4_Data_Quality.mp3

### Narration Script (exact words):

> "Now that the data is loaded, the first thing you want to know is — how clean is it?
>
> The Data Quality Scan checks your entire workbook across six categories: completeness, accuracy, consistency, formatting, outliers, and cross-references.
>
> Right at the top — a letter grade. [Read the actual grade]. That tells you at a glance whether your data is ready to work with.
>
> Below that, each category gets its own score and detail. If there are issues, it tells you exactly where — which sheet, which column, what the problem is.
>
> This scan has never been done manually — there was no practical way to do it. Now it takes about fifteen seconds."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds of silence
3. Press **Ctrl+Shift+M**
4. Type **7**
5. Play the audio
6. Click **Run**
7. Wait 2–5 seconds for the scan
8. The **Data Quality Report** sheet appears
9. **PAUSE 2–3 seconds** on the letter grade badge at the top. Do NOT talk or scroll yet.
10. When the audio says "each category gets its own score" — **slowly scroll down** through the category breakdown
11. If specific issues are flagged, **hover your cursor near** a finding
12. When the audio says "about fifteen seconds" — hold still
13. Wait 2 seconds of silence
14. Stop OBS recording

### [START CLIP]

- 2 sec silence → open Command Center → run Action 7 → report appears → HOLD on letter grade → scroll breakdown → hold still → 2 sec silence

### [END CLIP]

**DO NOT delete the Data Quality Report sheet yet.** It stays visible as you move to the next clip. The viewer sees that output sheets accumulate.

---

## CLIP 13 — Reconciliation Checks
**Sequential Number:** Clip 13 of 39
**Duration:** ~45 seconds
**Macro/Action:** Action 3 — Reconciliation Checks (modReconciliation)
**Audio Clip:** V2_S5_Reconciliation.mp3

### Narration Script (exact words):

> "Next step — make sure all the numbers tie out.
>
> The reconciliation engine runs a series of validation checks across every sheet — verifying that cross-sheet totals match, that revenue and expense lines balance, and that formulas are intact.
>
> Each check gets a clear PASS or FAIL. Green means it ties. Red means something needs attention.
>
> In this case — [describe what you see: e.g., 'all checks passing' or 'one item flagged']. Either way, you know exactly where you stand in ten seconds instead of two hours."

### Screen Actions — Step by Step:

1. Start OBS recording
2. Wait 2 seconds
3. Press **Ctrl+Shift+M**, type **3**, play audio, click **Run**
4. Wait 2–5 seconds
5. The **Checks** sheet populates with results
6. **Pause 2 seconds** on the PASS/FAIL scorecard — the green/red visual is immediately readable
7. **Slowly scroll through** all checks
8. If any FAIL items exist, hover cursor near them
9. Wait 2 seconds of silence
10. Stop OBS recording

### [START CLIP]

- 2 sec silence → run Action 3 → Checks sheet populates → hold on PASS/FAIL → scroll → 2 sec silence

### [END CLIP]

---

## CLIP 14 — Variance Analysis (Month over Month)
**Sequential Number:** Clip 14 of 39
**Duration:** ~40 seconds
**Macro/Action:** Action 6 — Run Variance Analysis (modVarianceAnalysis)
**Audio Clip:** V2_S6_Variance_Analysis.mp3

### Narration Script (exact words):

> "Your data is in, it's clean, and it reconciles. Now — what's actually happening in the numbers?
>
> The Variance Analysis compares each line item month over month and flags anything that moved more than fifteen percent. Revenue, expenses, margins — it checks everything.
>
> Items over the threshold are highlighted automatically. You can see the dollar change, the percentage change, and whether it's favorable or unfavorable. For expense items, the favorable/unfavorable logic is automatically reversed — a decrease in costs is flagged as favorable, not unfavorable.
>
> Instead of scanning hundreds of rows yourself, you get a filtered view of what actually needs your attention."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, type **6**, play audio, click **Run**
3. Wait 2–5 seconds
4. **Variance Analysis** sheet appears
5. Pause on the header row — let the viewer see the column structure
6. **Slowly scroll through**, pausing on highlighted/flagged items
7. **Hover cursor near** a flagged item to draw attention
8. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Action 6 → Variance Analysis appears → scroll through highlighted items → 2 sec silence

### [END CLIP]

---

## CLIP 15 — Variance Commentary (JAW-DROP MOMENT)
**Sequential Number:** Clip 15 of 39
**Duration:** ~45 seconds
**Macro/Action:** Action 46 — Generate Variance Commentary (modVarianceAnalysis)
**Audio Clip:** V2_S7_Variance_Commentary.mp3

### Narration Script (exact words):

> "This is one of the features I'm most excited about.
>
> You've got your flagged variances. Now the system can write the commentary for you.
>
> These are plain English narratives for the top five variances. Each one identifies the line item, states the dollar and percentage change, and describes what happened — in complete sentences, ready to paste into an email, a report, or a board deck.
>
> [pause for 2–3 seconds of silence while text is visible]
>
> Writing these manually — pulling the numbers, doing the comparison, putting it into words — that's typically an hour of work. This takes about five seconds."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, type **46**, play audio, click **Run**
3. Wait 3–5 seconds
4. **Variance Commentary** sheet appears
5. **PAUSE 2–3 SECONDS.** This is the jaw-drop moment. Do NOT move your mouse. Do NOT talk. Let the viewer read.
6. After the pause, **slowly scroll through** each narrative (there should be ~5)
7. **Hover cursor near** one narrative to draw the eye
8. When the audio says "about five seconds" — **hold still for a beat.** Let the impact sit.
9. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Action 46 → commentary appears → PAUSE 2-3 sec → scroll narratives → hover near one → hold after "five seconds" → 2 sec silence

### [END CLIP]

**THIS IS YOUR JAW-DROP MOMENT IN THIS VIDEO. Give it room to breathe.**

---

## CLIP 16 — YoY Variance
**Sequential Number:** Clip 16 of 39
**Duration:** ~30 seconds
**Macro/Action:** YoY Variance action (search in Command Center)
**Audio Clip:** V2_S8_YoY_Variance.mp3

### Narration Script (exact words):

> "Variance Analysis gives you month over month. But leadership often wants year over year — how does this year compare to last year, and how are we tracking against budget?
>
> This builds a full Year-over-Year comparison. Full-year total versus prior year, full-year total versus budget, with dollar and percentage variances for every line.
>
> Same idea — items beyond the threshold are flagged. You get a complete picture of where you're ahead, where you're behind, and by how much.
>
> What would normally take a couple of hours of pulling data from two different periods and building the comparison — done in about ten seconds."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, search for **"YoY"** or **"Year over Year"**, play audio, click **Run**
3. Wait for results
4. **YoY Variance Analysis** sheet appears
5. Pause on the header — let viewer see column structure (FY Total, Prior Year, Budget, $ Variance, % Variance)
6. Slowly scroll through, pausing on flagged items
7. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run YoY action → sheet appears → scroll flagged items → 2 sec silence

### [END CLIP]

---

## CLIP 17 — Dashboard Charts
**Sequential Number:** Clip 17 of 39
**Duration:** ~45 seconds
**Macro/Action:** Action 12 — Build Dashboard (modDashboard)
**Audio Clip:** V2_S9_Dashboard_Charts.mp3

### Narration Script (exact words):

> "You've done the analysis. Now you need to present it.
>
> One click builds eight branded charts in a grid layout — revenue trends, expense breakdowns, margin analysis, product mix, and more. All formatted in iPipeline colors, all properly labeled.
>
> These are the visuals you'd normally build one at a time in a separate PowerPoint or chart tool. Here, they're generated directly from your data in about fifteen seconds."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, type **12**, play audio, click **Run**
3. Wait 5–10 seconds (builds 8 charts)
4. The **Charts & Visuals** or dashboard sheet appears with charts
5. **Slowly scroll/pan** across all charts — pause 1–2 seconds on each chart
6. Let the branded colors and formatting speak for themselves
7. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Action 12 → charts appear → scroll through each chart → 2 sec silence

### [END CLIP]

---

## CLIP 18 — Executive Dashboard
**Sequential Number:** Clip 18 of 39
**Duration:** ~30 seconds
**Macro/Action:** Executive Dashboard action (may be part of Action 12 or separate)
**Audio Clip:** V2_S10_Executive_Dashboard.mp3

### Narration Script (exact words):

> "For a more focused leadership view, there's the Executive Dashboard.
>
> This puts everything on one sheet — KPI summary cards across the top, a waterfall chart showing how you get from budget to actual, and a product line comparison at the bottom.
>
> This is designed to be the one sheet you pull up when the CFO asks 'how are we doing this month.' One click, one sheet, full picture."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. If the Executive Dashboard is a separate action: press **Ctrl+Shift+M**, find "Executive Dashboard", click **Run**
3. If it was already created as part of Action 12: just click the **Executive Dashboard** tab
4. Play the audio
5. **Pause at the top** — KPI cards visible — hold 2 seconds
6. **Scroll down** to waterfall chart — hold 2 seconds
7. **Scroll down** to product comparison — hold 2 seconds
8. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → navigate to/create Executive Dashboard → KPI cards (2 sec) → waterfall (2 sec) → product comparison (2 sec) → 2 sec silence

### [END CLIP]

---

## CLIP 19 — PDF Export
**Sequential Number:** Clip 19 of 39
**Duration:** ~30 seconds
**Macro/Action:** Action 10 — Export Report Package (modPDFExport)
**Audio Clip:** V2_S11_PDF_Export.mp3

### Narration Script (exact words):

> "When you need a final deliverable — something you can email, save to a shared drive, or print — the PDF Export handles it.
>
> It takes seven key sheets from the workbook and compiles them into a single, clean PDF. Each page has proper headers and footers — the report title, the date, page numbers. Formatted for printing or sharing.
>
> Manually formatting and exporting seven sheets to a clean PDF — that's easily thirty minutes of adjusting print areas, fixing page breaks, and hoping nothing shifts. This takes about ten seconds."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, type **10**, play audio, click **Run**
3. A save dialog appears — choose your **Desktop** for easy access
4. Wait 5–10 seconds for the export
5. A success message appears showing the file path
6. **Optional but impressive:** Minimize Excel briefly (click the minimize button), open the PDF file on your Desktop, show the first page (clean formatting, headers/footers visible), scroll to page 2
7. Close the PDF, maximize Excel again
8. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Action 10 → save dialog → Desktop → export completes → (optional: show PDF) → 2 sec silence

### [END CLIP]

---

## CLIP 20 — Executive Brief
**Sequential Number:** Clip 20 of 39
**Duration:** ~40 seconds
**Macro/Action:** GenerateExecBrief (modExecBrief) — may or may not have a Command Center action number
**Audio Clip:** Part of V2_S12 or a dedicated clip

### Narration Script (exact words):

> "The Executive Brief scans five key areas of the workbook — Revenue and P&L highlights, Reconciliation status, Key Assumptions and Drivers, Product Line performance, and Workbook Health — and builds a one-page summary with color-coded status indicators.
>
> Green means healthy. Yellow means attention needed. Red means action required. One click, and you know the full status of your close."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, search for **"Executive Brief"** or find the action
   - If no action number: press **Alt+F8**, type `GenerateExecBrief`, click **Run**
3. Play the audio
4. Wait 3–5 seconds for the scan
5. **Executive Brief** sheet appears with 5 styled sections
6. **Slowly scroll through** each section:
   - Revenue & P&L Highlights
   - Reconciliation Status
   - Key Assumptions & Drivers
   - Product Line Overview
   - Workbook Health
7. Pause on color-coded indicators (green/yellow/red)
8. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Executive Brief → sheet appears → scroll 5 sections → pause on indicators → 2 sec silence

### [END CLIP]

---

## CLIP 21 — Executive Mode Toggle
**Sequential Number:** Clip 21 of 39
**Duration:** ~20 seconds
**Macro/Action:** Action 48 — Toggle Executive Mode (modNavigation)
**Audio Clip:** Part of V2_S12_Executive_Mode.mp3

### Narration Script (exact words):

> "When leadership needs to review the file, they don't need to see every technical sheet. Executive Mode cleans it up.
>
> One click hides all the working sheets and leaves only the presentation-ready views. Toggle it off and everything comes back.
>
> Simple, but it makes a big difference when you're sharing the file with someone who just wants the highlights."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. **Before pressing anything:** look at the tab bar at the bottom. You should see many sheet tabs (10+). Let the viewer see this full tab bar.
3. Press **Ctrl+Shift+M**, type **48**, play audio, click **Run**
4. Watch the tabs at the bottom — several sheets **HIDE**, leaving only key executive-level sheets visible
5. **PAUSE 2 seconds.** Let the viewer see the reduced, clean tab bar. Do NOT talk during this pause.
6. Press **Ctrl+Shift+M** again, type **48**, click **Run** (toggle back)
7. All sheets reappear
8. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → show full tab bar → run Action 48 → tabs disappear → PAUSE 2 sec → run Action 48 again → tabs return → 2 sec silence

### [END CLIP]

---

## CLIP 22 — Version Control: Save Snapshot
**Sequential Number:** Clip 22 of 39
**Duration:** ~30 seconds
**Macro/Action:** Action 32 — Save Version (modVersionControl)
**Audio Clip:** V2_S12 (Version Control section)

### Narration Script (exact words):

> "Version Control lets you save a snapshot of the entire workbook at any point — and compare or restore it later.
>
> You give it a name — 'Pre-Close' or 'March Draft 1' — and it saves the full state. If something goes wrong, or if someone overwrites your work, you go back to Version Control, pick a snapshot, and restore it.
>
> Every snapshot is timestamped and logged. You always know what changed and when."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, type **32**, play audio, click **Run**
3. An InputBox appears asking for a version name
4. Type: **March Close Draft 1** (type it slowly so the viewer can read)
5. Click **OK**
6. A success message confirms the snapshot was saved
7. Click **OK** on the success message
8. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Action 32 → type "March Close Draft 1" → OK → success message → 2 sec silence

### [END CLIP]

---

## CLIP 23 — What-If Scenario Demo (THE WOW MOMENT)
**Sequential Number:** Clip 23 of 39
**Duration:** ~90 seconds
**Macro/Action:** Action 63 — Run What-If Demo + Action 65 — Restore Baseline (modWhatIf)
**Audio Clip:** V2_S13_Scenario_Sensitivity.mp3

### Narration Script (exact words):

> "Now — one of the most powerful features. What-If Scenario Analysis.
>
> You pick a scenario — revenue drops fifteen percent, AWS costs spike twenty-five percent, best case, worst case — and the system modifies the assumptions, recalculates the entire workbook, and shows you the full ripple effect.
>
> [run the scenario]
>
> Look at this. Every P&L line item — before value, after value, dollar impact, percentage change. You can see exactly how a revenue drop cascades through the entire business.
>
> And when you're done, one click restores everything to the original baseline. No manual cleanup."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, type **63**, play audio, click **Run**
3. A menu appears with 9 scenario options (Revenue drops 15%, Revenue increases 10%, AWS costs increase 25%, etc.)
4. Type **1** (Revenue drops 15%) and click **OK**
5. Wait 3–5 seconds — the macro:
   - Saves the current Assumptions as a baseline (to hidden WhatIf_Baseline sheet)
   - Modifies the Assumptions sheet to reflect -15% revenue
   - Recalculates the workbook
   - Creates a "What-If Impact" sheet showing before/after comparison
6. Excel navigates to the **What-If Impact** sheet
7. **PAUSE 3–4 SECONDS.** Let the viewer absorb the impact numbers.
8. **Slowly scroll through** the impact report:
   - Which drivers changed
   - Before value → After value → Dollar impact → Percentage change
   - Ripple effects across P&L line items
9. Now restore: Press **Ctrl+Shift+M**, type **65**, click **Run**
10. A confirmation dialog appears — click **Yes**
11. The Assumptions sheet restores to original values
12. A success message confirms restoration — click **OK**
13. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Action 63 → choose scenario 1 → What-If Impact sheet → PAUSE 3-4 sec → scroll impact → run Action 65 to restore → confirm → success → 2 sec silence

### [END CLIP]

**CRITICAL GOTCHA:** After running What-If, the Assumptions sheet has modified values. If you DON'T restore before the next clip, all subsequent clips will show wrong numbers. **ALWAYS restore (Action 65) before moving on.**

---

## CLIP 24 — Integration Test
**Sequential Number:** Clip 24 of 39
**Duration:** ~30 seconds
**Macro/Action:** Action 44 — Run Full Integration Test (modIntegrationTest)
**Audio Clip:** V2_S14_Integration_Test.mp3

### Narration Script (exact words):

> "With this many automated actions, you need to know the system is working correctly. The Integration Test runs eighteen automated checks across the entire workbook — sheet existence, data integrity, formula health, macro functionality.
>
> Eighteen out of eighteen — all passing.
>
> This runs every time you want to verify the file is in a good state. Before a close, after making changes, anytime you want peace of mind — one click."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, type **44**, play audio, click **Run**
3. Wait 5–10 seconds (runs 18 tests)
4. **Integration Test Report** sheet appears
5. **Pause 3 seconds** on the 18/18 PASS summary. This is a confidence builder.
6. If results are listed individually, slowly scroll through them
7. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Action 44 → tests run → 18/18 PASS → hold 3 sec → scroll → 2 sec silence

### [END CLIP]

---

## CLIP 25 — Audit Log
**Sequential Number:** Clip 25 of 39
**Duration:** ~25 seconds
**Macro/Action:** Action 41 — View Audit Log (modLogger)
**Audio Clip:** V2_S14 (Audit Log section)

### Narration Script (exact words):

> "Every action you run is logged automatically.
>
> The Audit Log records a timestamp, the module that ran, and the result for every single action. If you need to know who ran what, and when — it's all here.
>
> You'll see entries from everything we just did — every import, every scan, every export. Full traceability."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, type **41**, play audio, click **Run**
3. The hidden **VBA_AuditLog** sheet becomes visible and activates
4. **Scroll through** the log entries — you should see timestamps from every macro you ran during this entire demo session
5. **Pause 2–3 seconds** — the viewer sees a full audit trail from the entire demo
6. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Action 41 → Audit Log sheet appears → scroll entries → hold 2-3 sec → 2 sec silence

### [END CLIP]

---

## CLIP 26 — Time Saved Calculator + Closing
**Sequential Number:** Clip 26 of 39
**Duration:** ~45 seconds + closing narration
**Macro/Action:** ShowTimeSavedReport (modTimeSaved)
**Audio Clip:** V2_S15_Closing.mp3

### Narration Script — Time Saved (exact words):

> "One more thing. The Time Saved Calculator compares manual effort versus automated effort for all sixty-two actions. Every row shows the action, how long it takes manually, how long the macro takes, and the time saved."

### Narration Script — Closing (exact words):

> "That's the full walkthrough. We imported GL data, checked data quality, ran reconciliation, analyzed variances month over month and year over year, generated written commentary, built dashboards and an executive view, exported a clean PDF, managed scenarios and sensitivity analysis, ran a full integration test, and reviewed the audit trail.
>
> All from one Excel file, all through the Command Center, all in a matter of minutes.
>
> If you want to explore the file yourself, it's available on SharePoint along with step-by-step training guides for everything you just saw. If you run into any questions, reach out — I'm happy to help.
>
> Thanks for watching."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Ctrl+Shift+M**, search for **"Time Saved"** — click **Run**
   - If no action number: press **Alt+F8**, type `ShowTimeSavedReport`, click **Run**
3. Play the audio
4. Wait 3–5 seconds
5. **Time Saved Analysis** sheet appears with a table of all 62 actions and their time comparisons
6. **Slowly scroll through** the table
7. **Scroll to the bottom** for the **Executive Summary box**:
   - "Manual: X hours per monthly close"
   - "Automated: Y hours per monthly close"
   - "Annual: Z hours per year"
8. **PAUSE 3–4 SECONDS on the annual savings number.** This is the mic-drop closing moment.
9. When the closing narration begins, navigate back to the **Report-->** landing page
10. Hold still on the landing page during the closing narration
11. When the audio says "Thanks for watching" — hold still for 3 seconds
12. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run Time Saved → sheet appears → scroll to Executive Summary → HOLD on annual savings 3-4 sec → navigate to Report page → closing narration → "Thanks for watching" → hold 3 sec → 2 sec silence

### [END CLIP]

---

### VIDEO 2 COMPLETE

**Congratulations — Video 2 is done.** You have recorded Clips 8–26.

**Before moving to Video 3:**
1. **Close the demo .xlsm file entirely** (File → Close, or close Excel)
2. Take a 15-minute break
3. When you come back, you will open the **Sample_Quarterly_Report.xlsm** file for Video 3

---
---

# VIDEO 3 — "Universal Tools" (Clips 27–39)

**Runtime Target:** 8:00–10:00
**Audience:** Anyone at iPipeline who uses Excel and wants to automate their own work
**Purpose:** Show that the universal tools work on ANY file — not just the demo
**Key Principle:** Uses a SEPARATE, SIMPLE sample Excel file — NOT the demo file
**File:** Sample_Quarterly_Report.xlsm (the messy sample file with universal toolkit modules imported)

---

## Video 3 File Prep

**CRITICAL — Video 3 uses a completely different file.** Close the demo file. You are now working with the sample file.

1. Close the demo .xlsm file if it is still open
2. Open **Sample_Quarterly_Report.xlsm** (the messy sample file you prepared during Pre-Setup)
3. Verify macros are enabled (if you see the yellow bar, click Enable Content)
4. Maximize Excel to fill the screen
5. Set zoom to **100%**
6. Navigate to the main data sheet (the one with the messy data — blank rows, merged cells, text-stored numbers visible)
7. Test that the universal toolkit macros are accessible: press **Alt+F8** — you should see a list of macros starting with names like `RunFullSanitize`, `HighlightByThreshold`, `ListAllSheetsWithLinks`, etc.
8. For the Python clips: open a **Command Prompt** window (Start → type `cmd` → Enter) but keep it minimized for now. You will bring it up during Clips 37–38.

**Confirm before recording:**
- [ ] Sample file is open with messy data visible
- [ ] Macros are enabled and accessible (Alt+F8 shows universal toolkit macros)
- [ ] The file has intentional mess: blank rows, text-stored numbers, merged cells, mixed dates, hidden sheets
- [ ] Command Prompt is open (minimized)
- [ ] Two test Excel files for Python comparison are on Desktop
- [ ] Sample PDF for PDF extractor is on Desktop
- [ ] Computer lockdown is still in effect

---

## CLIP 27 — Data Sanitizer: Preview + Full Clean
**Sequential Number:** Clip 27 of 39
**Duration:** ~60 seconds
**Macro/Action:** PreviewSanitizeChanges → RunFullSanitize (modUTL_DataSanitizer)
**Audio Clip:** V3_S1 (Data Cleanup chapter)

### Narration Script (exact words):

> "First problem — text-stored numbers and floating-point noise. The cell shows 1,250 but it's actually a text string. Or it shows 99.99999999997 instead of 100.
>
> Before fixing anything, the Sanitizer can do a dry run — a preview showing what WOULD change without touching your data.
>
> [run preview]
>
> Every cell it would fix, listed with the current value and the proposed fix. Once you're satisfied, run the full sanitize.
>
> [run full sanitize]
>
> It backs up your sheets first, then applies all fixes. Text-stored numbers are converted. Floating-point tails are cleaned. Your data is ready to use."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Show the sample data — scroll briefly to point out cells with green triangles (text-stored numbers) or long decimal numbers
3. Play the audio
4. Press **Alt+F8**, type `PreviewSanitizeChanges`, click **Run**
5. Wait 2–3 seconds
6. A new **UTL_Sanitizer_Preview** sheet appears showing what would change
7. **Scroll through** the preview — each row shows: Sheet, Cell, Current Value, Issue Type, Proposed Fix
8. **Pause 2 seconds** on the preview
9. Now run the full fix: Press **Alt+F8**, type `RunFullSanitize`, click **Run**
10. Click **Yes** on any confirmation dialog
11. Wait 3–5 seconds (it creates backup sheets first, then fixes)
12. Success message shows count of fixes applied — click **OK**
13. Navigate back to the data sheet — the green triangles are gone
14. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → show messy data → run PreviewSanitizeChanges → preview sheet → scroll → run RunFullSanitize → confirm → success → data is clean → 2 sec silence

### [END CLIP]

---

## CLIP 28 — Highlights: Threshold + Duplicates
**Sequential Number:** Clip 28 of 39
**Duration:** ~60 seconds
**Macro/Action:** HighlightByThreshold + HighlightDuplicateValues (modUTL_Highlights)
**Audio Clip:** V3_S2

### Narration Script (exact words):

> "Need to find every amount over ten thousand dollars? One click.
>
> [run threshold highlight]
>
> All cells above the threshold are highlighted instantly. You can do the same for below a threshold, or exact matches.
>
> Now duplicates — if you have a column where values should be unique but aren't, the duplicate highlighter finds them instantly.
>
> [run duplicate highlight]
>
> Orange means it appears more than once. Clear the highlights when you're done."

### Screen Actions — Step by Step:

**Part A — Threshold (30 sec):**
1. Start OBS recording, wait 2 seconds
2. Select a range of numeric cells (e.g., click the column header for an Amount column)
3. Play the audio
4. Press **Alt+F8**, type `HighlightByThreshold`, click **Run**
5. Type **10000** in the threshold InputBox, click OK
6. Type **above** in the direction InputBox, click OK
7. Wait 1 second — cells above $10,000 are highlighted
8. **Pause 2 seconds**

**Part B — Duplicates (30 sec):**
9. Select a column with duplicate values (e.g., a name or category column)
10. Press **Alt+F8**, type `HighlightDuplicateValues`, click **Run**
11. Wait 1 second — duplicate values highlighted in orange
12. **Pause 2 seconds**
13. Clear: press **Alt+F8**, type `ClearHighlights`, click **Run**, choose "Active Sheet"
14. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → select range → threshold highlight → PAUSE → select column → duplicate highlight → PAUSE → clear highlights → 2 sec silence

### [END CLIP]

---

## CLIP 29 — Comments: Extract and Count
**Sequential Number:** Clip 29 of 39
**Duration:** ~40 seconds
**Macro/Action:** CountComments → ExtractAllComments (modUTL_Comments)
**Audio Clip:** V3_S3

### Narration Script (exact words):

> "How many comments are hiding in this workbook? Let's find out.
>
> [run count]
>
> The count shows you how many comments on each sheet. Now extract them all to one place.
>
> [run extract]
>
> Every comment — sheet name, cell address, author, and the full text — all in one report you can review, filter, or share."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Play the audio
3. Press **Alt+F8**, type `CountComments`, click **Run**
4. A message box shows comment count per sheet — **pause 2 seconds** so the viewer can read
5. Click **OK**
6. Press **Alt+F8**, type `ExtractAllComments`, click **Run**
7. Wait 1–2 seconds
8. **UTL_CommentReport** sheet appears with all comments extracted
9. **Scroll through** the report — show Sheet, Cell, Author, Comment columns
10. **Pause 2 seconds**
11. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → count comments → message box → OK → extract comments → report sheet → scroll → 2 sec silence

### [END CLIP]

---

## CLIP 30 — Tab Organizer: Color + Reorder
**Sequential Number:** Clip 30 of 39
**Duration:** ~50 seconds
**Macro/Action:** ColorTabsByKeyword + ReorderTabs (modUTL_TabOrganizer)
**Audio Clip:** V3_S4

### Narration Script (exact words):

> "Your workbook has a dozen tabs. Let's organize them. First — color code by keyword.
>
> [run color tabs]
>
> Every tab with 'Sales' in the name turns green. You can do this for any keyword — 'Budget', 'Q1', 'Archive' — pick a keyword and a color.
>
> Now reorder — move tabs to the front, back, or after a specific sheet with one click.
>
> [run reorder]
>
> Your tab bar is organized and color-coded without dragging a single tab."

### Screen Actions — Step by Step:

**Part A — Color (25 sec):**
1. Start OBS recording, wait 2 seconds
2. Play the audio
3. Press **Alt+F8**, type `ColorTabsByKeyword`, click **Run**
4. Type **Sales** in the keyword InputBox, click OK
5. Choose a color number (e.g., **3** for green), click OK
6. Tabs with "Sales" in the name turn green — **pause 2 seconds**

**Part B — Reorder (25 sec):**
7. Press **Alt+F8**, type `ReorderTabs`, click **Run**
8. A numbered list of sheets appears — pick one to move
9. Choose position (front, back, or after a specific sheet)
10. Tabs reorder on screen — **pause 2 seconds** on the new tab order
11. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → color tabs by keyword → PAUSE → reorder tabs → PAUSE → 2 sec silence

### [END CLIP]

---

## CLIP 31 — Column Ops: Split and Merge
**Sequential Number:** Clip 31 of 39
**Duration:** ~50 seconds
**Macro/Action:** SplitColumn + CombineColumns (modUTL_ColumnOps)
**Audio Clip:** V3_S5

### Narration Script (exact words):

> "A column has combined data — 'Smith, John' or 'New York, NY 10001'. You need it split into separate columns.
>
> [run split]
>
> Pick the delimiter — comma, space, whatever — and it splits into clean separate columns.
>
> Going the other direction — combine two columns into one.
>
> [run combine]
>
> First Name plus Last Name becomes Full Name. Pick your separator and it's done."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Select the column header of a column with combined data (e.g., "Smith, John" format)
3. Play the audio
4. Press **Alt+F8**, type `SplitColumn`, click **Run**
5. Choose **Comma** as the delimiter
6. Wait 1–2 seconds — the column splits into 2+ columns
7. **Pause 2 seconds**
8. Select 2 columns to merge (e.g., the split columns)
9. Press **Alt+F8**, type `CombineColumns`, click **Run**
10. Choose **Space** as separator
11. Wait 1 second — a new column appears with combined values
12. **Pause 2 seconds**
13. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → select column → split → PAUSE → select 2 columns → combine → PAUSE → 2 sec silence

### [END CLIP]

---

## CLIP 32 — Sheet Tools: Sheet Index + Template Cloner
**Sequential Number:** Clip 32 of 39
**Duration:** ~50 seconds
**Macro/Action:** ListAllSheetsWithLinks + TemplateCloner (modUTL_SheetTools)
**Audio Clip:** V3_S6

### Narration Script (exact words):

> "Need an index of every sheet in the workbook? One click.
>
> [run sheet index]
>
> Every sheet listed with a clickable hyperlink. Click any link and it jumps you straight to that sheet.
>
> Need to clone a sheet? The Template Cloner makes exact copies — pick the sheet, pick how many copies, done.
>
> [run cloner]
>
> Three perfect copies, instantly."

### Screen Actions — Step by Step:

**Part A — Sheet Index (20 sec):**
1. Start OBS recording, wait 2 seconds
2. Play the audio
3. Press **Alt+F8**, type `ListAllSheetsWithLinks`, click **Run**
4. **UTL_SheetIndex** sheet appears with: Sheet names, Clickable hyperlinks, Visibility status
5. **Click one hyperlink** to show it works — it jumps to that sheet
6. Navigate back to UTL_SheetIndex

**Part B — Template Cloner (30 sec):**
7. Press **Alt+F8**, type `TemplateCloner`, click **Run**
8. A numbered list of sheets appears — pick a sheet (e.g., the main data sheet)
9. Type **3** for number of copies, click OK
10. Wait 2–3 seconds
11. Three new tabs appear in the tab bar — **pause 2 seconds** to show the clones
12. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → sheet index → click hyperlink → back → template cloner → 3 copies → PAUSE on tabs → 2 sec silence

### [END CLIP]

---

## CLIP 33 — Compare Sheets
**Sequential Number:** Clip 33 of 39
**Duration:** ~50 seconds
**Macro/Action:** CompareSheets (modUTL_Compare)
**Audio Clip:** V3_S7

### Narration Script (exact words):

> "Got two versions of the same sheet and need to find what changed? The Sheet Comparator does a cell-by-cell comparison.
>
> [run compare]
>
> Pick Sheet A, pick Sheet B, and it builds a diff report showing every difference — cell address, the value in each sheet, match or mismatch. It even highlights the differences on the source sheets in red."

### Screen Actions — Step by Step:

**Setup:** Before recording, go into one of the cloned sheets from Clip 32 and **change 3–5 cell values** so the compare tool has something to find.

1. Start OBS recording, wait 2 seconds
2. Play the audio
3. Press **Alt+F8**, type `CompareSheets`, click **Run**
4. A numbered list of sheets appears — pick the original sheet (Sheet A)
5. Pick the modified clone (Sheet B)
6. Click **Yes** on "Highlight differences on source sheets?"
7. Wait 2–5 seconds
8. **UTL_CompareReport** sheet appears with cell-by-cell comparison
9. **Scroll through** — show Cell Address, Sheet A Value, Sheet B Value, Match/Mismatch
10. Navigate to the original sheet — changed cells are highlighted in red
11. **Pause 3 seconds** on the diff report
12. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run CompareSheets → pick Sheet A → pick Sheet B → report appears → scroll diff → show red highlights → PAUSE 3 sec → 2 sec silence

### [END CLIP]

---

## CLIP 34 — Consolidate Sheets
**Sequential Number:** Clip 34 of 39
**Duration:** ~40 seconds
**Macro/Action:** ConsolidateSheets (modUTL_Consolidate)
**Audio Clip:** V3_S8

### Narration Script (exact words):

> "Need to stack data from multiple sheets into one? The Consolidator combines them vertically with a source tracking column so you know which sheet each row came from.
>
> [run consolidate]
>
> Headers from the first sheet, data from all selected sheets, and a Source column on the right. One consolidated view."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Play the audio
3. Press **Alt+F8**, type `ConsolidateSheets`, click **Run**
4. A numbered list of sheets appears — select 2–3 sheets (type their numbers comma-separated)
5. Click **Yes** on "Skip headers on sheets 2+?"
6. Click **Yes** on "Add Source Sheet column?"
7. Wait 2–3 seconds
8. **UTL_Consolidated** sheet appears with all data stacked
9. **Scroll down** to show data from different source sheets and the Source column
10. **Pause 2 seconds**
11. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → run ConsolidateSheets → select sheets → options → consolidated sheet appears → scroll → PAUSE → 2 sec silence

### [END CLIP]

---

## CLIP 35 — PivotTools + LookupBuilder + ValidationBuilder
**Sequential Number:** Clip 35 of 39
**Duration:** ~60 seconds
**Macro/Action:** Multiple tools from modUTL_PivotTools, modUTL_LookupBuilder, modUTL_ValidationBuilder
**Audio Clip:** V3_S9

### Narration Script (exact words):

> "Three quick power tools. First — the Lookup Builder writes VLOOKUP formulas for you. Pick your lookup keys, pick your source table, tell it which column to return. Done. Every formula wrapped in IFERROR so you never see #N/A.
>
> Second — the Validation Builder. Select a column, type your allowed values, and every cell gets a dropdown. Instant data validation.
>
> Third — Pivot Tools. List every pivot table in the workbook, refresh them all with one click, or restyle them."

### Screen Actions — Step by Step:

**Part A — Lookup Builder (25 sec):**
1. Start OBS recording, wait 2 seconds
2. Play the audio
3. Press **Alt+F8**, type `BuildVLOOKUP`, click **Run**
4. Follow the prompts: select lookup keys → select source table → enter column number → select output location
5. VLOOKUP formulas appear with values — **pause 2 seconds**

**Part B — Validation Builder (20 sec):**
6. Select a range of cells (e.g., a Status column)
7. Press **Alt+F8**, type `CreateDropdownList`, click **Run**
8. Type: **Open, Closed, Pending, Cancelled** — click OK
9. Click on one cell — a dropdown arrow appears with the 4 options — **pause 2 seconds**

**Part C — Pivot Tools (15 sec):**
10. Press **Alt+F8**, type `ListAllPivots`, click **Run**
11. If pivots exist: show the report. If not: message says "No pivot tables found" — that is fine.
12. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → VLOOKUP builder → formulas appear → PAUSE → validation dropdown → PAUSE → pivot tools → 2 sec silence

### [END CLIP]

---

## CLIP 36 — Universal Command Center
**Sequential Number:** Clip 36 of 39
**Duration:** ~50 seconds
**Macro/Action:** LaunchUTLCommandCenter (modUTL_CommandCenter)
**Audio Clip:** V3_S10

### Narration Script (exact words):

> "Everything you just saw — plus dozens more tools — is organized in the Universal Command Center. Just like the demo file's Command Center, but for the universal tools.
>
> Every tool categorized: Data Sanitize, Highlights, Comments, Tab Organizer, Column Ops, Sheet Tools, Compare, Consolidate, Pivot Tools, Lookup Builder, Validation. Over 140 tools total.
>
> Search for any keyword — 'duplicate', 'pivot', 'format' — and it filters instantly."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Press **Alt+F8**, type `LaunchUTLCommandCenter`, click **Run**
   - (Or if you set up `LaunchCommandCenter` for the universal toolkit, use that)
3. Play the audio
4. The Universal Command Center menu appears with all categories
5. **Slowly scroll through** the full menu — let the viewer see ALL categories
6. Type **duplicate** in the search — show filtered results — **pause 2 seconds**
7. Clear search, type **pivot** — show results — **pause 2 seconds**
8. Clear search
9. Close the Command Center
10. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → launch Universal Command Center → scroll full menu → search "duplicate" → search "pivot" → close → 2 sec silence

### [END CLIP]

**Gotcha:** Make sure you are running `LaunchUTLCommandCenter` (the universal version), NOT `LaunchCommandCenter` from the demo file's modFormBuilder.

---

## CLIP 37 — Python: File Comparison Script
**Sequential Number:** Clip 37 of 39
**Duration:** ~60 seconds
**Macro/Action:** compare_files.py (Python script — run from Command Prompt)
**Audio Clip:** V3_S11

### Narration Script (exact words):

> "Beyond VBA, the library includes 22 Python scripts for heavier-duty work. You don't need to be a programmer — each one has a step-by-step guide.
>
> Here's the File Comparator. You point it at two Excel files and it builds a color-coded diff report — every added row, every removed row, every changed cell.
>
> [run the script]
>
> Green means added. Red means removed. Yellow means changed. One command, complete report."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Bring up the **Command Prompt** window (click it on the taskbar, or Alt+Tab)
3. Play the audio
4. Type (or paste) the command — have this pre-typed in Notepad so you can just paste it:
   ```
   python compare_files.py "C:\Users\Connor\Desktop\Budget_v1.xlsx" "C:\Users\Connor\Desktop\Budget_v2.xlsx"
   ```
   (Adjust the paths to match your actual file locations and the path to compare_files.py)
5. Press **Enter**
6. Wait 3–5 seconds — the terminal shows progress messages
7. Output: "COMPARISON_REPORT.xlsx saved to ..."
8. Switch to Excel (**Alt+Tab**)
9. Open the COMPARISON_REPORT.xlsx from your Desktop (File → Open → navigate → open)
10. Show the SUMMARY sheet: Added/Removed/Changed counts
11. Click into a detail sheet showing cell-by-cell differences with color coding
12. **Pause 3 seconds**
13. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → Command Prompt → run compare_files.py → output → switch to Excel → open report → show summary → show diff detail → PAUSE 3 sec → 2 sec silence

### [END CLIP]

**Gotcha:** If Python is not in your PATH, use the full path to python.exe. Test the exact command beforehand.

---

## CLIP 38 — Python: PDF Extractor Script
**Sequential Number:** Clip 38 of 39
**Duration:** ~60 seconds
**Macro/Action:** pdf_extractor.py (Python script — run from Command Prompt)
**Audio Clip:** V3_S12

### Narration Script (exact words):

> "This one is especially powerful — the PDF Table Extractor. Point it at any PDF with data tables and it pulls them into Excel. No retyping, no copy-paste from a PDF viewer.
>
> [run the script]
>
> Every table gets its own sheet — Page 1 Table 1, Page 2 Table 1 — headers styled, data aligned. A PDF table is now an Excel table."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Bring up the **Command Prompt** window
3. Play the audio
4. Type (or paste) the command:
   ```
   python pdf_extractor.py "C:\Users\Connor\Desktop\SampleReport.pdf"
   ```
   (Adjust paths as needed)
5. Press **Enter**
6. Wait 3–10 seconds (depends on PDF size)
7. Terminal shows: "Found X tables across Y pages" and "Saved to PDF_EXTRACTED_TABLES.xlsx"
8. Switch to Excel
9. Open PDF_EXTRACTED_TABLES.xlsx from Desktop
10. Show the extracted tables — each table on its own sheet, headers styled, data aligned
11. **Pause 3 seconds**
12. Wait 2 seconds, stop recording

### [START CLIP]

- 2 sec silence → Command Prompt → run pdf_extractor.py → output → switch to Excel → open extracted tables → show results → PAUSE 3 sec → 2 sec silence

### [END CLIP]

**Gotcha:** The PDF must have selectable text — scanned image PDFs will not work. Test beforehand.

---

## CLIP 39 — Closing + Recap
**Sequential Number:** Clip 39 of 39
**Duration:** ~45 seconds
**Macro/Action:** None (narration only)
**Audio Clip:** V3_S13_Closing.mp3

### Narration Script (exact words):

> "That's a sample of what's in the universal tools library. We showed maybe ten percent of what's available.
>
> Here's how to get started:
>
> Everything is on SharePoint — the VBA code, the Python scripts, the SQL tools, and the sample files. Each tool has a step-by-step guide written for someone who's never touched code before. And if you need help, there's a master document with pre-built Copilot prompts that will walk you through using any tool, step by step.
>
> Start with the guide. Pick one tool that solves a problem you deal with every week. Try it. And if you have questions or ideas for new tools, reach out — I'm happy to help.
>
> Thanks for watching."

### Screen Actions — Step by Step:

1. Start OBS recording, wait 2 seconds
2. Switch back to the sample file in Excel
3. Play the closing audio
4. While the audio plays, **slowly scroll through the tab bar** at the bottom showing all the output sheets created during the demo:
   - UTL_Sanitizer_Preview
   - UTL_CommentReport
   - UTL_SheetIndex
   - UTL_CompareReport
   - UTL_Consolidated
   - (any others)
5. When the audio says "Thanks for watching" — hold still for 3 seconds
6. Wait 2 seconds of silence
7. Stop OBS recording

### [START CLIP]

- 2 sec silence → audio plays → scroll through tab bar showing all output sheets → "Thanks for watching" → hold 3 sec → 2 sec silence

### [END CLIP]

**Note:** The closing title card (branded slide with SharePoint link and tool counts) gets added in the video editor.

---

### VIDEO 3 COMPLETE

**Congratulations — ALL 39 CLIPS ARE RECORDED.**

---
---

# EMERGENCY RECOVERY — If Something Goes Wrong

## Macro errors out with a VBA error dialog
- Click **End** (NOT Debug)
- The action failed but Excel is fine
- Re-run the action. If it fails again, skip it and note it for re-recording later

## Excel freezes or hangs
- Wait 30 seconds — some macros are just slow
- If still frozen after 60 seconds, press **Ctrl+Break** to interrupt VBA
- If totally unresponsive, open Task Manager (Ctrl+Shift+Esc), find Excel, click **End Task**, and restart from the last saved state

## Wrong sheet is showing
- Just click the correct tab. You can trim the bad part out in the video editor.

## Command Center doesn't open (Ctrl+Shift+M not working)
- Use **Alt+F8** → type `LaunchCommandCenter` → **Run** instead
- To set up the shortcut for future use: Alt+F8 → type `SetupKeyboardShortcuts` → Run

## What-If won't restore baseline
1. Run Action 65 (Restore Baseline) manually
2. If that fails: right-click any tab → Unhide → find "WhatIf_Baseline" → unhide it → manually copy values back to Assumptions
3. Nuclear option: close without saving and reopen from the last saved version

## PDF export fails
- Check that a PDF printer is configured (Microsoft Print to PDF)
- Try exporting a single sheet manually: File → Export → Create PDF

## A notification pops up during recording
- Stop recording immediately
- Dismiss the notification
- Re-do Computer Lockdown (check Focus Assist, close the offending app)
- Re-record that clip from scratch

## One audio clip doesn't match the screen timing
- This is normal and expected. You will adjust timing in the video editor by sliding the audio track left or right on the timeline. Small 0.5–2 second adjustments are typical.

---
---

# QUICK REFERENCE — Action Numbers

Use this table to quickly find the Command Center action number for each macro used in the recordings.

| Action # | Action Name | VBA Module | Used In Clip(s) |
|----------|-------------|------------|-----------------|
| 3 | Reconciliation Checks | modReconciliation | Clip 13 |
| 6 | Variance Analysis (MoM) | modVarianceAnalysis | Clip 14 |
| 7 | Data Quality Scan | modDataQuality | Clips 4, 12 |
| 10 | PDF Export | modPDFExport | Clip 19 |
| 12 | Build Dashboard | modDashboard | Clips 6, 17 |
| 17 | Import Data Pipeline | modImport | Clip 11 |
| 32 | Save Version | modVersionControl | Clip 22 |
| 35 | List Versions | modVersionControl | (optional in Clip 22) |
| 41 | View Audit Log | modLogger | Clip 25 |
| 44 | Run Full Integration Test | modIntegrationTest | Clip 24 |
| 46 | Generate Variance Commentary | modVarianceAnalysis | Clips 5, 15 |
| 48 | Toggle Executive Mode | modNavigation | Clip 21 |
| 63 | Run What-If Demo | modWhatIf | Clip 23 |
| 65 | Restore Baseline | modWhatIf | Clip 23 |
| — | GenerateExecBrief | modExecBrief | Clip 20 |
| — | ShowTimeSavedReport | modTimeSaved | Clip 26 |

**Universal Toolkit macros (Video 3 — run via Alt+F8):**

| Macro Name | Module | Used In Clip |
|------------|--------|-------------|
| PreviewSanitizeChanges | modUTL_DataSanitizer | Clip 27 |
| RunFullSanitize | modUTL_DataSanitizer | Clip 27 |
| HighlightByThreshold | modUTL_Highlights | Clip 28 |
| HighlightDuplicateValues | modUTL_Highlights | Clip 28 |
| ClearHighlights | modUTL_Highlights | Clip 28 |
| CountComments | modUTL_Comments | Clip 29 |
| ExtractAllComments | modUTL_Comments | Clip 29 |
| ColorTabsByKeyword | modUTL_TabOrganizer | Clip 30 |
| ReorderTabs | modUTL_TabOrganizer | Clip 30 |
| SplitColumn | modUTL_ColumnOps | Clip 31 |
| CombineColumns | modUTL_ColumnOps | Clip 31 |
| ListAllSheetsWithLinks | modUTL_SheetTools | Clip 32 |
| TemplateCloner | modUTL_SheetTools | Clip 32 |
| CompareSheets | modUTL_Compare | Clip 33 |
| ConsolidateSheets | modUTL_Consolidate | Clip 34 |
| BuildVLOOKUP | modUTL_LookupBuilder | Clip 35 |
| CreateDropdownList | modUTL_ValidationBuilder | Clip 35 |
| ListAllPivots | modUTL_PivotTools | Clip 35 |
| LaunchUTLCommandCenter | modUTL_CommandCenter | Clip 36 |

**Python scripts (Video 3 — run from Command Prompt):**

| Script | Used In Clip |
|--------|-------------|
| compare_files.py | Clip 37 |
| pdf_extractor.py | Clip 38 |

---
---

# CLIP MASTER INDEX

All 39 clips at a glance:

| Clip | Video | Section | Duration | What Happens |
|------|-------|---------|----------|-------------|
| 1 | V1 | Title Card | 5 sec | Static branded title (add in editor) |
| 2 | V1 | Opening Hook | 30 sec | Scroll landing page |
| 3 | V1 | Command Center | 40 sec | Open CC, browse, search "variance" |
| 4 | V1 | Data Quality Scan | 40 sec | Run Action 7, show letter grade |
| 5 | V1 | Variance Commentary | 45 sec | Run Action 46, jaw-drop narratives |
| 6 | V1 | Executive Dashboard | 40 sec | Run Action 12, scroll charts |
| 7 | V1 | Bridge + Closing | 60 sec | Narration over landing page |
| 8 | V2 | Splash Screen | 15 sec | Open file, splash fires |
| 9 | V2 | Opening + Tour | 85 sec | Intro narration + click through tabs |
| 10 | V2 | Command Center | 45 sec | Open CC, scroll, search |
| 11 | V2 | GL Import | 45 sec | Run Action 17, show data |
| 12 | V2 | Data Quality | 50 sec | Run Action 7, letter grade + breakdown |
| 13 | V2 | Reconciliation | 45 sec | Run Action 3, PASS/FAIL checks |
| 14 | V2 | Variance Analysis | 40 sec | Run Action 6, highlighted variances |
| 15 | V2 | Variance Commentary | 45 sec | Run Action 46, JAW-DROP |
| 16 | V2 | YoY Variance | 30 sec | Run YoY action |
| 17 | V2 | Dashboard Charts | 45 sec | Run Action 12, 8 charts |
| 18 | V2 | Executive Dashboard | 30 sec | KPI cards, waterfall |
| 19 | V2 | PDF Export | 30 sec | Run Action 10, show PDF |
| 20 | V2 | Executive Brief | 40 sec | Run ExecBrief, 5 sections |
| 21 | V2 | Executive Mode | 20 sec | Toggle tabs on/off |
| 22 | V2 | Version Control | 30 sec | Save snapshot "March Close Draft 1" |
| 23 | V2 | What-If Scenario | 90 sec | Run scenario + restore baseline |
| 24 | V2 | Integration Test | 30 sec | 18/18 PASS |
| 25 | V2 | Audit Log | 25 sec | Show full audit trail |
| 26 | V2 | Time Saved + Closing | 45 sec | Annual savings + closing narration |
| 27 | V3 | Data Sanitizer | 60 sec | Preview + full sanitize |
| 28 | V3 | Highlights | 60 sec | Threshold + duplicates |
| 29 | V3 | Comments | 40 sec | Count + extract |
| 30 | V3 | Tab Organizer | 50 sec | Color + reorder tabs |
| 31 | V3 | Column Ops | 50 sec | Split + combine columns |
| 32 | V3 | Sheet Tools | 50 sec | Sheet index + template cloner |
| 33 | V3 | Compare Sheets | 50 sec | Cell-by-cell diff report |
| 34 | V3 | Consolidate | 40 sec | Stack sheets with source tracking |
| 35 | V3 | Power Tools | 60 sec | VLOOKUP + Validation + Pivots |
| 36 | V3 | Universal CC | 50 sec | Full toolkit Command Center |
| 37 | V3 | Python: Compare | 60 sec | compare_files.py |
| 38 | V3 | Python: PDF | 60 sec | pdf_extractor.py |
| 39 | V3 | Closing | 45 sec | Recap + CTA |

---

*Created: 2026-03-19 | Single-Day Recording Playbook for iPipeline Finance Automation Videos*
*Source: FinalExport/VideoRecording scripts + RECORDING_INSTRUCTIONS.md + Video Production Guide*
