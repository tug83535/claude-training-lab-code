# Master Director Macro — Setup & Execution Guide

**What This Is:** Step-by-step instructions for using the `modDirector.bas` VBA module to automate your entire video demo recording. This module acts as a puppeteer — it plays your AI audio clips, navigates sheets, triggers macros, scrolls, and pauses, all timed to your script. You press one button, start OBS, and record.

---

## Table of Contents

1. [What You Need Before Starting](#1-what-you-need-before-starting)
2. [Import the Director Module](#2-import-the-director-module)
3. [Set Your Audio File Path](#3-set-your-audio-file-path)
4. [Configure OBS Studio](#4-configure-obs-studio)
5. [Run the Quick Test](#5-run-the-quick-test)
6. [Record Video 1](#6-record-video-1)
7. [Record Video 2](#7-record-video-2)
8. [Record Video 3](#8-record-video-3)
9. [Testing Individual Clips](#9-testing-individual-clips)
10. [Adjusting Timing](#10-adjusting-timing)
11. [Troubleshooting](#11-troubleshooting)

---

## 1. What You Need Before Starting

Before you do anything, confirm you have all of these:

- [ ] **Demo Excel file** (`iPipeline_PnL_Demo.xlsm`) with all 39 VBA modules imported and working
- [ ] **Sample file** (`Sample_Quarterly_Report.xlsm`) with universal toolkit modules imported (for Video 3)
- [ ] **Audio clips** — all 41 MP3 files in `FinalExport/AudioClips/` organized in `Video1/`, `Video2/`, `Video3/` subfolders
- [ ] **OBS Studio** installed and working
- [ ] **The Director module** (`modDirector.bas`) — the file you are about to import
- [ ] **Macros enabled** in Excel (File > Options > Trust Center > Trust Center Settings > Enable all macros + Trust access to VBA project object model)
- [ ] **frmCommandCenter** UserForm built in the demo file (test: Ctrl+Shift+M should open the styled Command Center form)

---

## 2. Import the Director Module

### Step-by-step:

1. Open your demo Excel file (`iPipeline_PnL_Demo.xlsm`)
2. Press **Alt+F11** to open the VBA Editor
3. In the VBA Editor, click **File** > **Import File...**
4. Navigate to `FinalExport/DemoVBA/`
5. Select **`modDirector.bas`**
6. Click **Open**
7. You should now see `modDirector` in the Modules folder in the left panel
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

## 3. Set Your Audio File Path

The Director module needs to know where your audio clips are stored. **You must set this path before running anything.**

### Step-by-step:

1. Press **Alt+F11** to open the VBA Editor
2. In the left panel, double-click **modDirector** to open it
3. Near the top of the file (~line 65), find this line:

```vba
Private Const AUDIO_BASE_PATH As String = "C:\Users\connor.atlee\.claude\projects\claude-training-lab-code\FinalExport\AudioClips\"
```

4. Change the path to wherever YOUR AudioClips folder is. For example, if you copied the FinalExport folder to your Desktop:

```vba
Private Const AUDIO_BASE_PATH As String = "C:\Users\connor.atlee\Desktop\FinalExport\AudioClips\"
```

5. **The path MUST end with a backslash** (`\`)
6. **The folder MUST contain** `Video1/`, `Video2/`, and `Video3/` subfolders with the MP3 files
7. Press **Ctrl+S** to save
8. Press **Alt+Q** to close the VBA Editor

### Verify the path is correct:

Open File Explorer and navigate to your AudioClips folder. Confirm you see:
```
AudioClips\
├── Video1\
│   ├── V1_S1_Opening_Hook.mp3
│   ├── V1_S2_Command_Center.mp3
│   └── ... (7 files)
├── Video2\
│   ├── V2_S0_Opening.mp3
│   └── ... (21 files)
└── Video3\
    ├── V3_S0_Opening.mp3
    └── ... (13 files)
```

---

## 4. Configure OBS Studio

### OBS Settings for Director Macro Recording:

1. Open OBS Studio
2. Click **Settings** (bottom right)

**Output tab:**
- Recording Path: Choose a folder on your Desktop (e.g., `Desktop/VideoRecordings/`)
- Recording Format: **mp4**
- Recording Quality: **High Quality, Medium File Size**

**Video tab:**
- Base Resolution: **1920x1080**
- Output Resolution: **1920x1080**
- FPS: **30**

**Audio tab (IMPORTANT):**
- **Desktop Audio: ENABLED** — This captures the MP3 audio that the Director macro plays
- **Mic/Auxiliary Audio: DISABLED** — You do not want your microphone recorded
- If you do NOT want audio baked into the recording (adding in post instead), disable Desktop Audio too

3. Click **OK** to save settings

**Add a Display Capture source:**
1. In the Sources panel, click **+** > **Display Capture**
2. Name it "Excel Recording"
3. Select your main display
4. Click **OK**

**Test recording:**
1. Click **Start Recording** in OBS
2. Wait 5 seconds
3. Click **Stop Recording**
4. Find the file and play it back — verify it's clean 1080p

---

## 5. Run the Pre-Flight Check and Quick Test

Before recording anything for real, run **two checks** to verify everything works.

### Step A: Run the Pre-Flight Check

1. Press **Alt+F8** to open the Macro dialog
2. Select **`RunPreflight`**
3. Click **Run**
4. A report will appear showing:
   - Whether all audio files exist for all 3 videos
   - Whether the MCI audio subsystem works (and shows test clip duration)
   - Whether Excel is maximized, zoom is 100%, correct sheet is active
   - Whether all required sheets exist
5. **Everything should say OK, FOUND, or YES.** If anything says MISSING or FAILED, fix it before recording.

### Step B: Run the Quick Test

1. Make sure you are in the demo Excel file
2. Navigate to the **Report-->** sheet
3. Press **Alt+F8** to open the Macro dialog
4. Select **`QuickTest`**
5. Click **Run**

### What should happen:

- The pre-flight check runs automatically first
- You should **hear audio** playing (the opening hook clip)
- The screen should **scroll down smoothly** on the Report--> page
- A message box appears showing: audio status, measured clip duration, and pre-flight result

### If audio did NOT play:

- Check that `AUDIO_BASE_PATH` is correct (Step 3 above)
- Check that the MP3 files actually exist in the Video1 subfolder
- Try playing one of the MP3 files manually (double-click it) to verify it's not corrupted
- Make sure your computer's audio is not muted

### If scrolling did NOT work:

- Make sure the Report--> sheet is the active sheet
- Make sure Excel is not in Edit mode (press Escape first)

---

## 6. Record Video 1

Video 1 is "What's Possible" — approximately 5 minutes, 7 clips.

### Pre-recording checklist:

> **v2.0 Note:** The macro now runs an automatic pre-flight check before each video. It verifies maximized window, 100% zoom, correct sheet, A1 selected, and audio files. If anything is wrong, it tells you AND auto-fixes what it can. But you should still handle these yourself:

- [ ] All notifications are OFF (Windows Focus Assist > Alarms Only)
- [ ] Desktop is clean (no icons, taskbar auto-hidden)
- [ ] OBS is open and ready to record
- [ ] Command Center is **CLOSED**

### Step-by-step:

1. Complete the pre-recording checklist above
2. **Start OBS recording** (click Start Recording or use your hotkey)
3. **Wait 3 seconds** (so OBS captures a clean start)
4. In Excel, press **Alt+F8**
5. Select **`RunVideo1`**
6. Click **Run**
7. **Do NOT touch the mouse or keyboard** — the macro is driving everything
8. Watch the screen — you'll see:
   - Slow scroll on the landing page (with audio playing)
   - Command Center opening, categories scrolling, "variance" being typed
   - Data Quality Scan running, letter grade appearing
   - Variance Commentary generating (the jaw-drop moment)
   - Executive Dashboard building with charts
   - Return to landing page for closing
9. When the "Video 1 recording complete!" message appears:
   - Click **OK** to dismiss it
   - **Stop OBS recording**
10. Play back the recording to check quality

### If something goes wrong:

- If a macro errors out, click **End** (NOT Debug)
- Delete any output sheets that were created
- Navigate back to Report--> sheet
- Re-run `CleanupAllOutputSheets`
- Re-run `RunVideo1`

---

## 7. Record Video 2

Video 2 is "Full Demo Walkthrough" — approximately 18 minutes, 19 clips.

### Before starting:

1. Run `CleanupAllOutputSheets` to reset the file
2. Navigate to **Report-->** sheet
3. Select cell **A1**
4. Save the file

### Step-by-step:

1. **Start OBS recording**
2. Wait 3 seconds
3. Press **Alt+F8** > Select **`RunVideo2`** > Click **Run**
4. **Do NOT touch anything** — this one takes ~18 minutes
5. The macro will:
   - Tour the workbook (clicking through sheet tabs)
   - Open the Command Center and search it
   - Run Data Quality Scan, Reconciliation, Variance Analysis
   - Generate Variance Commentary (jaw-drop moment #2)
   - Build Dashboard Charts and Executive Dashboard
   - Run PDF Export, Executive Brief, Executive Mode toggle
   - Save a version, run What-If scenario with restore
   - Run Sensitivity Analysis, Integration Test
   - Show Audit Log and Time Saved Calculator
   - Return to Report--> for closing
6. When the completion message appears:
   - Click **OK**
   - **Stop OBS recording**

### Special notes for Video 2:

- **GL Import (Clip 11):** The macro skips the actual file dialog and just shows the GL sheet with existing data. This avoids a file picker appearing on camera.
- **PDF Export (Clip 19):** v2.0 bypasses the file dialog entirely — exports directly to your Desktop as `KBT_Report_Package_YYYYMMDD.pdf`. No dialog, no SendKeys, no risk.
- **Version Control (Clip 22):** v2.0 bypasses the InputBox entirely — saves directly as `versions/v1_[timestamp]_March_Draft_1.xlsx`. No dialog, no SendKeys.
- **What-If (Clip 23):** v2.0 bypasses both the InputBox AND the restore confirmation MsgBox entirely — calls `RunWhatIfPreset(1)` and `RestoreBaselineSilent` directly. No dialog, no SendKeys, zero risk. This is the "wow moment" clip so it must be bulletproof.

---

## 8. Record Video 3

Video 3 is "Universal Tools" — approximately 10 minutes, 12 clips.

### IMPORTANT: Video 3 runs on the SAMPLE file, not the demo file.

### Before starting:

1. **Close the demo file** (or keep it in the background)
2. **Open** `Sample_Quarterly_Report.xlsm`
3. Verify the sample file has:
   - Messy data (blank rows, text-stored numbers, merged cells, etc.)
   - Universal toolkit VBA modules imported (modUTL_*.bas files)
4. **Import modDirector.bas into the sample file too** (repeat Step 2 from this guide)
5. Set the `AUDIO_BASE_PATH` in the sample file's copy of modDirector (same as Step 3)

### Step-by-step:

1. Make sure you are in the **sample file** (not the demo file)
2. Navigate to the first data sheet
3. **Start OBS recording**
4. Wait 3 seconds
5. Press **Alt+F8** > Select **`RunVideo3`** > Click **Run**
6. The macro will warn you if it detects you're on the demo file
7. Watch the Universal Tools demo play through
8. When the completion message appears:
   - Click **OK**
   - **Stop OBS recording**

### Special notes for Video 3:

- Some universal toolkit macros prompt for input (which sheet to compare, what threshold, etc.). The macro uses `SendKeys` to pre-stage answers. If a dialog appears and doesn't auto-close, press Enter.
- Video 3 macros may create extra sheets (comparison reports, extracted comments, etc.). These are part of the demo.

---

## 9. Testing Individual Clips

You can test any single clip without running the full video. This is useful for:
- Checking timing on a specific clip
- Re-recording just one clip
- Debugging an issue

### How to test a single clip:

1. Press **Alt+F8**
2. Select **`TestClip`**
3. Click **Run**
4. A dialog will NOT appear — instead, you call it from the Immediate Window:
   - Press **Alt+F11** to open VBA Editor
   - Press **Ctrl+G** to open the Immediate Window
   - Type: `TestClip 4` (replace 4 with your clip number)
   - Press **Enter**

### Clip number reference:

| Clip # | Video | What It Does |
|--------|-------|-------------|
| 1 | V1 | Title Card (5 sec pause) |
| 2 | V1 | Opening Hook — scroll landing page |
| 3 | V1 | Command Center — open, search, close |
| 4 | V1 | Data Quality Scan — run + show report |
| 5 | V1 | Variance Commentary — jaw-drop feature |
| 6 | V1 | Executive Dashboard — charts + KPIs |
| 7 | V1 | Bridge + Closing — static landing page |
| 8 | V2 | Opening — scroll landing page |
| 9 | V2 | Workbook Tour — click through tabs |
| 10 | V2 | Command Center — search "reconciliation" |
| 11 | V2 | GL Import — show General Ledger data |
| 12 | V2 | Data Quality Scan |
| 13 | V2 | Reconciliation Checks — PASS/FAIL |
| 14 | V2 | Variance Analysis — flagged items |
| 15 | V2 | Variance Commentary — jaw-drop #2 |
| 16 | V2 | YoY Variance Analysis |
| 17 | V2 | Dashboard Charts — 8-chart grid |
| 18 | V2 | Executive Dashboard — KPIs + waterfall |
| 19 | V2 | PDF Export |
| 20 | V2 | Executive Brief |
| 21 | V2 | Executive Mode — toggle on/off |
| 22 | V2 | Version Control — save snapshot |
| 23 | V2 | What-If Scenario — THE WOW MOMENT |
| 24 | V2 | Sensitivity Analysis |
| 25 | V2 | Integration Test — 18/18 PASS |
| 26 | V2 | Audit Log + Time Saved + Closing |
| 27 | V3 | Opening — show messy sample file |
| 28 | V3 | Data Sanitizer — preview + clean |
| 29 | V3 | Highlights — threshold + duplicates |
| 30 | V3 | Comments — count + extract |
| 31 | V3 | Tab Organizer — color + reorder |
| 32 | V3 | Column Ops — split + combine |
| 33 | V3 | Sheet Tools — index + clone |
| 34 | V3 | Compare Sheets — cell-by-cell diff |
| 35 | V3 | Consolidate Sheets |
| 36 | V3 | Pivot Tools + Lookup/Validation |
| 37 | V3 | Universal Command Center |
| 38 | V3 | Closing |

---

## 10. Adjusting Timing

**v2.0: Audio clip durations are now measured automatically at runtime.** The macro reads each MP3 file's actual length before playing it, so the timing always matches your audio. You do NOT need to manually measure clip lengths or update duration constants — that's all automatic.

### What you CAN still adjust:

- **Scroll speed:** Change `SCROLL_STEP_DELAY` (default 250ms between scroll steps)
- **Typing speed:** Change `TYPING_DELAY_MS` (default 90ms between characters)
- **Silence padding:** Change `SILENCE_PAD_SEC` (default 2 seconds at start/end of each clip)

### Within individual clips:

Each clip sub has `WaitSec` calls that control specific pauses. For example, in `V1_Clip5_VarianceCommentary`:

```vba
' JAW-DROP MOMENT: Pause 3 seconds in silence
WaitSec 3
```

Change the `3` to `5` if you want a longer pause on the narratives.

---

## 11. Troubleshooting

### "Audio file not found" in the Immediate Window

- Check that `AUDIO_BASE_PATH` is set correctly
- Check that the subfolder names are exactly `Video1`, `Video2`, `Video3`
- Check that the MP3 filenames match exactly (case-sensitive on some systems)

### No audio plays but no error

- Check that your computer audio is not muted
- Check Windows volume mixer — Excel (or VBA) should have audio
- Try playing the MP3 file manually by double-clicking it
- **v2.0 fix:** The macro now auto-resets the MCI audio device at every entry point, so "stuck state" from interrupted runs is automatically cleared. If audio still fails, restart Excel.

### Macro errors out during a clip

- Click **End** (NOT Debug) on the error dialog
- Run `CleanupAllOutputSheets` to reset the file
- Navigate back to Report--> sheet
- Try the individual clip with `TestClip N` to isolate the issue
- If a specific macro crashes, that macro has a bug independent of the Director

### Command Center doesn't appear on screen

- The Director tries to show frmCommandCenter modeless (non-blocking)
- If the UserForm doesn't exist, it silently skips and calls the macro directly
- To fix: Ensure frmCommandCenter is built in the workbook (run `BuildCommandCenter` once)

### Screen looks jumpy or too fast

- Increase `SCROLL_STEP_DELAY` to 400 or 500 (slower scrolling)
- Increase `TYPING_DELAY_MS` to 120 or 150 (slower typing)
- Add more `WaitSec` calls in the clip subs where you want longer pauses

### SendKeys doesn't type into a dialog

- `SendKeys` can be unreliable with modal dialogs
- If a macro's InputBox appears and doesn't get the pre-staged keys, just type the answer manually and press Enter
- The Director will continue after the dialog closes

### OBS recording has no audio

- In OBS Settings > Audio, make sure **Desktop Audio** is enabled
- The Director plays audio through Windows (mciSendString), which counts as desktop audio
- Test: Play any MP3 file manually while OBS is recording, then check the recording

### I want to re-record just one clip

- Run `CleanupAllOutputSheets` first
- Get the file back to the right starting state for that clip
- Start OBS recording
- Run `TestClip N` from the Immediate Window
- Stop OBS recording

---

## Quick Reference: Running the Macros

| What You Want | Macro to Run | How to Run |
|---|---|---|
| Test audio + scrolling | `QuickTest` | Alt+F8 > QuickTest > Run |
| Record Video 1 (~5 min) | `RunVideo1` | Alt+F8 > RunVideo1 > Run |
| Record Video 2 (~18 min) | `RunVideo2` | Alt+F8 > RunVideo2 > Run |
| Record Video 3 (~10 min) | `RunVideo3` | Alt+F8 > RunVideo3 > Run (from sample file) |
| Record all 3 videos | `RunAllVideos` | Alt+F8 > RunAllVideos > Run |
| Test one specific clip | `TestClip N` | VBA Immediate Window: `TestClip 4` |
| Clean up after recording | `CleanupAllOutputSheets` | Alt+F8 > CleanupAllOutputSheets > Run |

---

## Recording Day Checklist

### Morning — Video 1 + Video 2
1. [ ] Computer lockdown (notifications off, desktop clean, taskbar hidden)
2. [ ] Demo file open, macros enabled, on Report--> sheet
3. [ ] Run `CleanupAllOutputSheets`
4. [ ] Run `QuickTest` — confirm audio and scrolling work
5. [ ] Start OBS recording
6. [ ] Run `RunVideo1` — watch it play through (~5 min)
7. [ ] Stop OBS when completion message appears
8. [ ] Run `CleanupAllOutputSheets` to reset
9. [ ] Start OBS recording
10. [ ] Run `RunVideo2` — watch it play through (~18 min)
11. [ ] Stop OBS when completion message appears

### Afternoon — Video 3
12. [ ] Switch to `Sample_Quarterly_Report.xlsm`
13. [ ] Verify modDirector is imported and audio path is set
14. [ ] Start OBS recording
15. [ ] Run `RunVideo3` — watch it play through (~10 min)
16. [ ] Stop OBS when completion message appears
17. [ ] Review all recordings

---

*Guide created: 2026-03-24 | Part of the Master Director Macro package*
