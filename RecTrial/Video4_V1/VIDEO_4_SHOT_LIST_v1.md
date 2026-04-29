# Video 4 — Shot List & Screen Recording Guide
## Python Automation for Finance — iPipeline

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Version:** v1.0 — 2026-04-28
**Companion doc:** VIDEO_4_NARRATION_SCRIPT_v1.md (same folder)

---

## BEFORE YOU RECORD — ONE-TIME SETUP CHECKLIST

- [ ] Run `smoke_test_video4_python.py` — confirm 5/5 PASS
- [ ] Open FinanceTools.xlsm in Excel — confirm Finance Tools button visible
- [ ] Set screen resolution to 1920x1080 (or your standard OBS resolution)
- [ ] Set Windows zoom to 100% — no scaling
- [ ] Open File Explorer and navigate to the ZeroInstall/outputs/ folder — have it ready
- [ ] Close all other apps — clean taskbar, no notification popups
- [ ] Disable Windows notifications (Focus Assist ON) for the recording session
- [ ] Have Notepad open and minimized — you'll need it for Chapter 4
- [ ] Have Excel open and minimized — you'll need it for Chapters 5 and 3
- [ ] Have Chrome/Edge open and minimized — for HTML report in Chapter 3

---

## CLIP-BY-CLIP SHOT LIST

---

### V4_C01 — Why Python After Excel and VBA? (~44 sec)

**What's on screen:** Static title card or simple two-column slide

| # | Action | Notes |
|---|---|---|
| 1 | Show a clean title card or static slide | Two columns: "Excel + VBA is for..." / "Python adds..." — keep it simple, no animation needed |
| 2 | Hold for full clip duration (~44 sec) | No mouse movement — just let the narration play over the static image |

**Tip:** If you don't have a title card ready, a plain white slide with two bullet columns in Arial works fine. This chapter has no demo — just set and forget.

---

### V4_C02 — Safety First (~60 sec)

**What's on screen:** PYTHON_SAFETY.md in Notepad → outputs folder path

| # | Action | Notes |
|---|---|---|
| 1 | Open PYTHON_SAFETY.md in Notepad | File is at `RecTrial\UniversalToolkit\python\PYTHON_SAFETY.md` |
| 2 | Slow scroll through the 14 rules | Scroll at reading pace — don't rush. Show at least rules 1–5 clearly |
| 3 | At narration "Each run creates a new timestamped folder..." | Switch to File Explorer showing the outputs/ folder |
| 4 | Show a timestamped folder name (e.g., `20260428_143212_revenue_leakage_finder`) | If outputs/ is empty, run a sample first to create one |
| 5 | Hold on the folder view until clip ends | |

**Tip:** Font size in Notepad — bump it to 14pt before recording so the rules are legible on screen.

---

### V4_C03a — Revenue Leakage Finder: Setup (~55 sec)

**What's on screen:** Excel with Finance Tools button → CLI menu → option 1 selected → processing

| # | Action | Notes |
|---|---|---|
| 1 | Show FinanceTools.xlsm open in Excel | Finance Tools button should be visible and prominent |
| 2 | At narration "Everything runs from this one button" | Hover mouse over the button briefly — don't click yet |
| 3 | At narration "which customers are being billed without a matching contract..." | Click the Finance Tools button — CLI menu window opens |
| 4 | Show the full numbered menu for ~3 seconds | Let viewer read the options |
| 5 | Type "1" and press Enter | Slow and deliberate — viewer should see the keypress |
| 6 | Show the script processing: a few output lines scrolling | Don't cut away immediately — let ~3–4 lines print |
| 7 | Show "Analysis complete." and the output folder path | End of C03a — pause here before starting C03b screen action |

**Tip:** The CLI window will close after the script finishes. Have the output folder open in Explorer in the background so you can switch instantly for C03b.

---

### V4_C03b — Revenue Leakage Finder: Results (~1 min 53 sec)

**What's on screen:** HTML report in browser → ARR waterfall → exceptions_ranked.csv in Excel

| # | Action | Notes |
|---|---|---|
| 1 | Switch to browser — HTML report should auto-open, or open it manually from outputs/ | Navigate to: `outputs/[latest_folder]/leakage_report.html` |
| 2 | Show the headline summary section for ~5 seconds | Exception counts by class, total expected vs billed |
| 3 | Scroll slowly down to the exception detail table | Pause on 2–3 specific rows — point out a clear leakage case |
| 4 | At narration "This is the ARR waterfall..." | Scroll to the waterfall chart section |
| 5 | Hold on waterfall for ~15–20 seconds | This is the closing visual artifact — give it time |
| 6 | At narration "Down at the row level..." | Switch to Excel — open exceptions_ranked.csv |
| 7 | Show the ranked rows in Excel — Column A (rank), customer name, class, priority score | Widen columns so text is readable |
| 8 | At narration "To run this against your own data..." | Hold on the CSV view until clip ends |

**[FLAG — stdlib waterfall]:** The waterfall shown is a CSS bar chart in the HTML report.
If matplotlib is approved later, the visual here will look different — re-record this section
only if the visual changes materially.

**Tip:** Before recording, run `revenue_leakage_finder.py --sample` so the output folder
and HTML report are already there. Don't generate them live during recording — too risky.

---

### V4_C04 — Data Contract Checker (~1 min 26 sec)

**What's on screen:** CLI menu → FAIL output → Notepad edit → PASS output

| # | Action | Notes |
|---|---|---|
| 1 | Return to CLI menu (reopen FinanceTools.xlsm and click button, or show menu already open) | |
| 2 | Type "2" and press Enter — Data Contract Checker runs in sample mode | |
| 3 | Show red FAIL output for ~5 seconds | The script intentionally produces a bad-file scenario in sample mode |
| 4 | At narration "In this example, a required column is missing..." | Highlight or zoom slightly on the specific error lines |
| 5 | Switch to Notepad — open the billing CSV (or a simple example file) | |
| 6 | Rename one column header — e.g., change `amount_billed` to `amount_billed_fixed` then back | Keep the edit fast and visible |
| 7 | Save the file in Notepad (Ctrl+S) | |
| 8 | Re-run the tool (return to menu, type "2" again) | |
| 9 | Show green PASS output | Hold for ~3 seconds |
| 10 | At narration "PASS means the file is safe to analyze..." | Hold on PASS output until clip ends |

**Tip:** Pre-stage the "bad" CSV file at a known path before recording. Don't improvise
the edit live — know exactly which column you're renaming.

---

### V4_C05 — Exception Triage Engine (~1 min 26 sec)

**What's on screen:** CLI menu → scored terminal output → top_10_action_list.csv in Excel

| # | Action | Notes |
|---|---|---|
| 1 | Return to CLI menu | |
| 2 | Type "3" and press Enter — Exception Triage Engine runs in sample mode | It auto-finds the most recent Revenue Leakage output |
| 3 | Show terminal output: exception class labels, customer names, scores appearing | Let ~6–8 lines print before cutting away |
| 4 | At narration "Each exception gets a priority score..." | Zoom in slightly on one scored row if possible |
| 5 | At narration "Row one is your highest-priority review..." | Switch to Excel — open `top_10_action_list.csv` from the output folder |
| 6 | Show the ranked rows — widen columns so all text is visible | Key columns: rank, customer_name, exception_class, priority_score, recommended_action |
| 7 | Scroll right slowly to show the recommended_action column | This is the payoff — "here's what to do" |
| 8 | Hold on the action list until clip ends | |

**Tip:** Run the tool in sample mode before recording so you know exactly which CSV
to open and what the columns look like. No surprises mid-recording.

---

### V4_C06 — Control Evidence Pack (~1 min 26 sec)

**What's on screen:** CLI menu → file list with hashes → evidence_summary.html in browser

| # | Action | Notes |
|---|---|---|
| 1 | Return to CLI menu | |
| 2 | Type "4" and press Enter — Control Evidence Pack runs in sample mode | It scans the most recent Revenue Leakage output folder |
| 3 | Show file names and SHA-256 hashes printing to terminal | Let the list complete — it's short (5–6 files) |
| 4 | At narration "It logs exactly which files were analyzed..." | Hold on the terminal output for ~5 seconds |
| 5 | At narration "This is the evidence summary..." | Switch to browser — open `evidence_summary.html` from the output folder |
| 6 | Show the full one-page evidence summary | Scroll slowly — let viewer read the file list and timestamps |
| 7 | Hold on the HTML page until clip ends | |

**Tip:** The evidence pack runs on the revenue leakage output folder. Run Chapter 3's
demo first, then run this — it has content to hash. In sample mode it handles this automatically.

---

### V4_C07 — Finance Automation Launcher (~60 sec)

**What's on screen:** Excel button → full menu → option 7 → outputs folder in Explorer → option 8

| # | Action | Notes |
|---|---|---|
| 1 | Return to FinanceTools.xlsm — show the Finance Tools button clearly | Full Excel window visible, not just the button |
| 2 | At narration "Click this button in Excel, and you get the numbered menu..." | Click the button — menu opens |
| 3 | Show all 8 options for ~5 seconds | Let viewer read the full menu |
| 4 | At narration "Option 7 opens your outputs folder directly in Explorer" | Type "7" and press Enter — File Explorer opens on outputs/ |
| 5 | Hold on File Explorer showing the timestamped output folders for ~5 seconds | |
| 6 | Return to the menu (reopen it) | |
| 7 | Type "8" and press Enter — menu closes | |
| 8 | Show Excel window in the foreground — clean ending | |

**Tip:** This is the chapter where the button is the hero. Linger on the Excel sheet
with the button before clicking — let the viewer absorb the "this is all you click" message.

---

### V4_C08 — How to Start (~32 sec)

**What's on screen:** Static text card with the four rules

| # | Action | Notes |
|---|---|---|
| 1 | Show a clean text card with the four rules | Same style as Chapter 1 title card if possible |
| 2 | Hold for full clip duration (~32 sec) | No animation, no demo — just the rules on screen |

**Four rules to show on screen:**
1. Run in sample mode first
2. Start with the supported workflows
3. Outputs go to the outputs folder — input files never touched
4. Questions? Contact Connor — Finance & Accounting

---

## POST-RECORDING CHECKLIST

- [ ] All 9 clips recorded (V4_C01, C02, C03a, C03b, C04, C05, C06, C07, C08)
- [ ] Playback each clip before finalizing — check sync between narration and screen action
- [ ] Confirm CLI window is readable (font size, contrast) in every chapter that shows it
- [ ] Confirm HTML reports are readable at your recording resolution
- [ ] Confirm Excel columns are wide enough to read in C03b and C05
- [ ] No personal file paths, real customer names, or sensitive data visible on screen
- [ ] outputs/ folder shown contains only sample data outputs — not real file names

---

## SCREEN SETUP REFERENCE

| Chapter | Apps needed open |
|---|---|
| C01 | Static slide or title card |
| C02 | Notepad (PYTHON_SAFETY.md), File Explorer (outputs/) |
| C03a | Excel (FinanceTools.xlsm), CLI window |
| C03b | Browser (leakage_report.html), Excel (exceptions_ranked.csv) |
| C04 | CLI window, Notepad (billing CSV), CLI window again |
| C05 | CLI window, Excel (top_10_action_list.csv) |
| C06 | CLI window, Browser (evidence_summary.html) |
| C07 | Excel (FinanceTools.xlsm), CLI window, File Explorer |
| C08 | Static text card |

---

*End of shot list. Version v1.0 — 2026-04-28.*
*Use alongside VIDEO_4_NARRATION_SCRIPT_v1.md when recording.*
