# VIDEO 4 — "Python Automation for Finance"
# Manual Recording Guide — Step by Step

**What This Is:** Complete step-by-step guide for manually recording Video 4. Unlike Videos 1-3 (which used the Director macro), Video 4 is recorded manually because the Python scripts run from Command Prompt, not Excel.

**Runtime Target:** 6-8 minutes
**Clips:** 10 (opening + 8 demos + closing)
**Recording Method:** You play each audio clip on your computer, run the Python command while OBS records, then open the output file to show the results.

---

## Table of Contents

1. [What You Need](#what-you-need)
2. [Folder Setup](#folder-setup)
3. [Computer Lockdown](#computer-lockdown)
4. [OBS Setup](#obs-setup)
5. [Python Setup](#python-setup)
6. [Recording — Clip by Clip](#recording--clip-by-clip)
7. [After Recording](#after-recording)
8. [Troubleshooting](#troubleshooting)

---

## What You Need

- [ ] All 10 audio clips in `RecTrial\AudioClips\Video4\`
- [ ] All demo input files in `RecTrial\Video4DemoFiles\`
- [ ] Sample PDF: `RecTrial\Video4DemoFiles\sample_report.pdf`
- [ ] Python installed (3.7+) with required packages
- [ ] OBS Studio configured
- [ ] A media player that can play MP3 files (Windows Media Player, VLC, or just double-click)

---

## Folder Setup

Everything is in one place:

```
RecTrial\
├── AudioClips\Video4\          (10 MP3 narration clips)
├── Video4DemoFiles\            (all input files for the demos)
│   ├── Q1_Revenue_v1.xlsx
│   ├── Q1_Revenue_v2.xlsx
│   ├── our_vendor_list.xlsx
│   ├── bank_vendor_list.xlsx
│   ├── gl_ledger.xlsx
│   ├── bank_statement.xlsx
│   ├── open_invoices.xlsx
│   ├── product_budget_vs_actual.xlsx
│   ├── monthly_revenue_history.xlsx
│   ├── sample_report.pdf
│   └── budget_files\           (7 department budget files)
├── Recordings\Video4\          (OBS saves here)
└── UniversalToolkit\python\    (the Python scripts)
```

---

## Computer Lockdown

Same as Videos 1-3:

1. Close everything except OBS and Command Prompt (and a media player for audio)
2. Turn off ALL notifications (Focus Assist → Alarms Only)
3. Hide desktop icons
4. Auto-hide taskbar
5. Display: 1920x1080, 100% scaling
6. Plug in laptop
7. Phone on silent

---

## OBS Setup

1. Open OBS Studio
2. Settings:
   - Recording Path: `RecTrial\Recordings\Video4\`
   - Format: MP4, 1920x1080, 30 FPS
   - **Desktop Audio: ENABLED** (captures the narration audio you play)
   - **Mic: DISABLED**
3. Display Capture source active

---

## Python Setup

Before recording, verify Python and packages work.

1. Open **Command Prompt** (Win+R → type `cmd` → Enter)
2. Navigate to the scripts folder:
   ```
   cd C:\Users\connor.atlee\RecTrial\UniversalToolkit\python
   ```
3. Test Python:
   ```
   python --version
   ```
   Should show Python 3.x

4. Install required packages (if not already):
   ```
   pip install pandas openpyxl pdfplumber thefuzz python-Levenshtein python-dateutil python-docx
   ```

5. Test one script to verify:
   ```
   python compare_files.py C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v1.xlsx C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v2.xlsx
   ```
   Should create a COMPARISON_REPORT.xlsx file. If it works, delete the output and proceed.

---

## Recording — Clip by Clip

### How each clip works:

1. **Arrange your screen:** Command Prompt on the left, File Explorer or Excel on the right (to show output files)
2. **Start OBS recording**
3. **Play the audio clip** (double-click the MP3 file — it plays through your speakers, OBS captures it)
4. **While audio plays:** Type and run the Python command in Command Prompt
5. **After the script finishes:** Open the output file in Excel to show the results on camera
6. **Wait for audio to finish** (hold still for 2-3 seconds after audio ends)
7. **Stop OBS recording**
8. **Delete the output file** before the next clip (clean state)

### Tip: Pre-type the commands

Before recording each clip, type the full command in Command Prompt but DO NOT press Enter yet. Then when you start recording and play the audio, just press Enter at the right moment. This avoids typos on camera.

---

### CLIP 1 — Opening (~30 sec)

**Audio:** `V4_S0_Opening.mp3`
**Screen:** Show the Video4DemoFiles folder in File Explorer (full screen)
**What to do:**
1. Open File Explorer to `RecTrial\Video4DemoFiles\`
2. Start OBS recording
3. Wait 2 seconds
4. Play V4_S0_Opening.mp3
5. While audio plays, slowly scroll through the files in the folder so the viewer sees all the input files
6. When audio finishes, wait 2 seconds
7. Stop OBS recording

---

### CLIP 2 — Compare Files (~50 sec)

**Audio:** `V4_S1_CompareFiles.mp3`
**Script:** `compare_files.py`
**Input:** `Q1_Revenue_v1.xlsx` + `Q1_Revenue_v2.xlsx`
**Output:** `COMPARISON_REPORT.xlsx`

**What to do:**
1. Command Prompt should be open and in the scripts folder:
   ```
   cd C:\Users\connor.atlee\RecTrial\UniversalToolkit\python
   ```
2. Pre-type (but don't press Enter yet):
   ```
   python compare_files.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v1.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v2.xlsx"
   ```
3. Start OBS recording
4. Wait 2 seconds
5. Play V4_S1_CompareFiles.mp3
6. When audio says "This script compares every cell" — press **Enter** to run the command
7. Wait for script to finish (5-10 seconds)
8. Open the output file (COMPARISON_REPORT.xlsx) in Excel
9. Scroll through the diff report — show the color-coded differences
10. When audio finishes, hold still 2 seconds
11. Stop OBS recording
12. **Reset:** Close Excel, delete COMPARISON_REPORT.xlsx

---

### CLIP 3 — PDF Extractor (~50 sec)

**Audio:** `V4_S2_PDFExtractor.mp3`
**Script:** `pdf_extractor.py`
**Input:** `sample_report.pdf`
**Output:** `PDF_EXTRACTED_TABLES.xlsx`

**What to do:**
1. Pre-type:
   ```
   python pdf_extractor.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\sample_report.pdf"
   ```
2. Start OBS recording
3. Wait 2 seconds
4. Play V4_S2_PDFExtractor.mp3
5. When audio says "This script reads the P-D-F" — press **Enter**
6. Wait for extraction (5-10 seconds)
7. Open output file in Excel — show the extracted tables (Revenue by Product, Expenses by Department)
8. When audio finishes, hold still 2 seconds
9. Stop OBS recording
10. **Reset:** Close Excel, delete output file

---

### CLIP 4 — Fuzzy Lookup (~50 sec)

**Audio:** `V4_S3_FuzzyLookup.mp3`
**Script:** `fuzzy_lookup.py`
**Input:** `our_vendor_list.xlsx` + `bank_vendor_list.xlsx`
**Output:** `FUZZY_MATCH_RESULTS.xlsx`

**What to do:**
1. Pre-type:
   ```
   python fuzzy_lookup.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\our_vendor_list.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\bank_vendor_list.xlsx" --source-col "Vendor Name" --lookup-col "Vendor Name"
   ```
2. Start OBS recording → wait 2 sec → play audio
3. When audio says "This script uses fuzzy matching" — press Enter
4. Wait for results → open output in Excel
5. Show the color-coded match results (green=exact, yellow=fuzzy, red=no match)
6. Hold still after audio → stop OBS
7. **Reset:** Close Excel, delete output

**Note:** Check the script's actual command-line arguments. If `--source-col` and `--lookup-col` aren't the right flags, run `python fuzzy_lookup.py --help` first to check.

---

### CLIP 5 — Bank Reconciler (~50 sec)

**Audio:** `V4_S4_BankReconciler.mp3`
**Script:** `bank_reconciler.py`
**Input:** `gl_ledger.xlsx` + `bank_statement.xlsx`
**Output:** Reconciliation report xlsx

**What to do:**
1. Pre-type:
   ```
   python bank_reconciler.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\gl_ledger.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\bank_statement.xlsx"
   ```
2. Start OBS → wait 2 sec → play audio
3. When audio says "This script matches them" — press Enter
4. Wait → open output in Excel
5. Show matched items (green), fuzzy matches (yellow), unmatched (red)
6. Hold still → stop OBS
7. **Reset:** Close Excel, delete output

**Note:** Check command-line args with `python bank_reconciler.py --help` before recording.

---

### CLIP 6 — Aging Report (~45 sec)

**Audio:** `V4_S5_AgingReport.mp3`
**Script:** `aging_report.py`
**Input:** `open_invoices.xlsx`
**Output:** Aging report xlsx

**What to do:**
1. Pre-type:
   ```
   python aging_report.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\open_invoices.xlsx" --date-col "Invoice Date" --amount-col "Amount"
   ```
2. Start OBS → wait 2 sec → play audio
3. When audio says "Give this script a file" — press Enter
4. Wait → open output in Excel
5. Show the aging buckets — Current, 0-30, 31-60, 61-90, 90+ with color coding
6. Hold still → stop OBS
7. **Reset:** Close Excel, delete output

**Note:** Check command-line args with `python aging_report.py --help`.

---

### CLIP 7 — Variance Decomposition (~50 sec)

**Audio:** `V4_S6_VarianceDecomp.mp3`
**Script:** `variance_decomposition.py`
**Input:** `product_budget_vs_actual.xlsx`
**Output:** Variance decomposition report xlsx

**What to do:**
1. Pre-type:
   ```
   python variance_decomposition.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\product_budget_vs_actual.xlsx"
   ```
2. Start OBS → wait 2 sec → play audio
3. When audio says "This script takes your actual and budget data" — press Enter
4. Wait → open output in Excel
5. Show Price Effect, Volume Effect, Mix Effect columns with color coding
6. Hold still → stop OBS
7. **Reset:** Close Excel, delete output

---

### CLIP 8 — Forecast Rollforward (~45 sec)

**Audio:** `V4_S7_ForecastRoll.mp3`
**Script:** `forecast_rollforward.py`
**Input:** `monthly_revenue_history.xlsx`
**Output:** Forecast report with chart xlsx

**What to do:**
1. Pre-type:
   ```
   python forecast_rollforward.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\monthly_revenue_history.xlsx" --date-col "Date" --value-col "Revenue"
   ```
2. Start OBS → wait 2 sec → play audio
3. When audio says "Give this script your historical actuals" — press Enter
4. Wait → open output in Excel
5. Show the combined actuals + forecast view and the line chart
6. Hold still → stop OBS
7. **Reset:** Close Excel, delete output

---

### CLIP 9 — Variance Analysis (~45 sec)

**Audio:** `V4_S8_VarianceAnalysis.mp3`
**Script:** `variance_analysis.py`
**Input:** `budget_files\` folder (7 department files)
**Output:** Consolidated variance report xlsx

**What to do:**
1. Pre-type:
   ```
   python variance_analysis.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\budget_files"
   ```
2. Start OBS → wait 2 sec → play audio
3. When audio says "Point this script at a folder" — press Enter
4. Wait → open output in Excel
5. Show the consolidated report with variance by department and bar chart
6. Hold still → stop OBS
7. **Reset:** Close Excel, delete output

---

### CLIP 10 — Closing (~30 sec)

**Audio:** `V4_S9_Closing.mp3`
**Screen:** Show the Video4DemoFiles folder again with all the OUTPUT files visible (don't delete them this time)

**What to do:**
1. Before recording, re-run a few scripts so output files are visible in the folder
2. Start OBS recording
3. Wait 2 seconds
4. Play V4_S9_Closing.mp3
5. Slowly scroll through File Explorer showing all the input AND output files
6. When audio finishes, hold still 3 seconds
7. Stop OBS recording

---

## After Recording

1. Review all 10 recordings in `RecTrial\Recordings\Video4\`
2. Check: audio plays, commands visible, output files shown, no errors
3. Note any clips that need re-recording
4. Re-record problem clips only (no need to redo all 10)

---

## Troubleshooting

### Script errors when running
- Check the command-line arguments — run `python script_name.py --help` to see what's expected
- Make sure file paths are correct (use full paths in quotes)
- Make sure packages are installed: `pip install pandas openpyxl pdfplumber thefuzz python-Levenshtein`

### Audio doesn't play
- Double-click the MP3 file — it should open in your default media player
- Make sure computer audio is not muted
- OBS captures Desktop Audio — make sure it's enabled in OBS settings

### Script runs but output file is empty
- Check the input file has data in the expected columns
- Check the column names match what the script expects
- Run `python script_name.py --help` for column name requirements

### Need to re-record just one clip
- Delete the old output file
- Re-run the setup for that specific clip
- Record just that clip with OBS
- Stitch it in during editing

### Command Prompt looks ugly on camera
- Maximize Command Prompt to full screen
- Right-click title bar → Properties → Font → set to Consolas 16pt
- Right-click title bar → Properties → Colors → set background to dark blue or black
- This makes the terminal look more professional on camera

---

## Pre-Recording Checklist

Before starting Clip 1:

- [ ] Python works (`python --version` shows 3.x)
- [ ] All packages installed (pandas, openpyxl, pdfplumber, thefuzz, etc.)
- [ ] All 10 audio clips in `AudioClips\Video4\`
- [ ] All input files in `Video4DemoFiles\`
- [ ] sample_report.pdf exists and has tables
- [ ] OBS configured (1920x1080, Desktop Audio ON, Mic OFF)
- [ ] Command Prompt open, navigated to `UniversalToolkit\python\`
- [ ] Computer lockdown done (no notifications, clean desktop)
- [ ] Test run of at least one script completed successfully
- [ ] Command Prompt font set to Consolas 16pt for readability

---

## Quick Reference — All Commands

```
# Navigate to scripts folder first:
cd C:\Users\connor.atlee\RecTrial\UniversalToolkit\python

# Clip 2: Compare Files
python compare_files.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v1.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v2.xlsx"

# Clip 3: PDF Extractor
python pdf_extractor.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\sample_report.pdf"

# Clip 4: Fuzzy Lookup
python fuzzy_lookup.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\our_vendor_list.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\bank_vendor_list.xlsx" --source-col "Vendor Name" --lookup-col "Vendor Name"

# Clip 5: Bank Reconciler
python bank_reconciler.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\gl_ledger.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\bank_statement.xlsx"

# Clip 6: Aging Report
python aging_report.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\open_invoices.xlsx" --date-col "Invoice Date" --amount-col "Amount"

# Clip 7: Variance Decomposition
python variance_decomposition.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\product_budget_vs_actual.xlsx"

# Clip 8: Forecast Rollforward
python forecast_rollforward.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\monthly_revenue_history.xlsx" --date-col "Date" --value-col "Revenue"

# Clip 9: Variance Analysis
python variance_analysis.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\budget_files"
```

---

*Recording Guide created: 2026-04-02*
*Video 4 of 4: Python Automation for Finance*
