# VIDEO 4 — Interactive Recording Guide (Updated)

**Use this document with Claude.ai to build an interactive follow-along checklist.**

**What Video 4 Is:** Manual recording from Command Prompt. You play each audio clip, run a Python command, and show the output in Excel. 10 clips total, ~6-8 minutes.

**Files needed:**
- Audio clips in `RecTrial\AudioClips\Video4\` (10 MP3s)
- Demo input files in `RecTrial\Video4DemoFiles\`
- Python scripts in `RecTrial\UniversalToolkit\python\`

---

## PRE-RECORDING SETUP

### Step 1: Install Python Packages
- [ ] Open Command Prompt (Win+R → cmd → Enter)
- [ ] Run: `cd C:\Users\connor.atlee\RecTrial\UniversalToolkit\python`
- [ ] Run: `pip install pandas openpyxl pdfplumber thefuzz python-Levenshtein python-dateutil python-docx`
- [ ] Verify: `python --version` shows 3.x

### Step 2: Test One Script
- [ ] Run: `python compare_files.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v1.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v2.xlsx"`
- [ ] Verify: COMPARISON_REPORT.xlsx is created
- [ ] Delete the output file

### Step 3: Check All Script Help
- [ ] Run `python fuzzy_lookup.py --help` and note the flags
- [ ] Run `python bank_reconciler.py --help` and note the flags
- [ ] Run `python aging_report.py --help` and note the flags
- [ ] Run `python variance_decomposition.py --help` and note the flags
- [ ] Run `python forecast_rollforward.py --help` and note the flags
- [ ] Run `python variance_analysis.py --help` and note the flags

### Step 4: Verify Demo Files Exist
- [ ] `Video4DemoFiles\Q1_Revenue_v1.xlsx` — for compare
- [ ] `Video4DemoFiles\Q1_Revenue_v2.xlsx` — for compare
- [ ] `Video4DemoFiles\our_vendor_list.xlsx` — for fuzzy lookup
- [ ] `Video4DemoFiles\bank_vendor_list.xlsx` — for fuzzy lookup
- [ ] `Video4DemoFiles\gl_ledger.xlsx` — for bank reconciler
- [ ] `Video4DemoFiles\bank_statement.xlsx` — for bank reconciler
- [ ] `Video4DemoFiles\open_invoices.xlsx` — for aging report
- [ ] `Video4DemoFiles\product_budget_vs_actual.xlsx` — for variance decomp
- [ ] `Video4DemoFiles\monthly_revenue_history.xlsx` — for forecast
- [ ] `Video4DemoFiles\sample_report.pdf` — for PDF extractor
- [ ] `Video4DemoFiles\budget_files\` — 7 department files for variance analysis

### Step 5: Style Command Prompt
- [ ] Maximize Command Prompt
- [ ] Right-click title bar → Properties → Font → Consolas 16pt
- [ ] Right-click title bar → Properties → Colors → dark background

### Step 6: Computer Lockdown + OBS
- [ ] Same as Videos 1-3 (notifications off, clean desktop, etc.)
- [ ] OBS Recording Path: `RecTrial\Recordings\Video4\`
- [ ] Desktop Audio: ENABLED, Mic: DISABLED

---

## RECORDING — Clip by Clip

### How each clip works:
1. Pre-type the command in Command Prompt (don't press Enter yet)
2. Start OBS recording
3. Wait 2 seconds
4. Play the audio clip (double-click MP3)
5. When narration mentions running the script → press Enter
6. Wait for output → open output in Excel to show results
7. When audio finishes → hold 2 seconds → stop OBS
8. Delete output file before next clip

---

### CLIP 1 — Opening (~30 sec)
**Audio:** V4_S0_Opening.mp3
**Action:** Show Video4DemoFiles folder in File Explorer, slowly scroll through files
**No command to run**

---

### CLIP 2 — Compare Files (~50 sec)
**Audio:** V4_S1_CompareFiles.mp3
**Pre-type:**
```
python compare_files.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v1.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\Q1_Revenue_v2.xlsx"
```
**When to press Enter:** When audio says "This script compares every cell"
**Show output:** Open COMPARISON_REPORT.xlsx → show color-coded diff
**Reset:** Close Excel, delete output

---

### CLIP 3 — PDF Extractor (~50 sec)
**Audio:** V4_S2_PDFExtractor.mp3
**Pre-type:**
```
python pdf_extractor.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\sample_report.pdf"
```
**When to press Enter:** When audio says "This script reads the P-D-F"
**Show output:** Open extracted Excel → show Revenue + Expense tables
**Reset:** Close Excel, delete output

---

### CLIP 4 — Fuzzy Lookup (~50 sec)
**Audio:** V4_S3_FuzzyLookup.mp3
**Pre-type:**
```
python fuzzy_lookup.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\our_vendor_list.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\bank_vendor_list.xlsx" --source-col "Vendor Name" --lookup-col "Vendor Name"
```
**Note:** Check `--help` first — flags may differ
**When to press Enter:** When audio says "This script uses fuzzy matching"
**Show output:** Open results → show green/yellow/red matches
**Reset:** Close Excel, delete output

---

### CLIP 5 — Bank Reconciler (~50 sec)
**Audio:** V4_S4_BankReconciler.mp3
**Pre-type:**
```
python bank_reconciler.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\gl_ledger.xlsx" "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\bank_statement.xlsx"
```
**Note:** Check `--help` first
**When to press Enter:** When audio says "This script matches them"
**Show output:** Open results → show matched/fuzzy/unmatched items
**Reset:** Close Excel, delete output

---

### CLIP 6 — Aging Report (~45 sec)
**Audio:** V4_S5_AgingReport.mp3
**Pre-type:**
```
python aging_report.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\open_invoices.xlsx" --date-col "Invoice Date" --amount-col "Amount"
```
**Note:** Check `--help` first
**When to press Enter:** When audio says "Give this script a file"
**Show output:** Open aging report → show buckets (Current, 0-30, 31-60, 61-90, 90+)
**Reset:** Close Excel, delete output

---

### CLIP 7 — Variance Decomposition (~50 sec)
**Audio:** V4_S6_VarianceDecomp.mp3
**Pre-type:**
```
python variance_decomposition.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\product_budget_vs_actual.xlsx"
```
**When to press Enter:** When audio says "This script takes your actual and budget data"
**Show output:** Open results → show Price/Volume/Mix effects color-coded
**Reset:** Close Excel, delete output

---

### CLIP 8 — Forecast Rollforward (~45 sec)
**Audio:** V4_S7_ForecastRoll.mp3
**Pre-type:**
```
python forecast_rollforward.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\monthly_revenue_history.xlsx" --date-col "Date" --value-col "Revenue"
```
**Note:** Check `--help` first
**When to press Enter:** When audio says "Give this script your historical actuals"
**Show output:** Open forecast → show actuals + forecast + line chart
**Reset:** Close Excel, delete output

---

### CLIP 9 — Variance Analysis (~45 sec)
**Audio:** V4_S8_VarianceAnalysis.mp3
**Pre-type:**
```
python variance_analysis.py "C:\Users\connor.atlee\RecTrial\Video4DemoFiles\budget_files"
```
**When to press Enter:** When audio says "Point this script at a folder"
**Show output:** Open results → show consolidated variance + bar chart
**Reset:** Close Excel, delete output

---

### CLIP 10 — Closing (~30 sec)
**Audio:** V4_S9_Closing.mp3
**Action:** Re-run a couple scripts so output files are visible in folder. Show File Explorer with all outputs. Slowly scroll.
**No command to run**

---

## AFTER RECORDING

- [ ] Review all 10 recordings in `RecTrial\Recordings\Video4\`
- [ ] Check each clip: audio plays, command visible, output shown, no errors
- [ ] Note any clips to re-record
- [ ] Re-record only the problem clips (not all 10)

---

*Updated: 2026-04-15*
