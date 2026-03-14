# ExeTest — Standalone Distribution Packages

This folder contains everything needed to build two distribution packages
for coworkers who want to use the KBT Universal Tools without any setup.

---

## Folder Contents

```
ExeTest/
  README.md                  <-- You are here
  python_exe/
    requirements.txt         <-- Python dependencies (install first)
    build_all_exe.bat        <-- Double-click to build all 22 .exe files
  xlam_kit/
    BUILD_XLAM_INSTRUCTIONS.md  <-- Step-by-step .xlam build guide
    modUTL_*.bas (23 files)     <-- All universal VBA modules
```

---

## Package 1: Python .exe Tools (22 standalone executables)

**What:** 22 Python scripts converted to standalone `.exe` files. Coworkers
double-click or run from command line — no Python installation needed.

**How to build:**
1. Install Python 3.10+ on your machine
2. Open Command Prompt in `ExeTest\python_exe\`
3. Run: `pip install -r requirements.txt`
4. Double-click `build_all_exe.bat`
5. Wait for all 22 builds to finish (~10-15 minutes)
6. Find all `.exe` files in the `dist\` folder

**The 22 tools:**

| # | .exe Name | What It Does |
|---|-----------|-------------|
| 1 | KBT_CleanData | Clean any Excel file (trim, dedupe, fix numbers, standardize dates) |
| 2 | KBT_CompareFiles | Cell-by-cell comparison of two Excel files with diff report |
| 3 | KBT_ConsolidateFiles | Combine hundreds of Excel files from a folder into one master |
| 4 | KBT_ConsolidateBudget | Merge department budget files with variance columns |
| 5 | KBT_VarianceAnalysis | Actual vs Budget comparison with waterfall breakdown |
| 6 | KBT_VarianceDecomposition | Price/Volume/Mix variance decomposition (FP&A standard) |
| 7 | KBT_AgingReport | AR/AP aging buckets: Current, 0-30, 31-60, 61-90, 90+ days |
| 8 | KBT_BankReconciler | Fuzzy-match bank statement to ledger (catches near-matches) |
| 9 | KBT_GLReconciliation | GL vs Sub-ledger matching by amount + date + reference |
| 10 | KBT_ReconExceptions | Output only unmatched exceptions between two files |
| 11 | KBT_FuzzyLookup | Fuzzy string matching between two datasets (vendor dedup, etc.) |
| 12 | KBT_MasterDataMapper | SQL-style joins between Excel files (replaces nested VLOOKUPs) |
| 13 | KBT_ForecastRollforward | 12-month rolling forecast (moving avg, growth rate, or flat) |
| 14 | KBT_BatchProcess | Run the data cleaner on every file in a folder automatically |
| 15 | KBT_UnpivotData | Convert wide pivot format to tall database format |
| 16 | KBT_RegexExtractor | Extract invoices, emails, phone numbers, dates from free text |
| 17 | KBT_PDFExtractor | Pull tables out of PDF documents into Excel |
| 18 | KBT_WordReport | Generate formatted Word documents from Excel data |
| 19 | KBT_SQLQueryTool | Run SQL queries against CSV/Excel files (SELECT, JOIN, WHERE) |
| 20 | KBT_MultiFileConsolidator | Smart multi-file combine with column mismatch handling |
| 21 | KBT_TwoFileReconciler | Two-file reconciliation with row-level exception detail |
| 22 | KBT_DateFormatUnifier | Standardize all date formats across a file to one format |

**How coworkers use them:**
- Open Command Prompt
- Run: `KBT_CleanData.exe "C:\path\to\your_file.xlsx"`
- Or run: `KBT_CleanData.exe --help` to see all options

---

## Package 2: Excel Add-In (.xlam) — 23 VBA Modules, ~140+ Tools

**What:** All 23 universal VBA modules packaged as an Excel Add-In. Install
once, and every tool is available from a menu in every workbook you open.

**How to build:** See `xlam_kit/BUILD_XLAM_INSTRUCTIONS.md` for the full
step-by-step guide.

**Quick version:**
1. Open Excel > Alt+F11 > Import all 23 `.bas` files from `xlam_kit/`
2. Debug > Compile (should be clean)
3. File > Save As > Excel Add-In (.xlam) > name it `KBT_UniversalTools.xlam`
4. File > Options > Add-Ins > Go > Browse > select the `.xlam` > OK

---

## What to Share With Coworkers

| Audience | Give Them |
|----------|-----------|
| Everyone (easiest) | The `.xlam` file — install and use from Excel menu |
| Power users | The `.exe` files — run from command line for batch processing |
| Both | Share both — they complement each other |

---

## Notes
- The `.exe` files are large (~30-50 MB each) because they bundle Python inside
- The `.xlam` file is small (~500 KB) and works in Excel 2016+
- Both packages are 100% standalone — zero dependencies after installation
