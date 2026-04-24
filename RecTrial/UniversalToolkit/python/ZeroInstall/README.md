# Zero-Install Python Tools

Seven small command-line tools that run on **plain Python 3.8+** with **no `pip install` required.** They use only the Python standard library — no pandas, no openpyxl, no third-party packages.

Designed for coworkers whose corporate laptops don't allow installing Python packages, but still want useful Finance automation from the command prompt.

## The tools

| Script | Input | What it does | Output |
|---|---|---|---|
| `profile_workbook.py` | `.xlsx` or `.xlsm` file | Inventories every sheet: name, dimensions, named ranges, whether the workbook contains VBA | Markdown summary printed to stdout |
| `sanitize_dataset.py` | `.csv` file | Normalizes text (trim), numbers (strip currency/commas/parens), dates (to ISO) | Cleaned CSV to stdout or file |
| `compare_workbooks.py` | Two `.xlsx` files | Row-level diff between workbooks, sheet-by-sheet | CSV diff report |
| `build_exec_summary.py` | `.csv` (finance data) | Totals, top groups by value, plain-English "suggested talking points" | Markdown exec summary |
| `variance_classifier.py` | `.csv` with Actual + Baseline columns | Classifies each row: Material increase/decrease, Watch, Normal | Labelled CSV |
| `scenario_runner.py` | `.csv` with a numeric metric column | Applies percentage shocks (e.g. +10%, -5%) and summarizes impact | Scenario results CSV |
| `sheets_to_csv.py` | `.xlsx` | Extracts every sheet to a separate CSV file in an output folder | Folder of CSVs |

## How to use

From a command prompt:

```
cd C:\path\to\ZeroInstall\
python profile_workbook.py C:\path\to\your_file.xlsx
python sanitize_dataset.py C:\path\to\your_data.csv > cleaned.csv
python build_exec_summary.py C:\path\to\your_data.csv > summary.md
```

Every script supports `--help` for its full argument list:

```
python profile_workbook.py --help
```

## When to use these vs the full toolkit

| Situation | Use |
|---|---|
| Your laptop allows `pip install` | `UniversalToolkit\python\` (the full suite, using pandas/openpyxl — more features, more formats) |
| Your laptop blocks `pip install`, or you just want something portable | **ZeroInstall/** (these scripts, stdlib only) |

The full-toolkit scripts and the zero-install scripts are **complementary**, not replacements. You can use both side by side.

## Requirements

- Python 3.8 or newer (check with `python --version`)
- Nothing else

## Notes

- Excel `.xlsm` files: `profile_workbook.py` reads the structure but does not execute VBA macros.
- `compare_workbooks.py` handles `.xlsx` only (not `.xlsm`). Save your `.xlsm` as `.xlsx` first if needed.
- All scripts print helpful error messages if they fail — no silent failures.
