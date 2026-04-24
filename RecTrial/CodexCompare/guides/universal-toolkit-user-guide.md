# Universal Toolkit User Guide

## 1) Audience

This guide is for non-developers who want to run the universal toolkit in Excel.

## 2) Prerequisites

- Excel desktop app with macros enabled
- Access to the workbook/add-in that contains toolkit modules
- A copy of your file for testing

## 3) First-time setup

1. Open Excel.
2. Open your test workbook copy.
3. Import the universal modules (`vba/universal/*.bas`) into the VBA project, or load the toolkit add-in.
4. Save as `.xlsm`.
5. Run `BuildCommandCenter`.

Expected result: a sheet named `UTL_CommandCenter` appears.

## 4) Quickstart (15 minutes)

Run these in order:
1. **Preview Sanitizer Impact**
2. **Run Full Workbook Sanitizer**
3. **Create Workbook Profile**
4. **Consolidate Visible Sheets**
5. **Build Executive One-Pager**

## 5) What each action does

- **Preview Sanitizer Impact**: estimates fixes without changing data.
- **Run Full Workbook Sanitizer**: cleans text/number/date quality issues.
- **Create Workbook Profile**: maps sheets, headers, and ranges.
- **Consolidate Visible Sheets**: appends visible sheets into one combined output.
- **Build Executive One-Pager**: builds a KPI summary sheet for leadership review.

## 6) Safety rules

- Always run on a copy first.
- Keep output non-destructive.
- Review `UTL_RunLog` after each action.

## 7) Troubleshooting

- If a macro fails, read the message box and check `UTL_RunLog`.
- If command center buttons fail, rebuild by running `BuildCommandCenter` again.
- If your workbook has merged headers, run sanitizer first before compare/consolidation.

## 8) FAQ

**Q: Can I run this on any workbook?**  
A: Yes, universal modules are designed to be file-agnostic.

**Q: Will it change my source sheets?**  
A: Some actions write to new sheets; sanitizer updates values in-place, so run on a copy first.

**Q: Where do outputs go?**  
A: Mostly new sheets in the same workbook, plus optional PDF files in the workbook folder.
