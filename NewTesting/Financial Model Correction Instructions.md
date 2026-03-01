# Financial Model Correction Instructions

Please update and regenerate the Keystone BenefitTech P&L Model to fix the following structural and mathematical issues. Stop treating the output as raw text tables and format it as a dynamic, production-ready Excel workbook.

## 1. Fix Data Type Errors (Critical)
* **Issue:** In the `Assumptions` sheet, the *InsureSight Software Split* (0.22) is stored as text. 
* **Fix:** Convert this, and all other numerical drivers, to strictly numeric data types (Percentages/Decimals) to prevent downstream `#VALUE!` errors.

## 2. Eliminate Floating-Point Anomalies (Critical)
* **Issue:** The model relies on raw, unrounded math (e.g., AWS Compute Share sums to `0.99999999999999989`, Contribution Margin is `0.5787034647912902`).
* **Fix:** Wrap ALL financial calculations in `=ROUND([formula], 2)` for currency and `=ROUND([formula], 4)` for percentages. Apply proper Excel cell formatting so they display cleanly.

## 3. Implement Check Tolerances (Critical)
* **Issue:** The `Reconciliation Checks` tab looks for exact zero differences, which will eventually trigger false "FAIL" statuses due to Excel's floating-point precision quirks.
* **Fix:** Rewrite the check formulas to include a penny-tolerance threshold: `=IF(ABS([Sheet A Value] - [Sheet B Value]) < 0.01, "PASS", "FAIL")`.

## 4. Consolidate Flat Monthly Tabs (Structural)
* **Issue:** The model has separate static tabs for `Functional P&L Summary - Jan 25`, `Feb 25`, and `Mar 25`. This is inefficient and hard to scale.
* **Fix:** Delete the individual month tabs. Create one single, dynamic `Functional P&L Summary` tab. Provide the exact `SUMIFS`, `INDEX/MATCH`, or `XLOOKUP` formulas required to populate this tab based on a Data Validation drop-down menu where the user selects the month.

## 5. Remove Hardcoded File Paths (Navigation)
* **Issue:** The `Report-->` (Table of Contents) tab uses local hard drive paths for links (e.g., `file:///C:\Users\connor.atlee\Downloads\...`).
* **Fix:** Replace all external links with native Excel internal hyperlinks using the format: `=HYPERLINK("#'SheetName'!A1", "Link Description")`.

## 6. Upgrade Data Ingestion to Power Query (Structural)
* **Issue:** The model relies on a massive, static data dump in `CrossfireHiddenWorksheet`. 
* **Fix:** Do not generate raw mock transaction data. Instead, write the exact **Power Query M-Code** required to ingest the raw CSV, clean it, map it to the `Data Dictionary`, and load it into the Data Model. 
