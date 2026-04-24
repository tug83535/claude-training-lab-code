# Demo Walkthrough Guide (P&L Workbook)

## 1) Purpose

This guide walks you through the file-specific demo using `samples/ExcelDemoFile_adv.xlsm`.

## 2) Demo sequence

Run in this order:
1. `BuildDemoCommandCenter`
2. `RunDemoReconciliation`
3. `GenerateDemoVarianceNarrative`
4. `RunDemoWhatIfScenarios`
5. `BuildDemoExecutiveBriefPack`

## 3) Expected outputs

- `Checks` updated with current run checks
- `Exec_Variance_Narrative` created/refreshed
- `Scenario_Compare` created/refreshed
- `Exec_Brief` created/refreshed
- PDF brief exported when workbook path is available

## 4) Presenter talking points

- Reconciliation confirms data trust before insight generation.
- Variance narrative turns numbers into plain-English finance commentary.
- Scenario comparison supports decision speed for leadership.
- Executive brief packs key KPIs and controls in one branded artifact.

## 5) Reset between demo runs

1. Close workbook without saving.
2. Reopen clean copy.
3. Rebuild command center.

## 6) Troubleshooting

- Missing sheet error: confirm workbook has required tabs from `modDemo_Config`.
- PDF export error: save workbook first so `ThisWorkbook.Path` is not empty.
- Empty narrative output: verify monthly trend sheet has numeric start/end period values.
