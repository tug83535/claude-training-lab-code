# KBT P&L Toolkit — Quick Start Guide

> **Time to first result: ~10 minutes**
> You will: enable macros, open the Command Center, and run your first 5 commands.

---

## Prerequisites

- Windows PC with Excel Desktop (O365 or Excel 2019+)
- The workbook: `KeystoneBenefitTech_PL_Model.xlsm`
- The 29 VBA `.bas` module files (provided in the `03_Code/VBA/` folder)

---

## Step 1 — Enable Macros (2 min)

1. Open Excel (do **not** open the workbook yet).
2. Go to **File → Options → Trust Center → Trust Center Settings**.
3. Click **Macro Settings** on the left.
4. Select **"Enable all macros"** (or "Disable with notification" if your IT policy requires).
5. Check **"Trust access to the VBA project object model"** (needed for automatic form building).
6. Click **OK → OK** to close both dialogs.
7. Restart Excel.

---

## Step 2 — Import VBA Modules (5 min)

1. Open `KeystoneBenefitTech_PL_Model.xlsm`. If prompted, click **Enable Content**.
2. Press **Alt+F11** to open the VBA Editor.
3. In the left panel, right-click on **VBAProject** → **Import File...**
4. Navigate to the `03_Code/VBA/` folder.
5. Select all 29 `.bas` files and click **Open**.
6. Verify: the left panel should now show all `mod*` modules under "Modules".
7. Close the VBA Editor (Alt+Q or click the X).
8. Save the workbook (**Ctrl+S**) — Excel may prompt to save as `.xlsm`.

---

## Step 3 — Launch the Command Center (1 min)

Press **Ctrl+Shift+M** anywhere in the workbook.

You should see the Command Center — a UserForm with categories on the left and 50 available actions on the right:

```
╔══════════════════════════════════════════════════════════╗
║  KBT P&L Toolkit v2.1.0 — Command Center               ║
╠══════════════════════════════════════════════════════════╣
║                                                          ║
║  Categories:          Available Actions:                 ║
║  ┌──────────────┐    ┌───┬─────────────────────────┐    ║
║  │ All Actions   │    │ # │ Name                    │    ║
║  │ Monthly Ops   │    ├───┼─────────────────────────┤    ║
║  │ Analysis      │    │ 1 │ Generate Monthly Tabs   │    ║
║  │ Data Quality  │    │ 2 │ Delete Generated Tabs   │    ║
║  │ Reporting     │    │ 3 │ Run Reconciliation      │    ║
║  │ Utilities     │    │ 4 │ Export Recon Report     │    ║
║  │ Data & Import │    │ 5 │ Sensitivity Analysis    │    ║
║  │ Forecasting   │    │ 6 │ Variance Analysis       │    ║
║  │ Scenarios     │    │ 7 │ Scan Data Quality       │    ║
║  │ Allocation    │    │...│ ...                     │    ║
║  │ Consolidation │    │50 │ About This Toolkit      │    ║
║  │ Version Ctrl  │    └───┴─────────────────────────┘    ║
║  │ Governance    │                                       ║
║  │ Admin & Test  │    Search: [_______________]          ║
║  │ Advanced      │                                       ║
║  └──────────────┘    [  Run  ]  [ Cancel ]               ║
║                                                          ║
║  50 actions shown                                        ║
╚══════════════════════════════════════════════════════════╝
```

> **Didn't get the form?** If an InputBox appears instead, that means the UserForm
> hasn't been built yet. Type the action number and press OK. See the
> Implementation Guide for UserForm setup instructions.

---

## Step 4 — Run Your First 5 Commands

These are the essential operations for any month. Run them in this order:

| Order | # | Command | What It Does |
|-------|---|---------|--------------|
| 1 | **7** | Scan Data Quality | Scans the GL for blanks, duplicates, text-stored numbers, and date issues. Always run this first to check your data. |
| 2 | **3** | Run Reconciliation | Runs PASS/FAIL checks across all sheets. Verifies GL totals match P&L Trend, checks allocation shares, and flags discrepancies. |
| 3 | **6** | Variance Analysis | Compares the current month to the prior month. Flags any line items that moved more than 15%. |
| 4 | **12** | Build Dashboard | Creates 3 charts: revenue trend, contribution margin %, and revenue mix pie chart. |
| 5 | **10** | Export Report Package | Exports all report sheets as a professional PDF with headers, footers, and page numbers. |

---

## Step 5 — Verify Everything Worked

After running the 5 commands above, you should see:

- **Data Quality Report** sheet — listing any data issues found
- **Checks** sheet — with PASS/FAIL results for each reconciliation
- **Variance Analysis** sheet — showing MoM dollar and percent changes
- **Dashboard** sheet — with 3 charts
- A PDF file saved to your workbook's folder

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| **Ctrl+Shift+M** | Open Command Center |
| **Ctrl+Shift+H** | Go Home (jump to Report sheet) |
| **Ctrl+Shift+J** | Quick Jump (pick any sheet) |
| **Ctrl+Shift+R** | Run All Reconciliation Checks |

---

## Optional: Python Analytics

If you want the parallel Python analytics suite (interactive dashboard, forecasting, AP matching), see the Implementation Guide's "Python Setup" section.

Quick version:
```
pip install -r requirements.txt
python pnl_runner.py --help
python pnl_runner.py dashboard        # Opens interactive Streamlit dashboard
```

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| "Macros have been disabled" | Enable macros in Trust Center (Step 1) |
| "Compile error: Sub not found" | Not all modules imported — reimport from `03_Code/VBA/` |
| InputBox appears instead of UserForm | Run `modFormBuilder.BuildCommandCenter` once from VBA Editor, or follow Mode B manual install |
| "Subscript out of range" error | Sheet name mismatch — verify tab names match exactly (see modConfig constants) |
| Charts appear blank | Ensure at least 1 month of data exists in the GL tab |

---

## Next Steps

- **Full setup details** → `IMPLEMENTATION_GUIDE.md`
- **All 50 commands explained** → `USER_TRAINING_GUIDE.md`
- **Monthly workflow** → `OPERATIONS_RUNBOOK.md`
- **Version history** → `CHANGELOG.md`
