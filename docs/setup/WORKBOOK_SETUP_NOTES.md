# Workbook Setup Notes

> **File:** `KeystoneBenefitTech_PL_Model.xlsm`
> **Format:** Excel Macro-Enabled Workbook

---

## Preparing the Workbook

### Step 1 — Convert to .xlsm

The source workbook is distributed as `.xlsx` (no macros). To use VBA automation:

1. Open `KeystoneBenefitTech_PL_Model.xlsx` in Excel.
2. Go to **File → Save As**.
3. Change **Save as type** to **"Excel Macro-Enabled Workbook (*.xlsm)"**.
4. Save to your working directory.
5. Close the `.xlsx` and work only in the `.xlsm` going forward.

### Step 2 — Import VBA Modules

See `IMPLEMENTATION_GUIDE.md`, Section 2 for the full procedure. Summary:

1. Press **Alt+F11** to open the VBA Editor.
2. Right-click the project → **Import File...**
3. Select all 29 `.bas` files from `03_Code/VBA/`.
4. Press **Ctrl+S** to save.

### Step 3 — Build the Command Center

See `IMPLEMENTATION_GUIDE.md`, Section 3. Summary:

- **Mode A (automatic):** In Immediate Window, type `modFormBuilder.BuildCommandCenter`
- **Mode B (manual):** Insert UserForm named `frmCommandCenter`, paste code from `frmCommandCenter_code.txt`

### Step 4 — Verify

Press **Ctrl+Shift+M**. The Command Center should appear with 50 actions.

---

## Workbook Structure

The workbook ships with 13 sheets. Do not rename, delete, or reorder them.

| Sheet | Purpose | Editable? |
|-------|---------|-----------|
| CrossfireHiddenWorksheet | Raw GL data (510 transactions) | Import only (Command 17) |
| Assumptions | Allocation shares, drivers, parameters | Yes — update monthly |
| Data Dictionary | Column and field definitions | Reference only |
| AWS Allocation | AWS cost allocation by product | Formula-driven |
| Report--> | Navigation hub / table of contents | Auto-generated |
| P&L - Monthly Trend | 12-month P&L with formulas | Formula-driven |
| Product Line Summary | Revenue/cost by product line | Formula-driven |
| Functional P&L - Monthly Trend | Department-level 12-month trend | Formula-driven |
| Functional P&L Summary - Jan 25 | January detail by department | Formula-driven |
| Functional P&L Summary - Feb 25 | February detail | Formula-driven |
| Functional P&L Summary - Mar 25 | March detail (template for cloning) | Formula-driven |
| US January 2025 Natural P&L | Natural classification P&L | Formula-driven |
| Checks | Reconciliation PASS/FAIL results | Formula-driven |

### Layout Contract

Every report sheet follows this layout:
- **Row 1:** Company title ("Keystone BenefitTech, Inc.") in column A
- **Row 4:** Column headers
- **Row 5+:** Data

Exceptions:
- **Assumptions:** Headers row 5, data row 6+
- **GL (CrossfireHidden):** Headers row 1, data row 2+
- **Checks:** Headers row 4, data row 5+, status in column E

---

## Data Integrity Notes

- All formulas reference the GL sheet as their data source. Do not edit formula cells on report sheets directly.
- The Assumptions sheet is the only sheet that should be manually edited by users.
- If you need to modify GL data, use Command 17 (Import) rather than editing cells directly.
- Save a snapshot (Command 20) before any significant changes.

---

## File Size Management

As the toolkit generates additional sheets (Dashboard, Variance Analysis, etc.), the workbook grows. Periodically:

1. Delete unused generated sheets (they can be regenerated anytime)
2. Clear the audit log (Command 43, after exporting with Command 42)
3. Use **File → Info → Check for Issues → Inspect Document** to remove hidden metadata

---

## Backup Recommendations

- Keep the original `.xlsx` as an unmodified backup
- Save `.xlsm` snapshots before and after each monthly close
- Export the audit log monthly (Command 42)
- Store backups in a separate location from the working copy
