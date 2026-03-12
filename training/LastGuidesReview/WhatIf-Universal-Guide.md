# What-If Scenario Tool — Universal Toolkit Guide

**Module:** `modUTL_WhatIf.bas`
**Location:** `UniversalToolsForAllFiles/vba/`
**Last Updated:** 2026-03-12

---

## What This Does

The Universal What-If tool lets you apply a percentage change to **any selected cells** in **any Excel file**. It saves a backup of the original values, applies the change, builds a styled impact report, and lets you undo everything with one click.

Unlike the demo version (which only works with the P&L demo file's Assumptions sheet), this works on literally any spreadsheet — budgets, forecasts, pricing models, headcount plans, anything with numbers.

---

## Prerequisites

- Import `modUTL_WhatIf.bas` into any Excel workbook (Alt+F11 > File > Import File)
- That's it — zero dependencies, no other modules needed

---

## Tool #1 — RunWhatIfPresets (Quick Preset Menu)

### How to Run

**From the Immediate Window (Ctrl+G in VBE):**
```
modUTL_WhatIf.RunWhatIfPresets
```

**Or from Alt+F8 (Macros dialog):**
Select `RunWhatIfPresets` and click Run.

### Step-by-Step

1. **First:** Select the cells you want to change
   - Click a single cell, or drag to select a range
   - You can select an entire column of numbers, a row, or scattered cells (Ctrl+Click)
   - Only numeric cells will be changed — text, blanks, and formulas are skipped... wait, formulas ARE included if they contain a number. Best practice: select cells with hardcoded numbers
2. **Run the macro** using one of the methods above
3. **Pick a preset** from the InputBox:

   | # | Preset |
   |---|--------|
   | 1 | Increase 5% |
   | 2 | Increase 10% |
   | 3 | Increase 25% |
   | 4 | Decrease 5% |
   | 5 | Decrease 10% |
   | 6 | Decrease 25% |
   | 7 | Custom % (you type the number) |

4. **Confirm** the change — a dialog shows:
   - How many numeric cells will be changed
   - Which sheet they're on
   - The cell range
5. Click **Yes** to proceed

### What You Should See

1. A new sheet called **"UTL_WhatIf_Impact"** appears with:
   - Title: "What-If Impact Report"
   - Scenario description (e.g., "+10% applied to 12 cell(s)")
   - Source sheet name and timestamp
   - A styled table with columns: **Label | Cell | Original Value | New Value | Change | Change %**
   - Green font for increases, red font for decreases
   - Alternating row shading (light gray every other row)
   - **TOTALS row** at the bottom (navy background, white text) — shows total original, total new, total change
   - "HOW TO UNDO" instructions in red at the bottom
2. The selected cells on your original sheet now have the new values
3. A hidden backup sheet (**"UTL_WhatIf_Backup"**) is created with all original values
4. A MsgBox confirms the change

### Where Do the Labels Come From?

The tool automatically grabs the label from **column A** of each changed cell's row. For example, if you change cell D5, it looks at A5 for the label. If column A is empty, it shows "Row 5" instead.

---

## Tool #2 — RunWhatIf (Direct Custom %)

### How to Run

```
modUTL_WhatIf.RunWhatIf
```

### Step-by-Step

1. **Select your cells first** (same as above)
2. Run the macro
3. Type your percentage:
   - `10` = increase by 10%
   - `-15` = decrease by 15%
   - `25` = increase by 25%
   - `-50` = cut in half
4. Confirm and see the same impact report

This is the same as picking option 7 (Custom) from the presets menu — just a shortcut if you already know the % you want.

---

## Tool #3 — RestoreBaseline (Undo Everything)

### How to Run

```
modUTL_WhatIf.RestoreBaseline
```

### Step-by-Step

1. Run the macro
2. A Yes/No confirmation appears: "Restore all original values from the last What-If?"
3. Click **Yes**
4. All changed cells are set back to their original values
5. The backup sheet (**UTL_WhatIf_Backup**) and impact report sheet (**UTL_WhatIf_Impact**) are both deleted
6. The workbook recalculates
7. A MsgBox confirms how many cells were restored

### Important Notes

- You **must** run a What-If first — if no backup exists, you'll see "No baseline saved"
- The backup stores the **sheet name + cell address + original value** for every cell that was changed
- After restoring, you can run another What-If — it creates a fresh backup each time
- If you run a What-If twice without restoring, the old backup is **replaced** with the current values (not the original originals). So always restore before running a new scenario if you want to get back to the true starting point

---

## Tool #4 — ViewBaseline (Inspect the Backup)

### How to Run

```
modUTL_WhatIf.ViewBaseline
```

### What It Does

- Makes the hidden backup sheet visible so you can see what values are saved
- The sheet has 4 columns: **Sheet | Cell Address | Original Value | Label**
- Useful if you want to verify what the restore will do before running it
- The sheet stays visible until you restore (which deletes it) or manually hide it

---

## Example Walkthrough

**Scenario:** You have a budget spreadsheet. Column B has annual budget amounts for 15 departments. You want to see what happens if every department gets a 10% cut.

1. Open your budget file
2. Import `modUTL_WhatIf.bas` (Alt+F11 > File > Import File > select the .bas file)
3. Go back to your budget sheet
4. Select cells **B2:B16** (the budget amounts)
5. Press **Alt+F8**, select `RunWhatIfPresets`, click Run
6. Pick **5** (Decrease 10%)
7. Click **Yes** to confirm
8. The impact report shows every department's old budget, new budget, and the dollar change
9. The totals row at the bottom shows the total budget reduction
10. When done, press **Alt+F8**, select `RestoreBaseline`, click Run
11. Click **Yes** — all 15 cells go back to their original values

---

## Troubleshooting

| Problem | Cause | Fix |
|---------|-------|-----|
| "Select the cells you want to change first" | Nothing was selected before running | Go to your sheet, select the cells with numbers, then run the macro |
| "No numeric values found in selection" | You selected text cells or empty cells | Make sure your selection contains actual numbers (not text that looks like numbers) |
| "No baseline saved" on restore | No What-If was run yet | Run a What-If scenario first |
| Labels show "Row 5" instead of real names | Column A is empty for that row | This is normal — the tool tries column A first, falls back to row number |
| Impact report has wrong labels | Column A has something other than the row label | This is cosmetic only — the actual values and changes are always correct |
| "Enter a number" error on custom % | You typed text instead of a number | Type just the number, no % sign. Example: type `15` not `15%` |

---

## Difference from the Demo Version

| Feature | Demo (modWhatIf) | Universal (modUTL_WhatIf) |
|---------|------------------|---------------------------|
| Works on | Only the P&L demo file | Any Excel file |
| What it changes | Assumptions sheet drivers | Any cells you select |
| Presets | 7 P&L-specific scenarios (Revenue, AWS, Headcount, etc.) | 6 generic % presets (+/-5/10/25%) |
| Category matching | Scans driver names for keywords | N/A — you pick the cells |
| Dependencies | modConfig, modPerformance, modLogger | None |
| Impact report sheet | "What-If Impact" | "UTL_WhatIf_Impact" |
| Backup sheet | "WhatIf_Baseline" | "UTL_WhatIf_Backup" |
| In Command Center | Yes (actions 63-65) | No (run from Alt+F8 or Immediate Window) |
