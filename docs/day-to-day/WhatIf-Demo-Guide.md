# What-If Scenario Demo — Command Center Guide

**Module:** `modWhatIf_v2.1.bas`
**Command Center Actions:** #63, #64, #65 (What-If Demo category)
**Last Updated:** 2026-03-12

---

## What This Does

The What-If Demo lets you change Assumptions drivers (revenue, expenses, AWS costs, headcount) by a percentage, instantly recalculate the entire P&L model, and see a styled impact report. When you're done, one click restores everything back to the original values.

This is designed for the live CFO/CEO demo — show them "What if revenue drops 15%?" and the entire model updates in real time.

---

## Prerequisites

- All 39 VBA modules imported into the Excel demo file
- Command Center rebuilt after importing (run `BuildCommandCenter` from the Immediate Window)
- The **Assumptions** sheet must exist in the workbook with driver names in column A and values in column B

---

## Action #63 — Run What-If Scenario Demo

### How to Run

1. Press **Ctrl+Shift+M** to open the Command Center
2. Click the **What-If Demo** category on the left
3. Select **"Run What-If Scenario Demo"** (action #63)
4. Click **Run Selected** (or double-click the action)

**Or from the Immediate Window:**
```
modWhatIf.RunWhatIfDemo
```

### What Happens

An InputBox pops up with 9 options:

| # | Scenario | What It Changes |
|---|----------|----------------|
| 1 | Revenue drops 15% | All revenue-related drivers on Assumptions sheet get multiplied by 0.85 |
| 2 | Revenue increases 10% | All revenue-related drivers get multiplied by 1.10 |
| 3 | AWS costs increase 25% | All AWS/cloud/compute drivers get multiplied by 1.25 |
| 4 | Headcount grows 20% | All headcount/salary/FTE/compensation/payroll drivers get multiplied by 1.20 |
| 5 | All expenses cut 10% | All non-revenue drivers get multiplied by 0.90 |
| 6 | Best case: Revenue +15%, Expenses -5% | Combo — revenue drivers up 15%, expense drivers down 5% |
| 7 | Worst case: Revenue -20%, Expenses +15% | Combo — revenue drivers down 20%, expense drivers up 15% |
| 8 | Custom | Jumps to action #64 (Custom What-If — you pick the driver and %) |
| 9 | Restore | Jumps to action #65 (Restore Baseline) |

### What You Should See

1. **First run:** A hidden sheet called **"WhatIf_Baseline"** is created automatically — this saves your original Assumptions values so they can be restored later
2. The Assumptions sheet values are **actually changed** — the model recalculates
3. A new sheet called **"What-If Impact"** is created with:
   - Title: "Keystone BenefitTech, Inc."
   - Subtitle: "What-If Impact Analysis: [scenario name]"
   - Timestamp
   - A styled table showing: **Driver | Original Value | New Value | Change | Change %**
   - Green font = value went up, Red font = value went down
   - Alternating row shading
   - "NEXT STEPS" section at the bottom
4. A MsgBox confirms the scenario was applied and how many drivers changed
5. You land on the "What-If Impact" sheet automatically

### For the Demo (Recommended Flow)

1. Start on the Assumptions sheet so the audience can see the current values
2. Run action #63, pick option **1** (Revenue drops 15%)
3. Walk through the impact report — point out the red numbers
4. Switch to **P&L Monthly Trend** sheet to show how the full P&L changed
5. Then run action #65 (Restore Baseline) to reset everything
6. Run action #63 again, pick option **6** (Best Case) to show a positive scenario
7. Restore again when done

This makes a great "wow" moment in the demo — the CFO asks "What if revenue drops?" and you answer in 3 seconds.

---

## Action #64 — Custom What-If Analysis

### How to Run

1. Press **Ctrl+Shift+M** to open the Command Center
2. Click the **What-If Demo** category
3. Select **"Custom What-If Analysis"** (action #64)
4. Click **Run Selected**

**Or from the Immediate Window:**
```
modWhatIf.QuickWhatIf
```

### What Happens

1. An InputBox shows **all drivers** from the Assumptions sheet with their current values (numbered list)
2. You type the number of the driver you want to change
3. A second InputBox asks for the percentage change:
   - Type `10` for +10%
   - Type `-15` for -15%
   - Type `25` for +25%
4. The baseline is saved, the driver value is changed, and the impact report is built — same as action #63

### When to Use This

- When someone in the audience asks about a specific driver that isn't in the 7 presets
- Example: "What if our licensing revenue goes up 20%?" — pick the licensing driver, type 20

---

## Action #65 — Restore Baseline (Undo What-If)

### How to Run

1. Press **Ctrl+Shift+M** to open the Command Center
2. Click the **What-If Demo** category
3. Select **"Restore Baseline (Undo What-If)"** (action #65)
4. Click **Run Selected**

**Or from the Immediate Window:**
```
modWhatIf.RestoreBaseline
```

### What Happens

1. A Yes/No confirmation box appears: "Restore original Assumptions values?"
2. Click **Yes**
3. All Assumptions drivers are set back to their original values (from the hidden WhatIf_Baseline sheet)
4. The **"WhatIf_Baseline"** and **"What-If Impact"** sheets are both deleted
5. The model recalculates
6. You land on the Assumptions sheet
7. A MsgBox confirms how many drivers were restored

### Important Notes

- You **must** run a What-If scenario first — if no baseline exists, you'll see "No baseline saved"
- The baseline only saves once per session — if you run multiple scenarios without restoring, the baseline still holds the **original** values (not the intermediate ones), which is correct
- After restoring, you can run another What-If scenario — it will save a fresh baseline

---

## Troubleshooting

| Problem | Cause | Fix |
|---------|-------|-----|
| "Assumptions sheet not found" | Sheet name doesn't match | Check that the sheet is named exactly **"Assumptions"** (must match `SH_ASSUMPTIONS` in modConfig) |
| "No baseline saved" on restore | No What-If was run yet | Run a scenario first (action #63 or #64) |
| Impact report shows 0 changes | No drivers matched the category | Check that your Assumptions sheet has driver names containing keywords like "revenue", "aws", "headcount", etc. |
| Numbers didn't change back after restore | You may have run restore twice | The second restore deletes the baseline — values should already be back from the first restore |
| What-If Impact sheet already exists | Previous run wasn't restored | The tool auto-deletes the old one before creating a new one — this is normal |

---

## How the Category Matching Works (Behind the Scenes)

The 7 presets match drivers by scanning their names for keywords:

| Category | Keywords Matched |
|----------|-----------------|
| Revenue ("rev") | revenue, rev share, sales, growth+revenue |
| AWS ("aws") | aws, cloud, compute |
| Headcount ("head") | headcount, salary, fte, compensation, payroll |
| Expense ("expense") | Everything that is NOT revenue/rev share/sales |

The "expense" category is a catch-all — if a driver doesn't match revenue, it's treated as an expense. This means options 5, 6, and 7 will change most non-revenue drivers.
