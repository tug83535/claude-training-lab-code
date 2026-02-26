# KBT P&L Toolkit — Implementation Guide

> **Audience:** IT support, power users, or anyone performing first-time setup.
> Covers Trust Center, module import, UserForm build, Python setup, named ranges, and fiscal year rollover.

---

## 1. Trust Center Macro Settings

Excel macros are disabled by default. You must enable them before the toolkit will function.

### Step-by-Step (Excel Desktop, Windows)

1. Open Excel (any blank workbook).
2. Navigate: **File → Options → Trust Center → Trust Center Settings...**
3. In the left panel, click **Macro Settings**.
4. Select one of:
   - **Enable all macros** (recommended for development/internal use)
   - **Disable all macros with notification** (shows a yellow bar each time you open)
5. Check the box: **"Trust access to the VBA project object model"**
   - This is required for Mode A automatic UserForm building.
   - If your IT policy blocks this, use Mode B (manual install) instead.
6. Click **OK**, then **OK** again to close Options.
7. **Restart Excel** for changes to take effect.

### Trusted Locations (Alternative)

If your organization uses Trusted Locations instead of macro settings:

1. In Trust Center Settings, click **Trusted Locations**.
2. Click **Add new location...**
3. Browse to the folder containing the workbook.
4. Check **"Subfolders of this location are also trusted"**.
5. Click **OK**.

---

## 2. Importing VBA Modules

All 29 `.bas` module files are in the `03_Code/VBA/` folder.

### Import Procedure

1. Open `KeystoneBenefitTech_PL_Model.xlsm`.
2. Press **Alt+F11** to open the VBA Editor.
3. In the **Project Explorer** (left panel), find your workbook under "VBAProject".
4. Right-click on **Modules** → **Import File...**
5. Navigate to `03_Code/VBA/`, select all `.bas` files, click **Open**.
6. Verify all 29 modules appear under the Modules folder.
7. Press **Alt+Q** to return to Excel.
8. Press **Ctrl+S** to save (choose `.xlsm` format if prompted).

### Module Inventory (29 files)

**Foundation (4):**
modConfig, modPerformance, modLogger, ThisWorkbook (events)

**Core Features (15):**
modNavigation, modMonthlyTabGenerator, modReconciliation, modDataQuality,
modVarianceAnalysis, modSensitivity, modDashboard, modPDFExport,
modAWSRecompute, modMasterMenu, modValidation, modSnapshot,
modConditionalFormat, modEmailSummary, modSearch

**Advanced (10):**
modAdmin, modAllocation, modForecast, modFormatting, modFormBuilder,
modImport, modIntegrationTest, modRefresh, modScenario, modSetup

### Verifying the Import

After importing, press **Alt+F11** and check:
- Each module starts with `Option Explicit`
- The Immediate Window (Ctrl+G) shows no compile errors
- Press **Alt+Q**, then **Ctrl+Shift+M** — the Command Center should open

---

## 3. Building the frmCommandCenter UserForm

The Command Center is a UserForm that provides point-and-click access to all 50 toolkit actions. There are two ways to create it.

### Mode A — Automatic Build (Recommended)

**Prerequisite:** "Trust access to the VBA project object model" must be enabled (Step 1.5 above).

1. Open the VBA Editor (**Alt+F11**).
2. Open the Immediate Window (**Ctrl+G**).
3. Type: `modFormBuilder.BuildCommandCenter` and press **Enter**.
4. You should see a success message in the Immediate Window.
5. Verify: a new form called `frmCommandCenter` appears under "Forms" in Project Explorer.
6. Close VBA Editor (**Alt+Q**).
7. Test: press **Ctrl+Shift+M** — the full UserForm should appear.

### Mode B — Manual Install

If Trust Access is blocked or Mode A fails:

1. In the VBA Editor, right-click your VBAProject → **Insert → UserForm**.
2. In the Properties panel (bottom-left), set:
   - **Name:** `frmCommandCenter`
   - **Caption:** `KBT P&L Toolkit — Command Center`
   - **Width:** 620
   - **Height:** 480
3. Double-click the blank form to open its code window.
4. Open `frmCommandCenter_code.txt` from the `03_Code/VBA/` folder.
5. Copy the entire contents and paste into the form's code window.
6. Press **Ctrl+S** to save.
7. Close VBA Editor, then press **Ctrl+Shift+M** to test.

### InputBox Fallback

If the UserForm is not installed, pressing Ctrl+Shift+M will fall back to a 3-page InputBox menu. This provides the same 50 actions but without the search and category filtering. Type the action number and press OK.

---

## 4. Python Environment Setup (Optional)

The Python analytics suite runs externally alongside the workbook. It is entirely optional — all core functionality works through VBA alone.

### Requirements

- Python 3.9+ (3.11 or 3.12 recommended)
- pip (included with Python)

### Setup

```bash
# Navigate to the Python folder
cd 03_Code/Python/

# Install dependencies
pip install -r requirements.txt

# Verify installation
python pnl_config.py

# You should see:
#   KBT P&L Toolkit v2.1.0 — Configuration Module
#   Source File:      KeystoneBenefitTech_PL_Model.xlsx
#   Fiscal Year:      FY2025
#   Products:         iGO, Affirm, InsureSight, DocFast
#   Revenue shares sum:   1.00 ✓
#   AWS compute shares:   1.00 ✓
#   Headcount shares:     1.00 ✓
```

### Available Python Commands

```bash
python pnl_runner.py --help           # Show all commands
python pnl_runner.py dashboard        # Interactive Streamlit dashboard
python pnl_runner.py month-end        # Month-end close checklist
python pnl_runner.py forecast         # Statistical forecasting
python pnl_runner.py allocate         # Allocation what-if simulator
python pnl_runner.py snapshot list    # List saved snapshots
python pnl_runner.py match            # AP fuzzy matching
python pnl_runner.py email            # Generate executive email report
python pnl_runner.py test             # Run automated test suite
python pnl_runner.py config           # Show current configuration
```

### File Path Configuration

By default, Python scripts look for `KeystoneBenefitTech_PL_Model.xlsx` in the current directory. To use a different path:

```bash
# Option 1: CLI argument
python pnl_runner.py month-end --file /path/to/your/workbook.xlsx

# Option 2: Environment variable
export KBT_SOURCE_FILE=/path/to/your/workbook.xlsx
python pnl_runner.py month-end
```

---

## 5. Named Range Setup

Dynamic named ranges auto-expand as new data is added. Run this once after initial setup.

1. Open the Command Center (**Ctrl+Shift+M**).
2. This is handled by `modSetup.DynamicNamedRanges` — accessible from the VBA Editor Immediate Window:
   ```
   modSetup.DynamicNamedRanges
   ```
3. This creates OFFSET/COUNTA-based named ranges for key data areas.

### What Gets Created

| Named Range | Scope | Points To |
|-------------|-------|-----------|
| `rng_GL_Data` | Workbook | GL data area (auto-expanding) |
| `rng_Assumptions` | Workbook | Assumptions data rows |
| `rng_PLTrend_Data` | Workbook | P&L Trend data rows |
| `rng_Checks` | Workbook | Checks sheet data rows |

---

## 6. Fiscal Year Rollover

At the start of each fiscal year, update two locations:

### VBA (modConfig)

Open **Alt+F11 → modConfig** and change:

```vba
' Update these 3 constants:
Public Const FISCAL_YEAR    As String = "26"       ' was "25"
Public Const FISCAL_YEAR_4  As String = "2026"     ' was "2025"
Public Const FY_LABEL       As String = "FY2026"   ' was "FY2025"
```

### Python (pnl_config.py)

Open `pnl_config.py` and change:

```python
# Update these 3 constants:
FISCAL_YEAR = "26"          # was "25"
FISCAL_YEAR_4 = "2026"      # was "2025"
FY_LABEL = "FY2026"         # was "FY2025"
```

### Post-Rollover Steps

1. Update sheet tab names (e.g., "Functional P&L Summary - Jan 26").
2. Run `modMonthlyTabGenerator.GenerateMonthlyTabs` to create new monthly tabs.
3. Run `modReconciliation.RunAllChecks` to verify the new year's structure.
4. Save a snapshot: Command #20 (Save Current Scenario) with name "FY26_Opening".

---

## 7. Architecture Overview

```
┌──────────────────────────────────────────────────────────────┐
│                    USER INTERFACE LAYER                       │
│  ┌────────────────┐  ┌───────────────┐  ┌────────────────┐  │
│  │ frmCommandCenter│  │ InputBox      │  │ Keyboard       │  │
│  │ (50 actions)   │  │ (fallback)    │  │ Shortcuts      │  │
│  └───────┬────────┘  └──────┬────────┘  └───────┬────────┘  │
│          └──────────────────┼───────────────────┘            │
│                             ▼                                │
│              modFormBuilder.ExecuteAction(n)                  │
│              modMasterMenu.ExecuteMenuAction(n)               │
├──────────────────────────────────────────────────────────────┤
│                    FEATURE MODULES (15+10)                    │
│  Monthly Ops  │ Analysis    │ Reporting   │ Data Quality     │
│  Navigation   │ Sensitivity │ PDF Export  │ Dashboard        │
│  Import       │ Forecast    │ Scenarios   │ Allocation       │
│  Search       │ Snapshot    │ Email       │ Conditional Fmt  │
├──────────────────────────────────────────────────────────────┤
│                    FOUNDATION LAYER                           │
│  modConfig (constants, helpers) │ modPerformance (TurboOn/Off)│
│  modLogger (audit trail)        │ ThisWorkbook (events)       │
├──────────────────────────────────────────────────────────────┤
│                    DATA LAYER                                │
│  KeystoneBenefitTech_PL_Model.xlsx                           │
│  13 sheets: GL, Assumptions, P&L Trend, Product Summary,     │
│  Functional Trend, Functional Jan/Feb/Mar, AWS, Checks, etc. │
├──────────────────────────────────────────────────────────────┤
│              PYTHON ANALYTICS (Optional, External)            │
│  pnl_runner.py → dashboard │ month-end │ forecast │ allocate │
│                   snapshot  │ match     │ email    │ test     │
└──────────────────────────────────────────────────────────────┘
```

---

## 8. Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| "Compile error: Variable not defined" | `Option Explicit` + missing declaration | Reimport the module — the v2.1 files have all declarations |
| "Sub or Function not defined" | Missing module | Check that all 29 `.bas` files are imported |
| "Subscript out of range" on sheet access | Sheet name doesn't match constant | Verify tab names match `modConfig` constants exactly |
| "Programmatic access to VBA project" error | Trust Access not enabled | Enable in Trust Center (Step 1.5) or use Mode B manual install |
| UserForm controls misaligned | Screen DPI scaling | The form is built for 100% DPI; adjust `Width`/`Height` in Properties if needed |
| Python "ModuleNotFoundError" | Missing pip package | Run `pip install -r requirements.txt` |
| Python "FileNotFoundError" | Excel file not in current directory | Use `--file` argument or set `KBT_SOURCE_FILE` env var |
| PDF export creates blank pages | Print area not set | The export module sets print areas automatically — ensure data exists |
| Charts show "#REF!" | Deleted or renamed sheets | Regenerate charts via Command #12 after fixing sheet names |
