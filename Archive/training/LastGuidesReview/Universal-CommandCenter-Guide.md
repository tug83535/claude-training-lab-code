# Universal Command Center — User Guide

> **Audience:** Finance & Accounting staff who want a single menu to access all Universal Toolkit tools.
> **Module:** `modUTL_CommandCenter.bas`
> **Version:** 1.0 | **Date:** 2026-03-12

---

## What Is the Universal Command Center?

The Universal Command Center is a **single menu** that gives you access to every tool in the Universal Toolkit — plus any custom macros you add yourself. Instead of remembering which module has which tool, you open one menu, pick a category, and run the tool.

**Key features:**
- **81 built-in tools** organized into 13 categories
- **Auto-discovery** — if new `modUTL_*` modules are added, they appear automatically
- **Custom tool registry** — add your own macros to the menu
- **Keyword search** — find any tool by typing a word
- **Tool inventory sheet** — print the full list to a styled Excel sheet

---

## How to Open the Command Center

### Option A — Run the Macro Directly
1. Press **Alt+F8** (the Run Macro dialog)
2. Type `LaunchCommandCenter` in the search box
3. Click **Run**

### Option B — Assign a Keyboard Shortcut
1. Press **Alt+F8**
2. Select `modUTL_CommandCenter.LaunchCommandCenter`
3. Click **Options**
4. Assign a shortcut key (e.g., **Ctrl+Shift+U**)
5. Click **OK**, then **Cancel**
6. Now press **Ctrl+Shift+U** any time to open the Command Center

### Option C — Add a Button to the Ribbon or Quick Access Toolbar
1. Right-click the Ribbon → **Customize the Ribbon** (or **Customize Quick Access Toolbar**)
2. Under "Choose commands from," select **Macros**
3. Find `modUTL_CommandCenter.LaunchCommandCenter`
4. Click **Add** to add it to your toolbar
5. Click **OK**

---

## Using the Command Center

### Step 1 — Choose a Category

When you open the Command Center, you see a numbered list of categories:

```
Universal Command Center v1.0 (81 tools)

CATEGORIES:
----------------------------------------
  1. Audit (8 tools)
  2. Branding (2 tools)
  3. Data Cleaning (12 tools)
  4. Data Sanitizer (4 tools)
  5. Executive Brief (1 tool)
  6. Finance (14 tools)
  7. Formatting (9 tools)
  8. Progress Bar (3 tools)
  9. Sheet Tools (4 tools)
  10. Splash Screen (2 tools)
  11. What-If Analysis (4 tools)
  12. Workbook Management (15 tools)
  13. Command Center (3 tools)

  S. Search all tools by keyword
  L. List all tools to a sheet

Enter a number (or S/L):
```

**Type the number** of the category you want and press **Enter**.

### Step 2 — Choose a Tool

You see the tools in that category with descriptions:

```
Finance (14 tools)
----------------------------------------

  1. Duplicate Invoice Detector
     Find duplicate invoices by amount, date, vendor
  2. Auto-Balancing GL Validator
     Validate that debits equal credits in GL data
  3. Trial Balance Checker
     Verify trial balance totals to zero
  ...

Enter tool number to run (or 0 to go back):
```

**Type the number** of the tool you want and press **Enter**. The tool runs immediately.

**Type 0** to go back to the category menu.

### Step 3 — Done

The tool runs and shows its results. The Command Center closes after each tool runs. Open it again to run another tool.

---

## Searching for a Tool

If you don't know which category a tool is in:

1. Open the Command Center (Alt+F8 → `LaunchCommandCenter`)
2. Type **S** and press Enter
3. Type a keyword (e.g., `duplicate`, `format`, `PDF`, `audit`)
4. The search results show every tool matching your keyword
5. Type the number of the tool you want to run

**The search checks:** tool names, descriptions, categories, and macro names.

You can also run the search directly:
- **Alt+F8** → type `SearchTools` → **Run**

---

## Listing All Tools to a Sheet

Want to see every tool in a printable format?

1. Open the Command Center
2. Type **L** and press Enter
3. A new sheet called **UTL_ToolInventory** is created with:
   - Tool number, category, name, macro name, description, and source
   - iPipeline-branded header styling
   - Alternating row colors for readability

You can also run this directly:
- **Alt+F8** → type `ListAllTools` → **Run**

---

## Adding Your Own Macros (Custom Tools)

You can add **any macro** from any module in your workbook to the Command Center.

### Register a Custom Tool

1. Press **Alt+F8** → type `RegisterCustomTool` → **Run**
2. You are prompted for 4 things:
   - **Macro name** — The exact name of the macro (e.g., `MyModule.MyReport`)
   - **Display name** — What you want to see in the menu (e.g., `Weekly Status Report`)
   - **Category** — Which category to file it under (e.g., `My Tools`). Leave blank for "Custom Tools"
   - **Description** — A short description (optional)
3. The tool is saved and will appear in the Command Center from now on

**Important:** The macro must be a `Public Sub` in any module in your workbook. If the module name is unique, you can just type the sub name (e.g., `MyReport`). If multiple modules have the same sub name, use the full format: `ModuleName.SubName`.

### View Custom Tools

- **Alt+F8** → type `ViewCustomTools` → **Run**
- Shows all your registered custom tools with their details

### Remove a Custom Tool

- **Alt+F8** → type `RemoveCustomTool` → **Run**
- Shows a numbered list of your custom tools
- Type the number to remove and press Enter

### Where Are Custom Tools Saved?

Custom tools are saved on a hidden sheet called `UTL_CustomTools` in your workbook. This means:
- Your custom tools **persist** when you save and reopen the workbook
- They are **specific to this workbook** — other workbooks have their own custom tools
- The sheet is hidden but safe — do not delete it manually

---

## Auto-Discovery (Advanced)

If **Trust Access to the VBA Project Object Model** is enabled, the Command Center automatically scans for new `modUTL_*` modules and adds any Public Subs it finds that aren't already in the built-in registry.

### What This Means for You

- If someone gives you a new `modUTL_Something.bas` module, just import it
- The Command Center will find it and add its tools automatically
- No code changes needed — it just works

### How to Enable Trust Access (Optional)

1. Open Excel → **File** → **Options** → **Trust Center**
2. Click **Trust Center Settings**
3. Click **Macro Settings**
4. Check **Trust access to the VBA project object model**
5. Click **OK** → **OK**

**Note:** Trust Access is **optional**. The Command Center works perfectly fine without it — it just uses the built-in registry instead of scanning for new modules.

---

## Complete Tool Inventory

### Audit (8 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | External Link Finder | Find all external links in the workbook |
| 2 | Circular Reference Detector | Detect circular references across all sheets |
| 3 | Workbook Error Scanner | Scan for #REF, #VALUE, #N/A and other errors |
| 4 | Data Quality Scorecard | Generate a data quality score for the workbook |
| 5 | Named Range Auditor | Audit all named ranges for validity |
| 6 | Data Validation Checker | Check data validation rules across all sheets |
| 7 | Inconsistent Formulas Auditor | Find formulas that break row/column patterns |
| 8 | External Link Severance | Break and remove all external links safely |

### Branding (2 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Apply iPipeline Branding | Apply iPipeline brand colors and fonts to active sheet |
| 2 | Set iPipeline Theme Colors | Set workbook theme to iPipeline brand palette |

### Data Cleaning (12 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Unmerge and Fill Down | Unmerge cells and fill down with the merged value |
| 2 | Fill Blanks Down | Fill blank cells with the value above |
| 3 | Convert Text to Numbers | Convert text-stored numbers to real numbers |
| 4 | Remove Spaces | Remove leading and trailing spaces from all cells |
| 5 | Delete Blank Rows | Delete all completely blank rows |
| 6 | Replace Error Values | Replace #REF, #N/A, etc. with zero or blank |
| 7 | Highlight Duplicate Rows | Highlight duplicate rows in yellow |
| 8 | Remove Duplicate Rows | Remove duplicate rows (keeps first occurrence) |
| 9 | Multi-Replace Data Cleaner | Find and replace multiple values at once |
| 10 | Formula to Value | Convert formulas to their current values |
| 11 | Phantom Hyperlink Purger | Remove invisible/phantom hyperlinks |
| 12 | Numbers to Words | Convert numbers to written words |

### Data Sanitizer (4 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Run Full Sanitize | Fix text-numbers, floating-point tails, and integer formats |
| 2 | Preview Sanitize Changes | Dry-run report showing what would change (no edits) |
| 3 | Fix Floating-Point Tails | Fix floating-point noise (e.g., 10.000000001) |
| 4 | Convert Text-Stored Numbers | Convert text-stored numbers to real numbers (smart detection) |

### Executive Brief (1 tool)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Generate Exec Brief | Scan any workbook and build an executive brief report |

### Finance (14 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Duplicate Invoice Detector | Find duplicate invoices by amount, date, vendor |
| 2 | Auto-Balancing GL Validator | Validate that debits equal credits in GL data |
| 3 | Trial Balance Checker | Verify trial balance totals to zero |
| 4 | Journal Entry Validator | Validate journal entries for completeness and balance |
| 5 | Flux Analysis | Period-over-period flux analysis with thresholds |
| 6 | AP Aging Summary | Generate accounts payable aging summary |
| 7 | AR Aging Summary | Generate accounts receivable aging summary |
| 8 | Aging Bucket Calculator | Calculate aging buckets (Current, 30, 60, 90+) |
| 9 | Variance Analysis Template | Build a variance analysis template |
| 10 | Quick Corkscrew Builder | Build opening-additions-disposals-closing rollforward |
| 11 | Period Roll Forward | Roll financial periods forward |
| 12 | Multi-Currency Consolidation | Consolidate multi-currency data |
| 13 | Ratio Analysis Dashboard | Calculate and display key financial ratios |
| 14 | GL Journal Mapper | Map GL entries to journal categories |

### Formatting (9 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Auto-Fit All Columns/Rows | Auto-fit column widths and row heights |
| 2 | Freeze Top Row (All Sheets) | Freeze the top row on every sheet |
| 3 | Number Format Standardizer | Standardize number formats across selection |
| 4 | Currency Format Standardizer | Standardize currency formats |
| 5 | Date Format Standardizer | Standardize date formats across selection |
| 6 | Highlight Negatives Red | Highlight negative numbers in red |
| 7 | Financial Number Formatting | Apply financial number formatting |
| 8 | Conditional Format Purger | Remove all conditional formatting rules |
| 9 | Print Header/Footer Standardizer | Standardize print headers and footers |

### Progress Bar (3 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Start Progress | Initialize progress bar (for use in your own macros) |
| 2 | Update Progress | Update progress bar percentage |
| 3 | End Progress | Close progress bar |

### Sheet Tools (4 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | List All Sheets with Links | Create a sheet index with clickable hyperlinks |
| 2 | Template Cloner | Clone any sheet multiple times (1-50 copies) |
| 3 | Generate Unique Customer IDs | Fill blank cells with sequential unique IDs |
| 4 | Create Folders from Selection | Create Windows folders from selected cell values |

### Splash Screen (2 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Show Splash | Show a branded welcome splash screen |
| 2 | Show Custom Splash | Show a customizable splash screen |

### What-If Analysis (4 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Run What-If Presets | Apply preset percentage changes to selected cells |
| 2 | Run Custom What-If | Apply a custom percentage change to selected cells |
| 3 | Restore Baseline | Undo all What-If changes and restore original values |
| 4 | View Baseline | View the saved baseline values |

### Workbook Management (15 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Unhide All Sheets/Rows/Columns | Unhide all hidden sheets, rows, and columns |
| 2 | Export All Sheets (Combined PDF) | Export all sheets to a single PDF |
| 3 | Find & Replace (All Sheets) | Find and replace text across all sheets |
| 4 | Search All Sheets | Search for text across all sheets |
| 5 | Batch Sheet Renamer | Rename multiple sheets at once |
| 6 | Sort Sheets Alphabetically | Sort all worksheet tabs alphabetically |
| 7 | Create Table of Contents | Create a clickable table of contents sheet |
| 8 | Protect All Sheets | Protect all sheets with a password |
| 9 | Unprotect All Sheets | Unprotect all sheets |
| 10 | Lock All Formula Cells | Lock cells containing formulas |
| 11 | Export Active Sheet PDF | Export the active sheet to PDF |
| 12 | Export All Sheets (Individual PDFs) | Export each sheet to its own PDF |
| 13 | Reset All Filters | Clear all AutoFilter and Advanced Filter settings |
| 14 | Build Distribution-Ready Copy | Create a clean copy for distribution |
| 15 | Workbook Health Check | Run a comprehensive health check on the workbook |

### Command Center (3 tools)
| # | Tool | What It Does |
|---|------|-------------|
| 1 | Register Custom Tool | Add your own macro to the Command Center |
| 2 | Remove Custom Tool | Remove a custom-registered tool |
| 3 | View Custom Tools | View all custom-registered tools |

---

## Troubleshooting

### "Could not run macro" Error
**Cause:** The module containing the tool is not imported into this workbook.
**Fix:** Import the `.bas` file into the VBA Editor (Alt+F11 → File → Import File → select the `.bas` file).

### "No tools found in registry"
**Cause:** The module loaded but failed to build the registry (very rare).
**Fix:** Close and reopen the workbook. If the issue persists, re-import `modUTL_CommandCenter.bas`.

### Custom tools are gone after reopening
**Cause:** The hidden `UTL_CustomTools` sheet was deleted, or the workbook was not saved after registering tools.
**Fix:** Always save the workbook after registering custom tools. The sheet is hidden — do not unhide and delete it.

### Auto-discovered tools show "(Discovered)" category
**Cause:** Trust Access is enabled and the Command Center found new modules that aren't in the built-in registry.
**This is normal.** The tools work fine. The "(Discovered)" label just means they were found automatically.

### I added a new modUTL module but it doesn't appear
**Cause:** Trust Access is not enabled, so auto-discovery is off.
**Fix:** Either enable Trust Access (see the Auto-Discovery section above) or use `RegisterCustomTool` to manually add the tools.

---

## Quick Reference

| Action | How to Do It |
|--------|-------------|
| Open Command Center | Alt+F8 → `LaunchCommandCenter` → Run |
| Search for a tool | Alt+F8 → `SearchTools` → Run |
| Print tool inventory | Alt+F8 → `ListAllTools` → Run |
| Add your own macro | Alt+F8 → `RegisterCustomTool` → Run |
| Remove a custom macro | Alt+F8 → `RemoveCustomTool` → Run |
| View custom macros | Alt+F8 → `ViewCustomTools` → Run |
