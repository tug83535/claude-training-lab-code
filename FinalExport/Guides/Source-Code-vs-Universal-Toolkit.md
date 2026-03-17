# Source Code vs. Universal Toolkit -- What's the Difference?

**A quick guide so you grab the right files**

---

## The Short Version

| | Source Code | Universal Toolkit |
|---|---|---|
| **Works with** | The P&L Demo File only | ANY Excel file you have |
| **Where to find it** | Source Code folder | Universal Toolkit folder |
| **Do I need it to use the demo file?** | No -- it's already built into the .xlsm | No -- these are separate tools |
| **Why would I want it?** | To read the code, learn from it, or adapt pieces for your own projects | To add automation tools to your own Excel files |

---

## Source Code -- The Demo File's Engine

The **Source Code** folder contains the raw code that powers the P&L Demo File. Think of it like looking under the hood of a car.

**What's in it:**
- **39 VBA modules** -- these are the macros that run the 65 Command Center actions (data quality scans, variance analysis, dashboards, reconciliation, etc.)
- **13 Python scripts** -- standalone tools for forecasting, Monte Carlo simulation, data cleanup, and more
- **4 SQL scripts** -- database queries for staging, transforming, and validating financial data

**Important:** You do NOT need to download these files to use the demo file. The VBA code is already inside the .xlsm file. When you open the demo file and press Ctrl+Shift+M, everything just works.

**So why is the source code here?** Two reasons:

1. **Learning.** If you want to understand HOW a macro works -- maybe you want to build something similar for your own file -- you can open the .bas file and read the code. It's like having the recipe, not just the meal.

2. **CoPilot / AI adaptation.** The CoPilot Prompt Guide (in the Quick Reference folder) teaches you how to take a piece of this code, paste it into CoPilot or Claude, and say "adapt this for my file." The source code gives you the raw material to work with.

**Example:** You love how the Variance Commentary macro writes plain-English explanations of budget variances. You open `modVarianceAnalysis_v2.1.bas` from Source Code, copy the `GenerateCommentary` function, paste it into CoPilot, and say: "Adapt this macro to work on my department's budget spreadsheet. My data starts in row 3 and my months are in columns D through O."

---

## Universal Toolkit -- Tools for Any File

The **Universal Toolkit** folder contains standalone tools that work on ANY Excel file -- not just the demo. Think of it like a toolbox you can carry to any job.

**What's in it:**
- **23 VBA modules** with 140+ tools organized by category:
  - Data Sanitizer (fix text-stored numbers, floating point errors, blank rows)
  - Highlights (color cells by threshold, find duplicates, top/bottom N)
  - Tab Organizer (sort, color, group, rename, reorder tabs)
  - Column Ops (split, merge, move, swap columns)
  - Compare (side-by-side sheet comparison with color-coded diff report)
  - Consolidate (merge data from multiple sheets with source tracking)
  - Pivot Tools (create, refresh, style pivot tables)
  - Lookup Builder (auto-generate VLOOKUP and INDEX-MATCH formulas)
  - Validation Builder (create dropdown lists, number ranges, date rules)
  - Comments (extract, count, clear, convert all comments)
  - Branding (apply iPipeline brand colors and fonts to any sheet)
  - And more
- **22 Python scripts** for file comparison, PDF table extraction, data cleanup, reconciliation, and more

**How to use the VBA tools:**
1. Open YOUR Excel file (any file -- a budget, a report, a data dump, anything)
2. Press Alt+F11 to open the VBA Editor
3. Go to File > Import File
4. Pick the .bas module you want (e.g., `modUTL_DataSanitizer.bas`)
5. Close the VBA Editor
6. Press Alt+F8 to see the available macros
7. Run the one you want

The **Universal Toolkit Guide** (included in the Universal Toolkit folder) walks through this step by step with screenshots.

---

## Side-by-Side Example

Let's say you want to clean up text-stored numbers (numbers Excel thinks are text -- they have the little green triangle in the corner).

**Using Source Code:** You open `modDataQuality_v2.1.bas`, find the `FixTextNumbers` sub, and adapt it for your file using CoPilot. This takes some work because the demo code references specific sheet names and settings that only exist in the demo file.

**Using Universal Toolkit:** You import `modUTL_DataSanitizer.bas` into your file, run `ConvertTextStoredNumbers`, and it just works. No adaptation needed. It scans whatever file you have open.

**Bottom line:** If you want a tool that works right now on your own file, grab it from Universal Toolkit. If you want to learn how something works or build something custom, look at Source Code and use the CoPilot Prompt Guide.

---

## Still Not Sure?

| I want to... | Go to... |
|---|---|
| Use the demo file as-is | Just download the .xlsm from Demo File -- you're done |
| Clean up my own messy spreadsheet | Universal Toolkit -- import the tool you need |
| Learn how a specific macro works | Source Code -- read the .bas file |
| Build a custom macro for my team's file | Source Code + CoPilot Prompt Guide |
| Compare two Excel files for differences | Universal Toolkit > Python Scripts > compare_files.py |
| Extract tables from a PDF into Excel | Universal Toolkit > Python Scripts > pdf_extractor.py |

---

*Questions? Ask Connor or check the CoPilot Prompt Guide for AI-assisted help.*

*iPipeline Finance & Accounting -- 2026*
