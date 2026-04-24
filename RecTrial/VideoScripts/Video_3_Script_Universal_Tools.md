# Video 3 Script — "Universal Tools"

**Runtime Target:** 8:00–10:00
**Format:** Screen recording, no webcam, voice-over narration, chaptered
**Audience:** Anyone at iPipeline who uses Excel and wants to automate their own work
**Purpose:** Show examples of the universal tools running on a plain sample file, then point people to SharePoint
**Key principle:** Uses a separate, simple sample Excel file — NOT the demo file
**Music:** Subtle corporate/tech track during title card, chapter cards, and closing card only

---

## Pre-Recording Checklist

Before hitting record, make sure:

- [ ] Excel is the only application open
- [ ] Desktop is clean — no icons, taskbar auto-hidden
- [ ] Excel is maximized to full screen
- [ ] Zoom level set to 100% or 110% (SAME as Videos 1 and 2 — consistency across all three)
- [ ] Windows display scaling set to 100%
- [ ] All notifications silenced
- [ ] Sample file (Sample_Quarterly_Report.xlsx) is open and ready
- [ ] The sample file has some intentional "mess" baked in:
  - [ ] A few merged cells in column A
  - [ ] Some text-stored numbers (numbers formatted as text)
  - [ ] A few blank rows scattered in the data
  - [ ] Some extra spaces in text cells
  - [ ] At least one column of dates in mixed formats
  - [ ] A few error values (#N/A, #REF!) in formulas
  - [ ] At least one hidden sheet
  - [ ] Some unstyled headers (no formatting)
- [ ] Script visible on second monitor or printed
- [ ] Audio test done
- [ ] Screen recording running, 1920×1080, 30fps

**IMPORTANT:** The sample file should look generic and relatable — something any iPipeline employee might have on their desktop. Name it something like "Sample_Quarterly_Report.xlsx" or "Team_Data_Export.xlsx". Use realistic-looking but fictional data (employee names, departments, amounts, dates).

---

## Title Card (5 seconds)

**On screen:**

```
iPipeline Finance Automation
Universal Tools — For Any Excel File
```

**Audio:** Brief music sting (2–3 seconds), fade to silence
**Style:** iPipeline Blue (#0B4779) background, white text, Arial font

---

## Opening

**Duration:** 0:05–0:50 (45 seconds)
**On screen:** The sample file is open, showing a typical messy spreadsheet

### Script:

> "The P&L demo file showed what automation can do for a specific workflow. But the code library behind it includes over 75 VBA tools and 22 Python scripts that work on any Excel file — not just that one.
>
> In this video, I'm going to show you a handful of those tools running on a regular spreadsheet. Nothing special about this file — it's just a typical data export with some common problems. Messy formatting, text stored as numbers, blank rows, inconsistent dates — the kind of thing you deal with every day.
>
> Every tool you see here is available on SharePoint. Grab what you need, use the step-by-step guides, and if you get stuck, there are pre-built Copilot prompts to help."

### Screen Actions:
- File is open showing the sample data
- Slow scroll through the file so the viewer can see the "mess" — blank rows, mixed formatting, errors visible
- No clicking yet — just visual context

### Production Notes:
- "Over 75 VBA tools and 22 Python scripts" — say this clearly. The scale is impressive.
- The scroll should make the viewer feel the familiar pain of a messy spreadsheet. Don't rush it.
- Keep this under 45 seconds. You're setting context, not explaining.

---

## CHAPTER CARD: Chapter 1

**On screen (3 seconds):**

```
Chapter 1
Data Cleanup
Fixing the mess in seconds
```

---

## Chapter 1: Data Cleanup

**Duration:** 2:30
**These tools come from: modUTL_DataCleaning, modUTL_DataCleaningPlus, modUTL_DataSanitizer**

This chapter shows 4 cleanup tools, each one solving a specific, common Excel problem.

### Demo 1A: Delete Blank Rows

**Duration:** 0:30

### Script:

> "First problem — blank rows scattered through the data. Every filter, every formula, every sort breaks when you have random empty rows in the middle of your table.
>
> [run Delete Blank Rows]
>
> One click — it finds and removes every completely empty row. It creates a backup copy of the sheet first, just in case. [Read the confirmation: X blank rows deleted.]"

### Screen Actions:
1. Show the data with visible blank rows (scroll briefly to point them out)
2. Run DeleteBlankRows from the macro menu (Alt+F8 or developer tab)
3. Confirm the dialog
4. Show the result — clean data, no gaps

### Production Notes:
- The backup-before-destructive-action detail builds trust. Mention it.
- Quick and snappy. This is an appetizer.

---

### Demo 1B: Remove Leading/Trailing Spaces

**Duration:** 0:30

### Script:

> "Next — invisible spaces hiding in your text cells. You can't see them, but they break VLOOKUP matches, they mess up filters, and they make duplicates that aren't really duplicates.
>
> [select a column, run Remove Leading/Trailing Spaces]
>
> [Read confirmation: X cells cleaned.]
>
> Those phantom spaces are gone. Every VLOOKUP that was failing because of a trailing space — fixed."

### Screen Actions:
1. Select a column of text data
2. Run RemoveLeadingTrailingSpaces
3. Show the confirmation message

### Production Notes:
- "Breaks VLOOKUP matches" — this hits home for anyone who's fought with VLOOKUP. Name the pain specifically.

---

### Demo 1C: Convert Text to Numbers

**Duration:** 0:30

### Script:

> "One of the most common data import problems — numbers stored as text. The cell shows 1,250 but it's actually a text string. Your SUM formula returns zero. Your chart is blank.
>
> [select a range, run ConvertTextToNumbers]
>
> [Read confirmation: X text-stored numbers converted.]
>
> Now they're real numbers. SUM works. Charts work. Everything downstream works."

### Screen Actions:
1. Show a column where numbers have the green triangle (text-stored number indicator) or demonstrate that SUM returns 0
2. Select the range
3. Run ConvertTextToNumbers
4. Show the confirmation

### Production Notes:
- If you can show a SUM formula going from 0 to the correct total after conversion, that's a powerful visual. Consider setting this up in the sample file.

---

### Demo 1D: Unmerge Cells & Fill Down

**Duration:** 0:30

### Script:

> "Merged cells. Every data person's least favorite thing. They break sorting, filtering, copy-paste — essentially everything.
>
> [select a range with merged cells, run Unmerge And Fill Down]
>
> It unmerges every cell in the selection and fills the value down into the blanks that are left behind. Your data is flat and filterable now."

### Screen Actions:
1. Show the merged cells in column A (department names merged across rows)
2. Select the range
3. Run UnmergeAndFillDown
4. Show the result — flat, clean data with values filled down

### Production Notes:
- "Every data person's least favorite thing" — a moment of shared frustration. The audience will nod.

---

## CHAPTER CARD: Chapter 2

**On screen (3 seconds):**

```
Chapter 2
Formatting & Standardization
Making every file look professional
```

---

## Chapter 2: Formatting & Standardization

**Duration:** 2:00
**These tools come from: modUTL_Formatting, modUTL_Branding**

### Demo 2A: AutoFit All Columns & Rows

**Duration:** 0:20

### Script:

> "Quick one — AutoFit across the entire workbook. Every column, every row, every sheet — properly sized in one click.
>
> [run AutoFit All Columns & Rows — choose "Yes" for all sheets]
>
> Done. No more scrolling sideways to read truncated headers."

### Screen Actions:
1. Show some columns that are too narrow or too wide
2. Run AutoFitAllColumnsRows
3. Show the immediate visual improvement

### Production Notes:
- Fast. 20 seconds max. The visual before/after speaks for itself.

---

### Demo 2B: Apply iPipeline Branding

**Duration:** 0:40

### Script:

> "This one is my favorite for making files look professional fast.
>
> [run Apply iPipeline Branding]
>
> It automatically detects your header row, applies the official iPipeline Blue background with white text, sets alternating row colors for readability, and styles any total or summary rows in Navy Blue. All in official brand fonts and colors.
>
> [scroll through the formatted sheet]
>
> Five seconds ago this was a plain spreadsheet. Now it looks like it came from the corporate template library."

### Screen Actions:
1. Show the sheet with plain, unstyled data (default Excel look)
2. Run ApplyiPipelineBranding
3. Confirm the dialog
4. Pause on the result — the transformation should be visually dramatic
5. Slowly scroll through to show headers, alternating rows, and total rows

### Production Notes:
- THIS IS YOUR VISUAL WOW MOMENT in Video 3. The before/after contrast of plain Excel → branded professional table is instantly impressive.
- Pause for 2–3 seconds after the branding applies. Let the visual register.
- "Came from the corporate template library" — that's the reaction you want from viewers.

---

### Demo 2C: Date Format Standardizer

**Duration:** 0:30

### Script:

> "Mixed date formats — some cells show MM/DD/YYYY, others show DD-MMM-YY, and a few are just serial numbers from an import. It's a mess.
>
> [run DateFormatStandardizer]
>
> [Read confirmation: X date cells standardized to MM/DD/YYYY.]
>
> Every date in the workbook is now in the same format. No more guessing whether 03/05 is March 5th or May 3rd."

### Screen Actions:
1. Show a column with visibly mixed date formats
2. Run DateFormatStandardizer
3. Show the confirmation
4. Scroll through the column — all dates now consistent

---

### Demo 2D: Highlight Negatives in Red

**Duration:** 0:20

### Script:

> "Standard Finance formatting — every negative number should be red and bold. One click applies this across every sheet in the workbook.
>
> [run Highlight Negatives Red — choose "Yes" for all sheets]
>
> Instant visual scan — you can see where the losses and shortfalls are without reading a single number."

### Screen Actions:
1. Run HighlightNegativesRed
2. Show the result — negative numbers now jump out in red

---

## CHAPTER CARD: Chapter 3

**On screen (3 seconds):**

```
Chapter 3
Audit & Investigation
Finding problems before they find you
```

---

## Chapter 3: Audit & Investigation

**Duration:** 1:30
**These tools come from: modUTL_Audit, modUTL_AuditPlus, modUTL_WorkbookMgmt**

### Demo 3A: Workbook Health Check

**Duration:** 0:40

### Script:

> "Before you start working with any file you've received, run the Workbook Health Check.
>
> [run WorkbookHealthCheck]
>
> It scans the entire workbook and gives you a diagnostic report — how many sheets, how many formulas, how many errors, how many external links, how many blank cells. If anything needs attention, it flags it.
>
> [read a few key lines from the report]
>
> Think of it as a checkup for your spreadsheet. Ten seconds and you know exactly what you're working with."

### Screen Actions:
1. Run WorkbookHealthCheck
2. Read the message box report — pause so the viewer can see the stats
3. Brief pause at the end

### Production Notes:
- This is a credibility builder. It says "these tools are thorough and professional."

---

### Demo 3B: External Link Finder

**Duration:** 0:25

### Script:

> "If a file has formulas pointing to other workbooks — external links — you want to know about it before those links break.
>
> [run ExternalLinkFinder]
>
> It creates a report listing every cell that references an external file, with the exact sheet, cell address, and linked file path. If there are none, it tells you the workbook is self-contained."

### Screen Actions:
1. Run ExternalLinkFinder
2. If links are found, show the report sheet. If not, show the clean message.

---

### Demo 3C: Unhide All Sheets, Rows & Columns

**Duration:** 0:25

### Script:

> "Ever received a file and suspected there were hidden sheets or rows? One click reveals everything.
>
> [run UnhideAllSheetsRowsColumns]
>
> [Read confirmation: X hidden sheets revealed, all hidden rows and columns shown.]
>
> No more right-clicking and unhiding one sheet at a time."

### Screen Actions:
1. Show the tab bar — the sample file should have at least one hidden sheet
2. Run UnhideAllSheetsRowsColumns
3. Show the hidden sheet appearing in the tab bar

---

## CHAPTER CARD: Chapter 4

**On screen (3 seconds):**

```
Chapter 4
Python & SQL Power Tools
Command-line tools for bigger jobs
```

---

## Chapter 4: Python & SQL Power Tools

**Duration:** 1:30
**These tools come from the 22 Python scripts**

### Script (intro — 15 seconds):

> "Beyond VBA, the library includes 22 Python scripts for heavier-duty work — data consolidation, bank reconciliation, fuzzy matching, PDF extraction, even running SQL queries against your Excel files.
>
> These are command-line tools, but you don't need to be a programmer. Each one has a step-by-step guide, and the Copilot prompts can walk you through it. Let me show you what a couple of these look like."

### Demo 4A: Universal Data Cleaner (clean_data.py)

**Duration:** 0:30

### Script:

> "The Python Data Cleaner does in one command what would take five or six VBA macros — it removes empty rows and columns, trims spaces, converts text-stored numbers, standardizes dates, removes duplicates, and gives you a before-and-after summary.
>
> [show the terminal command and output]
>
> You point it at a file, it creates a cleaned copy, and it tells you exactly what it changed. No manual steps."

### Screen Actions:
1. Open a command prompt or terminal alongside Excel
2. Run: `python clean_data.py "Sample_Quarterly_Report.xlsx"`
3. Show the output summary (rows removed, cells cleaned, etc.)
4. Briefly open the cleaned output file to show the result

### Production Notes:
- The terminal might feel intimidating to non-technical viewers. Keep the narration reassuring: "one command, one file path, done."
- The step-by-step guide covers exactly how to open a terminal and type this command. Mention that.

---

### Demo 4B: SQL Query Tool (sql_query_tool.py)

**Duration:** 0:30

### Script:

> "This one is especially powerful — it lets you run SQL queries directly on your Excel or CSV files. No database needed. Your spreadsheet becomes a queryable table.
>
> [show a SQL query running against the sample file]
>
> Filter, aggregate, join two files together — anything SQL can do, you can do on your spreadsheets. If you know SQL, this will change how you work with Excel data."

### Screen Actions:
1. In the terminal, run a SQL query against the sample file
2. Example: `python sql_query_tool.py "Sample_Quarterly_Report.xlsx" --query "SELECT Department, SUM(Amount) FROM data GROUP BY Department ORDER BY SUM(Amount) DESC"`
3. Show the output

### Production Notes:
- This is a power-user feature. Not everyone will use it — but the people who do will love it.
- "If you know SQL, this will change how you work with Excel data" — that's a strong statement for the right audience. Let it land.

---

### Brief mention (15 seconds — narration only, no live demo):

> "The library also includes tools for bank reconciliation with fuzzy matching, PDF table extraction, budget vs. actual consolidation, forecast roll-forwards, variance decomposition, and more. Twenty-two scripts in total. Every one documented, every one available on SharePoint."

### Production Notes:
- This is a verbal inventory, not a demo. You're painting the picture of scale.
- Show the SharePoint folder or a list of script names on screen while narrating this.

---

## Chapter 5: Closing + Call to Action

**Duration:** 0:45
**On screen:** Sample file (now cleaned and formatted) or closing card

### Script:

> "That's a sample of what's in the universal tools library. We showed maybe ten percent of what's available.
>
> Here's how to get started:
>
> Everything is on SharePoint — the VBA code, the Python scripts, the SQL tools, and the sample files. Each tool has a step-by-step guide written for someone who's never touched code before. And if you need help, there's a master document with pre-built Copilot prompts that will walk you through using any tool, step by step.
>
> Start with the guide. Pick one tool that solves a problem you deal with every week. Try it. And if you have questions or ideas for new tools, reach out — I'm happy to help.
>
> Thanks for watching."

### Screen Actions:
- During "Everything is on SharePoint" — cut to closing card showing:

```
SharePoint: [folder path]

VBA Tools (75+) | Python Scripts (22) | SQL Tools
Step-by-Step Guides | Copilot Prompts

Questions? Contact Connor [last name / email]
```

### Production Notes:
- "Pick one tool that solves a problem you deal with every week" — this is a great CTA because it's specific and low-commitment. One tool. One problem. Try it.
- The closing card should stay on screen for 5+ seconds
- Brief music sting on the closing card

---

## Closing Title Card (5 seconds)

**On screen:**

```
iPipeline Finance Automation

Universal Tools — For Any Excel File
Available on SharePoint
[SharePoint location]

Questions? Contact Connor [last name / email]
```

**Audio:** Brief music, fade out

---

## Total Runtime Breakdown

| Section | Duration | Cumulative |
|---------|----------|------------|
| Title Card | 0:05 | 0:05 |
| Opening | 0:45 | 0:50 |
| Ch 1 Card | 0:03 | 0:53 |
| Ch 1: Data Cleanup (4 demos) | 2:30 | 3:23 |
| Ch 2 Card | 0:03 | 3:26 |
| Ch 2: Formatting (4 demos) | 2:00 | 5:26 |
| Ch 3 Card | 0:03 | 5:29 |
| Ch 3: Audit (3 demos) | 1:30 | 6:59 |
| Ch 4 Card | 0:03 | 7:02 |
| Ch 4: Python/SQL (2 demos + verbal list) | 1:30 | 8:32 |
| Closing + CTA | 0:45 | 9:17 |
| Closing Card | 0:05 | 9:22 |
| **TOTAL** | | **~9:22** |

Buffer for natural pauses and command execution time: expect **9:30–10:30** in practice.

---

## Tools Selected for Demo (and Why)

### Chapter 1 — Data Cleanup (problems everyone has):
| Tool | Source Module | Why This One |
|------|-------------|--------------|
| Delete Blank Rows | modUTL_DataCleaning | Universal pain point, visual before/after |
| Remove Spaces | modUTL_DataCleaning | Invisible problem with visible consequences (broken VLOOKUPs) |
| Text to Numbers | modUTL_DataCleaning | Extremely common import problem, satisfying fix |
| Unmerge & Fill Down | modUTL_DataCleaning | Everyone hates merged cells — instant relatability |

### Chapter 2 — Formatting (visual transformation):
| Tool | Source Module | Why This One |
|------|-------------|--------------|
| AutoFit All | modUTL_Formatting | Quick win, visible improvement |
| iPipeline Branding | modUTL_Branding | VISUAL WOW MOMENT — plain sheet → branded professional |
| Date Standardizer | modUTL_Formatting | Common problem, clean fix |
| Highlight Negatives | modUTL_Formatting | Finance standard, instant visual value |

### Chapter 3 — Audit (trust and investigation):
| Tool | Source Module | Why This One |
|------|-------------|--------------|
| Workbook Health Check | modUTL_WorkbookMgmt | Comprehensive diagnostic, builds confidence |
| External Link Finder | modUTL_Audit | Solves a real problem (broken links), generates a report |
| Unhide All | modUTL_WorkbookMgmt | Simple, dramatic, universally useful |

### Chapter 4 — Python/SQL (power tools):
| Tool | Source Script | Why This One |
|------|-------------|--------------|
| Universal Data Cleaner | clean_data.py | Shows Python's power — 6 operations in one command |
| SQL Query Tool | sql_query_tool.py | Game-changer for SQL users, impressive for non-SQL viewers |

### Not Demoed (but mentioned or available):
The following are available on SharePoint but not shown in the video:
- 14 Finance-specific VBA tools (Duplicate Invoice Detector, GL Validator, Trial Balance Checker, Ratio Dashboard, etc.)
- Fuzzy Match / Fuzzy Lookup
- Bank Reconciler
- PDF Table Extractor
- Budget vs. Actual Consolidator
- Forecast Roll-Forward
- Variance Decomposition
- Word Report Generator
- 40+ additional VBA tools across all modules

---

## Sample File Requirements

The sample file (Sample_Quarterly_Report.xlsx) needs the following "problems" baked in for the demos to work:

| Problem | Where | For Which Demo |
|---------|-------|---------------|
| 5-10 blank rows scattered in data | Rows 15, 28, 43, etc. | Delete Blank Rows |
| Leading/trailing spaces in text cells | Column B (names or descriptions) | Remove Spaces |
| Numbers stored as text (green triangles) | Column D or E (amounts) | Text to Numbers |
| Merged cells | Column A (department names) | Unmerge & Fill Down |
| Narrow/wide columns | Various | AutoFit All |
| No header formatting (plain default look) | Row 1 | iPipeline Branding |
| Mixed date formats | Column C | Date Standardizer |
| Negative numbers (plain, not red) | Column D or E | Highlight Negatives |
| At least 1 external link formula | Any cell | External Link Finder |
| At least 1 hidden sheet | A sheet named "Notes" or "Archive" | Unhide All |
| A few #N/A or #REF! errors | Scattered | Workbook Health Check |

**Suggested data structure for the sample file:**

| Column A | Column B | Column C | Column D | Column E | Column F |
|----------|----------|----------|----------|----------|----------|
| Department | Employee Name | Date | Amount | Budget | Variance |

~100-150 rows of data. Fictional names, realistic departments (Engineering, Sales, Marketing, Finance, Operations). Amounts in the $500-$50,000 range. Some negative variances.

---

## Recording Tips Specific to This Video

1. **Record each chapter as a separate clip.** Same as Videos 1 and 2.

2. **Practice the VBA macro runs.** Open the sample file, run each macro in order, confirm they all work correctly on the specific data in the sample file. Do this BEFORE recording.

3. **Practice the Python demos.** Open the terminal, run the commands, make sure the output looks clean and the scripts execute without errors. Have the exact commands ready to paste.

4. **The iPipeline Branding demo is your key visual moment.** If any single demo needs to be perfect, it's this one. Practice it twice.

5. **For the Python chapter, consider pre-recording the terminal output** and splicing it in during editing. Terminal text can be hard to read in a screen recording — you may want to zoom in or increase font size in the terminal.

6. **Keep the demo pattern consistent:** state the problem → run the tool → show the result. Every demo follows this exact rhythm. The viewer learns the pattern and starts anticipating the payoff.

---

*Script created: 2026-03-06 | Part of Video Demo Master Plan*
*Based on full review of: 23 VBA modules (~140+ tools) + 22 Python scripts*
