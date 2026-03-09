# Excel Automation Prompt Library

**VBA and Python Support Prompts for Coworkers**

World-class internal guide for using M365 Copilot to troubleshoot and tailor Excel automation code.

**Version:** 2.0
**Last Updated:** March 07, 2026

---

## Appendix A — Quick Reference

Use this section when you want the fastest help with the least typing. Click a prompt name to jump to it. Each description explains when to use the prompt.

### Recommended Prompt

| Prompt | When to Use |
|--------|-------------|
| [All-in-One Prompt](#recommended-all-in-one-prompt) | Best single prompt — fix the code for your workbook and get beginner run steps and tests |

### Universal Prompt

| Prompt | When to Use |
|--------|-------------|
| [Context Header](#universal-context-header) | Optional header that improves accuracy by providing context, expectations, and environment details |

### A — Quick Start Prompts

| Prompt | When to Use |
|--------|-------------|
| [A1 — Run a VBA Macro](#a1--run-a-vba-macro-correctly) | Run a VBA macro successfully and learn the prerequisites and required workbook structure |
| [A2 — Run a Python Script](#a2--run-a-python-script-correctly) | Run a Python script step by step and confirm required setup and workbook expectations |
| [A3 — Check if Code is Universal](#a3--confirm-whether-the-code-is-universal) | Check what the code assumes about sheet names, headers, paths, and what would break |

### B — Error and Bug Fix Prompts

| Prompt | When to Use |
|--------|-------------|
| [B1 — Fix a VBA Error](#b1--fix-a-vba-error) | Diagnose a VBA crash and apply the smallest fix with a clear explanation |
| [B2 — Fix a Python Error](#b2--fix-a-python-error) | Diagnose a Python crash and apply the smallest fix with a clear explanation |
| [B3 — Wrong Output](#b3--it-runs-but-the-output-is-wrong) | Trace logic to find why results are wrong and add an early warning check |
| [B4 — Works on One File But Not Another](#b4--it-works-on-one-file-but-not-another) | Compare workbook structure versus code requirements and choose code change or workbook change |

### C — Workbook Fit Prompts

| Prompt | When to Use |
|--------|-------------|
| [C1 — Map Workbook to Code](#c1--map-my-workbook-to-the-code-requirements) | Create a mapping of expected versus actual sheets and columns, then update the code to match |
| [C2 — Sheet Names Are Different](#c2--my-sheet-names-are-different) | Update all sheet references and make the code more resilient to name changes |
| [C3 — Columns Are Different](#c3--my-columns-are-different-or-in-a-different-order) | Update the code to locate columns by header name instead of fixed positions |
| [C4 — Multiple Tables or Sections](#c4--my-workbook-has-multiple-tables-or-sections) | Loop through repeated sections or matching sheets and produce a processing summary |

### D — Security and Environment Prompts

| Prompt | When to Use |
|--------|-------------|
| [D1 — Macros Are Blocked](#d1--macros-are-blocked) | Enable macros safely and confirm trust settings so the macro can run |
| [D2 — File Path Issues](#d2--file-path-or-network-drive-issues) | Remove hardcoded paths and use current folder or a folder picker for save/open actions |
| [D3 — Mac vs Windows Issues](#d3--it-fails-only-on-mac-or-only-on-windows) | Identify platform-specific parts and adjust for cross-platform compatibility |

### E — Quality Upgrade Prompts

| Prompt | When to Use |
|--------|-------------|
| [E1 — Add Error Messages](#e1--add-clear-error-messages-for-beginners) | Add friendly messages and upfront validation checks for sheets and required columns |
| [E2 — Add Logging](#e2--add-logging-so-we-can-debug-faster) | Add a log sheet for VBA and a log file plus console output for Python |
| [E3 — Make It Faster](#e3--make-it-faster-on-large-files) | Rewrite slow sections and explain performance improvements |
| [E4 — Make It Universal](#e4--make-it-more-universal-and-less-fragile) | Refactor to remove hardcoding and add a configuration section for easy edits |

### F — Feature Change Prompts

| Prompt | When to Use |
|--------|-------------|
| [F1 — Change Output Location](#f1--change-where-results-go) | Control output sheet and start cell using one setting |
| [F2 — Add a Filter or Condition](#f2--add-a-filter-or-condition) | Apply a new rule for which rows to include and show where to edit it later |
| [F3 — Add a Calculated Column](#f3--add-a-new-calculated-column) | Add a new computed field and validate placement and behavior |
| [F4 — Run Across Multiple Files](#f4--make-it-work-across-multiple-files) | Run the automation across a folder of files and generate a summary report |

### G — Learning Prompts

| Prompt | When to Use |
|--------|-------------|
| [G1 — Explain This Code](#g1--explain-this-code-like-i-am-new) | Beginner explanation of code purpose, inputs, outputs, and failure points |
| [G2 — Teach Me to Edit It](#g2--teach-me-how-to-safely-edit-it) | Identify safe sections to edit plus a simple testing checklist |
| [G3 — Create a Tutorial](#g3--create-a-mini-tutorial-using-my-workbook) | Create a short tutorial tailored to the uploaded workbook and code |

### H — Conversion Prompts

| Prompt | When to Use |
|--------|-------------|
| [H1 — VBA to Python](#h1--convert-vba-logic-to-python) | Translate VBA to Python with the same inputs and outputs and clear run steps |
| [H2 — Python to VBA](#h2--convert-python-logic-to-vba) | Translate Python to VBA and explain any limitations or alternatives |
| [H3 — Recommend Best Tool](#h3--recommend-vba-python-or-both) | Choose the best tool or combination and outline the workflow |

### Mega Prompt

| Prompt | When to Use |
|--------|-------------|
| [Mega Prompt](#one-copy-paste-mega-prompt) | End-to-end workflow — understand, validate assumptions, fix, harden, and give run instructions |

---

## Recommended All-in-One Prompt

Attach your workbook and your code file, then copy and paste the prompt below.

> I am uploading:
> 1. My Excel workbook: **[FILE NAME]**
> 2. My code file: **[FILE NAME]** (VBA or Python)
>
> **My goal:**
> In plain words, I want the code to do this:
> [DESCRIBE THE RESULT YOU WANT]
>
> **What I expected:**
> [WHAT YOU THOUGHT WOULD HAPPEN]
>
> **What actually happened:**
> [WHAT HAPPENED INSTEAD]
>
> **Error text (if any, copy-paste):**
> [PASTE ERROR OR WRITE "NO ERROR, JUST WRONG OUTPUT"]
>
> Please do ALL of the following, in this exact order:
>
> **Part 1 — Understand the code**
> 1. Explain what the code is trying to do in simple beginner language.
> 2. List everything the code assumes about my workbook.
>    Examples: sheet names, column headers, table names, named ranges, where data starts, where output goes, any required buttons, any file paths.
>
> **Part 2 — Compare the assumptions to my workbook**
> 3. Check my workbook and tell me which assumptions match and which ones do not match.
> 4. If something does not match, describe it clearly.
>    Example: "Code expects sheet 'Data' but my workbook has sheet 'Raw Data'."
>
> **Part 3 — Fix the code to work with my workbook**
> 5. Update the code so it works with my workbook as it is now.
> 6. Keep the change as small as possible first.
> 7. Then also give an improved version that is more universal and less likely to break in other workbooks.
> 8. Show me the final updated code I should use.
>
> **Part 4 — Give beginner-level step-by-step instructions**
> 9. Give me step-by-step instructions to run it, written for a beginner.
>    - Include every click and menu name.
>    - Include where I should start in Excel.
>    - Include exactly where to paste the code and how to save the file correctly.
>    - Include what to enable if something is blocked by security.
>    - Include what I should see when it works.
>
> **Part 5 — Confirm it worked and help me test**
> 10. Give me a simple test checklist.
> 11. Give me 3 common problems and exactly how to fix each one.
> 12. If you need any missing info from me, ask only the smallest number of questions needed.
>
> **Important constraints:**
> - Use plain language. Assume I am new.
> - Do not skip steps.
> - If there are multiple ways to run it, pick the easiest option and label it "Recommended."

### Optional Add-Ons

If you know these details, add them under the prompt. If you do not know them, write "not sure."

- My Excel version: [Windows or Mac] [Office version if known]
- How I am trying to run it: [Macro button, Macro list, Python terminal, notebook, or other]
- Where the input data is located: [sheet name]
- Where I want the output to go: [sheet name and starting cell]

---

# Chapter 1 — Purpose and How to Use This Document

## Purpose

This document gives you copy-paste prompts you can use with M365 Copilot to get help with Excel automation. It supports two types of automation:

- **VBA macros** that run inside Excel
- **Python scripts** that read or write Excel files

The prompts are designed for beginners. They force the right details so Copilot can fix your code and explain exactly how to run it.

## Who This Is For

Anyone in the company who is using shared VBA or Python code to automate Excel tasks.

## How to Use This Document

1. **Step 1** — Find the prompt that matches your situation (use the [Quick Reference](#appendix-a--quick-reference) at the top).
2. **Step 2** — Fill in the bracket fields like **[THIS]**.
3. **Step 3** — Attach your workbook and code file.
4. **Step 4** — Paste the prompt into Copilot and send.
5. **Step 5** — Follow the step-by-step instructions you receive.

## What to Upload

To get the best results, upload these items whenever you can:

- The Excel workbook you are working with
- The VBA module file (such as .bas) or the Python script file (such as .py)
- The exact error text, copied and pasted, if you have one
- A screenshot of the error popup if it is easier than copying text

## How to Upload Files to Copilot

1. **Step 1** — Open the Copilot chat.
2. **Step 2** — Use the paperclip or attach button.
3. **Step 3** — Select your Excel file and your code file.
4. **Step 4** — Wait until the upload completes.
5. **Step 5** — Paste your chosen prompt.
6. **Step 6** — Send the message.

## Privacy and Safe Sharing

Only upload files you are allowed to share. Remove customer data or sensitive data when required. If you are unsure, ask your manager or the data owner before uploading.

---

# Chapter 2 — Prompt Library

Pick the prompt that matches what you need. Each prompt is ready to copy and paste.

---

## Universal Context Header

Use this header before any prompt when you want the best results. Fill in what you can.

> **Context:**
> - Goal: [what you want the code to do]
> - What I expected to happen: [expected result]
> - What actually happened: [actual result]
> - Exact error text (copy-paste): [error or "none"]
> - Steps I took right before the issue: [what you did]
> - Excel version: [Windows or Mac, and Office version if known]
> - How I run the code: [Excel button, macro list, Python script, etc.]
> - Files I uploaded: [workbook name, code file name]
> - Anything special about the workbook: [password, protected sheets, large file, external links, etc.]

---

## A — Quick Start Prompts

### A1 — Run a VBA Macro Correctly

> I uploaded my Excel file and a VBA module. Please tell me exactly how to run the macro step by step in Excel.
>
> Also tell me what sheet I should be on when I run it and what prerequisites it assumes.
>
> If the macro expects certain sheet names, tables, or columns, list them clearly.

### A2 — Run a Python Script Correctly

> I uploaded my Excel file and a Python script. Please tell me how to run it step by step.
>
> Assume I am a beginner. If I need to install anything or enable anything, list it clearly.
>
> Also tell me what the script expects in the workbook (sheet names, column names, table layout).

### A3 — Confirm Whether the Code is Universal

> I uploaded code that I think is universal. Please review it and tell me what assumptions it makes about the workbook.
>
> List every dependency like sheet names, header names, table names, named ranges, file paths, and hardcoded values.
>
> Tell me what would break if a coworker has a different workbook layout.

---

## B — Error and Bug Fix Prompts

### B1 — Fix a VBA Error

> My VBA macro fails. I uploaded the workbook and the VBA code.
>
> Here is the exact error text:
> **[PASTE ERROR]**
>
> Please find the root cause and give me the smallest change needed to fix it.
> Explain why it fails in my workbook specifically.
> Then show the updated VBA code.

### B2 — Fix a Python Error

> My Python script fails. I uploaded the workbook and the Python code.
>
> Here is the exact error text:
> **[PASTE ERROR]**
>
> Please find the root cause and give me the smallest change needed to fix it.
> Explain why it fails in my workbook specifically.
> Then show the updated Python code.

### B3 — It Runs But the Output is Wrong

> The code runs without crashing but the results are wrong. I uploaded the workbook and code.
>
> - Expected result: [what you expected]
> - Actual result: [what you got]
>
> Please trace where the logic diverges, identify the faulty step, and propose a fix.
> Also suggest a simple check I can add to catch this earlier next time.

### B4 — It Works on One File But Not Another

> This code works on one workbook but fails on my workbook. I uploaded my workbook and the code.
>
> Please compare the code expectations versus my workbook structure. List the exact structural differences that matter and provide two options:
> 1. Change the code to support my workbook
> 2. Change my workbook layout to match the code

---

## C — Make the Code Fit Your Workbook

### C1 — Map My Workbook to the Code Requirements

> I uploaded my workbook and the code.
>
> Please create a mapping that shows what the code expects versus what my workbook has.
> For example: expected sheet name versus my sheet name, expected column header versus my header.
>
> Then update the code to use my names.

### C2 — My Sheet Names Are Different

> The code expects certain sheet names but my workbook uses different names.
>
> Please locate all sheet references in the code and update them to match my workbook.
> Also suggest a more robust approach so it does not break if names change again.

### C3 — My Columns Are Different or in a Different Order

> My workbook has different column headers or the columns are in a different order.
>
> Please update the code so it finds columns by header name rather than fixed column letters or numbers.
> If the code already does this, explain how it works and why it still might fail.

### C4 — My Workbook Has Multiple Tables or Sections

> My workbook has multiple similar sections and I need the code to run on all of them.
>
> Please modify the code to loop through each section or each sheet that matches a pattern.
> Also add a summary output that lists what it processed.

---

## D — Security, Permissions, and Environment

### D1 — Macros Are Blocked

> Excel will not let me run the macro. I see a security warning or the button is disabled.
>
> Please tell me the exact steps to enable it safely in Excel.
> Also tell me how to confirm the file is trusted and what settings matter.

### D2 — File Path or Network Drive Issues

> The code fails when it tries to open or save a file. I uploaded the workbook and code. The error mentions a path or permission.
>
> Please modify the code so it:
> 1. Uses the current workbook folder by default
> 2. Prompts the user to pick a folder if needed
> 3. Avoids hardcoded paths

### D3 — It Fails Only on Mac or Only on Windows

> This code behaves differently on Mac versus Windows. I uploaded the workbook and code.
>
> Please identify what parts are platform-specific and propose a version that works on both.
> If that is not possible, tell me what needs to be different for each platform.

---

## E — Quality Upgrades

### E1 — Add Clear Error Messages for Beginners

> Please improve this code so when it fails it shows a clear friendly message that tells the user what to fix. Do not show technical stack traces unless I ask.
>
> Also add checks at the start to validate sheet names and required columns.

### E2 — Add Logging So We Can Debug Faster

> Please add logging to this code.
>
> - **For VBA:** Log to a new sheet called "Log" with timestamp, step name, and status.
> - **For Python:** Log to console and also write a log file next to the workbook.
>
> Show me the updated code and where to look when something goes wrong.

### E3 — Make It Faster on Large Files

> This code is slow on my workbook. I uploaded the workbook and code.
>
> Please identify the slow parts and rewrite them for performance.
> Also explain what changed and why it is faster.

### E4 — Make It More Universal and Less Fragile

> Please refactor this code to be more universal:
> - Remove hardcoded sheet names where possible
> - Prefer finding tables by header, named ranges, or structured tables
> - Add a single configuration section at the top where a coworker can change key settings
>
> Show the final code and the configuration section clearly.

---

## F — Feature Change Requests

### F1 — Change Where Results Go

> I need the output to go to a different sheet and start cell.
>
> Please update the code so I can control the destination using a single setting.
> Also keep existing behavior as the default if the setting is blank.

### F2 — Add a Filter or Condition

> I need to change the logic so it only includes rows where:
> **[DESCRIBE CONDITION]**
>
> Please update the code and show exactly where the condition lives.
> Also suggest how I can change it later without breaking anything.

### F3 — Add a New Calculated Column

> I need to add a new calculated field:
> - **Name:** [column name]
> - **Formula in plain English:** [what it calculates]
> - **Where it should appear:** [position]
>
> Please implement it in the code and verify it works with my workbook structure.

### F4 — Make It Work Across Multiple Files

> I need this to run across many Excel files in a folder.
>
> Please update the code to:
> 1. Ask me to pick a folder
> 2. Loop through each workbook
> 3. Run the same logic
> 4. Produce one summary report at the end

---

## G — Learning Prompts

### G1 — Explain This Code Like I Am New

> Please explain what this code does in simple language. Go section by section.
>
> For each section, tell me:
> 1. Purpose
> 2. Inputs
> 3. Outputs
> 4. What could go wrong

### G2 — Teach Me How to Safely Edit It

> I want to customize this code but I am worried about breaking it.
>
> Please show me which parts are safe to edit and which parts I should not touch.
> Also suggest a simple checklist to test changes.

### G3 — Create a Mini Tutorial Using My Workbook

> Use my uploaded workbook and code to create a short tutorial. Include:
> 1. What the workbook must contain
> 2. How to run the code
> 3. What the output looks like
> 4. Three common errors and how to fix them
>
> Keep it beginner-friendly.

---

## H — Conversion and Hybrid Prompts

### H1 — Convert VBA Logic to Python

> Please convert this VBA logic to Python that works with Excel.
>
> Keep the same inputs and outputs. Explain what Python libraries you use and why.
> Also tell me how to run it.

### H2 — Convert Python Logic to VBA

> Please convert this Python logic to VBA that runs inside Excel.
>
> Keep the same inputs and outputs. If something cannot be done cleanly in VBA, tell me the closest alternative.

### H3 — Recommend VBA, Python, or Both

> Given this task and my workbook, should I use VBA, Python, or both?
>
> Please recommend the best approach and explain why.
> Then show what the final workflow would look like.

---

## One Copy-Paste Mega Prompt

> I uploaded my workbook and the code. Please do all of the following in order:
>
> 1. Identify what the code is trying to do.
> 2. List every workbook assumption it makes (sheet names, headers, tables, named ranges, paths).
> 3. Validate those assumptions against my workbook and list mismatches.
> 4. If there is an error, identify the root cause.
> 5. Provide the smallest fix first.
> 6. Provide an improved universal version second (more robust, fewer hardcoded dependencies).
> 7. Give me step-by-step instructions to run it.
> 8. Give me a quick test checklist to confirm it worked.
>
> **Constraints:**
> - Use simple language.
> - If anything is unclear, ask me specific questions.

---

# Chapter 3 — Tips for Getting the Best Results

## When to Use VBA Versus Python

- **Use VBA** when the automation needs to interact with Excel buttons, forms, and workbook events.
- **Use Python** when you need stronger data manipulation, file processing, or automation across many files.
- **Not sure?** Use the [H3 — Recommend VBA, Python, or Both](#h3--recommend-vba-python-or-both) prompt.

## What Makes Code Universal

Code is only universal when it does **not** assume fixed sheet names, fixed column positions, or hardcoded paths. Universal code usually finds data by looking for headers, table names, or named ranges.

## How to Describe Your Goal

The best description is the **business result**, not the technical steps.

**Good example:** "Create a summary sheet that shows total spend by vendor for the current month."

**Less helpful:** "Loop through column B and sum matching values into a new sheet."

## If You Cannot Share Your Workbook

If the workbook contains sensitive data:
1. Create a copy of the file.
2. Remove sensitive rows or replace sensitive values with dummy data.
3. Keep the same sheet names and column headers so the code behavior stays similar.

## Support

If you are still blocked after using these prompts, share the Copilot conversation and the latest version of your workbook and code with the project owner.

---

*Finance Automation — Copilot Prompt Library v2.0*
