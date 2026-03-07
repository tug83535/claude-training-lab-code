**Excel Automation Prompt Library**
**VBA and Python Support Prompts for Coworkers**
World class internal guide for using M365 Copilot to troubleshoot and tailor Excel automation code

Version 1.0
Last updated: March 04, 2026


| Appendix A  Quick Reference |
| --- |


Use this section when you want the fastest help with the least typing.
## Quick links
Click a prompt name to jump to it. Each description explains when to use the prompt.
### Recommended prompts
- : Best single prompt: fix the code for your workbook and give beginner run steps and tests.
### Universal prompts
- : Optional header that improves accuracy by providing context, expectations, and environment details.
### A prompts
- : Run a VBA macro successfully and learn the prerequisites and required workbook structure.
- : Run a Python script step by step and confirm required setup and workbook expectations.
- : Check what the code assumes about sheet names, headers, paths, and what would break.
### B prompts
- : Diagnose a VBA crash and apply the smallest fix with a clear explanation.
- : Diagnose a Python crash and apply the smallest fix with a clear explanation.
- : Trace logic to find why results are wrong and add an early warning check.
- : Compare workbook structure versus code requirements and choose code change or workbook change.
### C prompts
- : Create a mapping of expected versus actual sheets and columns, then update the code to match.
- : Update all sheet references and make the code more resilient to name changes.
- : Update the code to locate columns by header name instead of fixed positions.
- : Loop through repeated sections or matching sheets and produce a processing summary.
### D prompts
- : Enable macros safely and confirm trust settings so the macro can run.
- : Remove hardcoded paths and use current folder or a folder picker for save open actions.
- : Identify platform specific parts and adjust for cross platform compatibility.
### E prompts
- : Add friendly messages and upfront validation checks for sheets and required columns.
- : Add a log sheet for VBA and a log file plus console output for Python.
- : Rewrite slow sections and explain performance improvements.
- : Refactor to remove hardcoding and add a configuration section for easy edits.
### F prompts
- : Control output sheet and start cell using one setting.
- : Apply a new rule for which rows to include and show where to edit it later.
- : Add a new computed field and validate placement and behavior.
- : Run the automation across a folder of files and generate a summary report.
### G prompts
- : Beginner explanation of code purpose, inputs, outputs, and failure points.
- : Identify safe sections to edit plus a simple testing checklist.
- : Create a short tutorial tailored to the uploaded workbook and code.
### H prompts
- : Translate VBA to Python with the same inputs and outputs and clear run steps.
- : Translate Python to VBA and explain any limitations or alternatives.
- : Choose the best tool or combination and outline the workflow.
### Mega prompts
- : End to end workflow: understand, validate assumptions, fix, harden, and give run instructions.
Attach your workbook and your code file, then copy and paste the prompt below.
## Recommended all in one prompt
I am uploading:
1. My Excel workbook: [FILE NAME]
2. My code file: [FILE NAME] (VBA or Python)

My goal:
In plain words, I want the code to do this:
[DESCRIBE THE RESULT YOU WANT]

What I expected:
[WHAT YOU THOUGHT WOULD HAPPEN]

What actually happened:
[WHAT HAPPENED INSTEAD]

Error text (if any, copy paste):
[PASTE ERROR OR WRITE "NO ERROR, JUST WRONG OUTPUT"]

Please do ALL of the following, in this exact order:

Part 1  Understand the code
1. Explain what the code is trying to do in simple beginner language.
2. List everything the code assumes about my workbook.
   Examples: sheet names, column headers, table names, named ranges, where data starts, where output goes, any required buttons, any file paths.

Part 2  Compare the assumptions to my workbook
3. Check my workbook and tell me which assumptions match and which ones do not match.
4. If something does not match, describe it clearly.
   Example: code expects sheet "Data" but my workbook has sheet "Raw Data".

Part 3  Fix the code to work with my workbook
5. Update the code so it works with my workbook as it is now.
6. Keep the change as small as possible first.
7. Then also give an improved version that is more universal and less likely to break in other workbooks.
8. Show me the final updated code I should use.

Part 4  Give beginner level step by step instructions
9. Give me step by step instructions to run it, written for a beginner.
   Include every click and menu name.
   Include where I should start in Excel.
   Include exactly where to paste the code and how to save the file correctly.
   Include what to enable if something is blocked by security.
   Include what I should see when it works.

Part 5  Confirm it worked and help me test
10. Give me a simple test checklist.
11. Give me 3 common problems and exactly how to fix each one.
12. If you need any missing info from me, ask only the smallest number of questions needed.

Important constraints:
Use plain language. Assume I am new.
Do not skip steps.
If there are multiple ways to run it, pick the easiest option and label it "Recommended".
## Optional add ons
If you know these details, add them under the prompt. If you do not know them, write not sure.
My Excel version: [Windows or Mac] [Office version if known]
How I am trying to run it: [Macro button, Macro list, Python terminal, notebook, or other]
Where the input data is located: [sheet name]
Where I want the output to go: [sheet name and starting cell]


| Chapter 1  Purpose and How to Use This Document |
| --- |


## Purpose
This document gives you copy paste prompts you can use with M365 Copilot to get help with Excel automation.
It supports two types of automation:
- VBA macros that run inside Excel
- Python scripts that read or write Excel files
The prompts are designed for beginners. They force the right details so Copilot can fix your code and explain exactly how to run it.
## Who this is for
Anyone in the company who is using shared VBA or Python code to automate Excel tasks.
## How to use this document
Step 1  Find the prompt that matches your situation.
Step 2  Fill in the bracket fields like [THIS].
Step 3  Attach your workbook and code file.
Step 4  Paste the prompt into Copilot and send.
Step 5  Follow the step by step instructions you receive.
## What to upload
To get the best results, upload these items whenever you can:
- The Excel workbook you are working with
- The VBA module file such as BAS, or the Python script file such as PY
- The exact error text, copied and pasted, if you have one
- A screenshot of the error popup if it is easier than copying text
## How to upload files to Copilot
Step 1  Open the Copilot chat.
Step 2  Use the paperclip or attach button.
Step 3  Select your Excel file and your code file.
Step 4  Wait until the upload completes.
Step 5  Paste your chosen prompt.
Step 6  Send the message.
## Privacy and safe sharing
Only upload files you are allowed to share. Remove customer data or sensitive data when required.
If you are unsure, ask your manager or the data owner before uploading.


| Chapter 2  Prompt Library |
| --- |


Pick the prompt that matches what you need. Each prompt is ready to copy and paste.
## Universal context header
Use this header before any prompt when you want the best results. Fill in what you can.
Context
Goal:
What I expected to happen:
What actually happened:
Exact error text (copy paste):
Steps I took right before the issue:
Excel version (Windows or Mac, and Office version if known):
Where the code is run from (Excel button, macro list, Python script, and so on):
Files I uploaded (workbook name, code file name):
Anything special about the workbook (password, protected sheets, large file, external links):
## A  Quick start prompts
### A1  Run a VBA macro correctly
I uploaded my Excel file and a VBA module. Please tell me exactly how to run the macro step by step in Excel.
Also tell me what sheet I should be on when I run it and what prerequisites it assumes.
If the macro expects certain sheet names, tables, or columns, list them clearly.
### A2  Run a Python script correctly
I uploaded my Excel file and a Python script. Please tell me how to run it step by step.
Assume I am a beginner.
If I need to install anything or enable anything, list it clearly.
Also tell me what the script expects in the workbook (sheet names, column names, table layout).
### A3  Confirm whether the code is universal
I uploaded code that I think is universal. Please review it and tell me what assumptions it makes about the workbook.
List every dependency like sheet names, header names, table names, named ranges, file paths, and hardcoded values.
Tell me what would break if a coworker has a different workbook layout.
## B  Error and bug fix prompts
### B1  Fix a VBA error
My VBA macro fails. I uploaded the workbook and the VBA code.
Here is the exact error text:
[PASTE ERROR]
Please find the root cause and give me the smallest change needed to fix it.
Explain why it fails in my workbook specifically.
Then show the updated VBA code.
### B2  Fix a Python error
My Python script fails. I uploaded the workbook and the Python code.
Here is the exact error text:
[PASTE ERROR]
Please find the root cause and give me the smallest change needed to fix it.
Explain why it fails in my workbook specifically.
Then show the updated Python code.
### B3  It runs but the output is wrong
The code runs without crashing but the results are wrong.
I uploaded the workbook and code.
Expected result:
Actual result:
Please trace where the logic diverges, identify the faulty step, and propose a fix.
Also suggest a simple check I can add to catch this earlier next time.
### B4  It works on one file but not another
This code works on one workbook but fails on my workbook.
I uploaded my workbook and the code.
Please compare the code expectations versus my workbook structure.
List the exact structural differences that matter and provide two options:
1. Change the code to support my workbook
2. Change my workbook layout to match the code
## C  Make the code fit your workbook
### C1  Map my workbook to the code requirements
I uploaded my workbook and the code.
Please create a mapping that shows what the code expects versus what my workbook has.
For example: expected sheet name versus my sheet name, expected column header versus my header.
Then update the code to use my names.
### C2  My sheet names are different
The code expects certain sheet names but my workbook uses different names.
Please locate all sheet references in the code and update them to match my workbook.
Also suggest a more robust approach so it does not break if names change again.
### C3  My columns are different or in a different order
My workbook has different column headers or the columns are in a different order.
Please update the code so it finds columns by header name rather than fixed column letters or numbers.
If the code already does this, explain how it works and why it still might fail.
### C4  My workbook has multiple tables or sections
My workbook has multiple similar sections and I need the code to run on all of them.
Please modify the code to loop through each section or each sheet that matches a pattern.
Also add a summary output that lists what it processed.
## D  Security, permissions, and environment
### D1  Macros are blocked
Excel will not let me run the macro. I see a security warning or the button is disabled.
Please tell me the exact steps to enable it safely in Excel.
Also tell me how to confirm the file is trusted and what settings matter.
### D2  File path or network drive issues
The code fails when it tries to open or save a file.
I uploaded the workbook and code. The error mentions a path or permission.
Please modify the code so it:
1. Uses the current workbook folder by default
2. Prompts the user to pick a folder if needed
3. Avoids hardcoded paths
### D3  It fails only on Mac or only on Windows
This code behaves differently on Mac versus Windows.
I uploaded the workbook and code.
Please identify what parts are platform specific and propose a version that works on both.
If that is not possible, tell me what needs to be different for each platform.
## E  Quality upgrades
### E1  Add clear error messages for beginners
Please improve this code so when it fails it shows a clear friendly message that tells the user what to fix.
Do not show technical stack traces unless I ask.
Also add checks at the start to validate sheet names and required columns.
### E2  Add logging so we can debug faster
Please add logging to this code.
For VBA: log to a new sheet called Log with timestamp, step name, and status.
For Python: log to console and also write a log file next to the workbook.
Show me the updated code and where to look when something goes wrong.
### E3  Make it faster on large files
This code is slow on my workbook.
I uploaded the workbook and code.
Please identify the slow parts and rewrite them for performance.
Also explain what changed and why it is faster.
### E4  Make it more universal and less fragile
Please refactor this code to be more universal.
Remove hardcoded sheet names where possible.
Prefer finding tables by header, named ranges, or structured tables.
Add a single configuration section at the top where a coworker can change key settings.
Show the final code and the configuration section clearly.
## F  Feature change requests
### F1  Change where results go
I need the output to go to a different sheet and start cell.
Please update the code so I can control the destination using a single setting.
Also keep existing behavior as the default if the setting is blank.
### F2  Add a filter or condition
I need to change the logic so it only includes rows where:
[DESCRIBE CONDITION]
Please update the code and show exactly where the condition lives.
Also suggest how I can change it later without breaking anything.
### F3  Add a new calculated column
I need to add a new calculated field:
Name:
Formula in plain English:
Where it should appear:
Please implement it in the code and verify it works with my workbook structure.
### F4  Make it work across multiple files
I need this to run across many Excel files in a folder.
Please update the code to:
1. Ask me to pick a folder
2. Loop through each workbook
3. Run the same logic
4. Produce one summary report at the end
## G  Learning prompts
### G1  Explain this code like I am new
Please explain what this code does in simple language.
Go section by section.
For each section, tell me:
1. Purpose
2. Inputs
3. Outputs
4. What could go wrong
### G2  Teach me how to safely edit it
I want to customize this code but I am worried about breaking it.
Please show me which parts are safe to edit and which parts I should not touch.
Also suggest a simple checklist to test changes.
### G3  Create a mini tutorial using my workbook
Use my uploaded workbook and code to create a short tutorial.
Include:
1. What the workbook must contain
2. How to run it
3. What the output looks like
4. Three common errors and how to fix them
Keep it beginner friendly.
## H  Conversion and hybrid prompts
### H1  Convert VBA logic to Python
Please convert this VBA logic to Python that works with Excel.
Keep the same inputs and outputs.
Explain what Python libraries you use and why.
Also tell me how to run it.
### H2  Convert Python logic to VBA
Please convert this Python logic to VBA that runs inside Excel.
Keep the same inputs and outputs.
If something cannot be done cleanly in VBA, tell me the closest alternative.
### H3  Recommend VBA, Python, or both
Given this task and my workbook, should I use VBA, Python, or both
Please recommend the best approach and explain why.
Then show what the final workflow would look like.
## One copy paste mega prompt
I uploaded my workbook and the code.
Please do all of the following in order:

1. Identify what the code is trying to do.
2. List every workbook assumption it makes (sheet names, headers, tables, named ranges, paths).
3. Validate those assumptions against my workbook and list mismatches.
4. If there is an error, identify the root cause.
5. Provide the smallest fix first.
6. Provide an improved universal version second (more robust, fewer hardcoded dependencies).
7. Give me step by step instructions to run it.
8. Give me a quick test checklist to confirm it worked.

Constraints:
Use simple language.
If anything is unclear, ask me specific questions.


| Chapter 3  Tips for Getting the Best Results |
| --- |


## When to use VBA versus Python
Use VBA when the automation needs to interact with Excel buttons, forms, and workbook events.
Use Python when you need stronger data manipulation, file processing, or automation across many files.
If you are unsure, use the recommendation prompt in section H.
## What makes code universal
Code is only universal when it does not assume fixed sheet names, fixed column positions, or hardcoded paths.
Universal code usually finds data by looking for headers, table names, or named ranges.
## How to describe your goal
The best description is the business result, not the technical steps.
Example: Create a summary sheet that shows total spend by vendor for the current month.
## If you cannot share your workbook
If the workbook contains sensitive data, create a copy and remove sensitive rows or replace sensitive values.
Keep the same sheet names and column headers so the code behavior stays similar.
## Support
If you are still blocked after using these prompts, share the Copilot conversation and the latest version of your workbook and code with the project owner.


