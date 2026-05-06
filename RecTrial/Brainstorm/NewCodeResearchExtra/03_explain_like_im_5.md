# Report 3 — Explain This Like I’m 5
## iPipeline Finance Automation Video 4 + Toolkit Plan

This file explains the plan in very simple terms.

---

## 1. What you built so far

You built a big set of tools that help Finance people work faster.

Some tools are in Excel.

Some tools are in Python.

The Excel tools help with things like:
- cleaning messy spreadsheets;
- comparing files;
- finding formula problems;
- making summaries;
- checking for duplicate invoices;
- organizing tabs;
- making reports.

The Python tools help with things like:
- checking files;
- comparing lots of data;
- finding exceptions;
- making reports;
- creating evidence folders.

You also made videos showing people what the tools can do.

Videos 1, 2, and 3 are already done.

Video 4 is still being decided.

---

## 2. The big change

At first, the project sounded like it was for 2,000 people.

Now the real near-term audience is more like **50 to 150 coworkers**.

That changes the plan.

You do not need to make this perfect for the whole company yet.

You need to make it safe, clear, and useful for a smaller group of coworkers.

That is easier.

But it still needs to look professional.

---

## 3. What Video 4 should do

Video 4 should not try to show every Python script.

That would be too much.

Video 4 should show:

> “Python can help Finance find problems, check files, and create useful reports safely.”

That is the main idea.

---

## 4. The best hero demo

The best main demo is:

# Revenue Leakage Finder

That means:

> “Find places where the company may have billed the wrong amount.”

Example:

A customer should have been billed $10,000.

But the billing file says they were billed $8,500.

The tool finds the $1,500 possible problem.

Then it shows:
- who the customer is;
- what the expected amount was;
- what the actual billed amount was;
- the difference;
- how serious it is;
- what someone should check next.

That is a strong demo because people understand money problems.

---

## 5. Why not make ARR Waterfall the main demo?

ARR Waterfall is still useful.

But it mostly says:

> “Python made a nice finance chart.”

That is okay, but not as strong.

Revenue Leakage Finder says:

> “Python found a possible money problem.”

That is stronger.

Use ARR Waterfall only as a smaller side output if you still want it.

---

## 6. Should Video 4 be split into 4a and 4b?

Probably not.

Make one official Video 4.

Call it:

# Video 4 of 4 — Python Automation for Finance

Make it about 9 to 12 minutes.

Use chapters inside the video.

Later, you can make short extra recipe videos if you want.

But do not make the official series confusing with “4a” and “4b.”

---

## 7. Simple Video 4 structure

Here is the simple version:

### Part 1 — Why Python matters
Excel is great.

VBA is great for Excel.

Python is better when you need to:
- work with many files;
- check data rules;
- create repeatable reports;
- build output folders;
- make logs;
- handle bigger workflows.

### Part 2 — Safety
Say clearly:

- The scripts run on your computer.
- They do not send data to the internet.
- They do not use external AI.
- They do not change the original input files.
- They write outputs to a separate folder.
- They create logs so you know what happened.

### Part 3 — Revenue Leakage Finder
Show the main demo.

Input files go in.

The tool checks them.

A report comes out.

### Part 4 — Data Contract Checker
Show a bad file.

The tool says FAIL.

Fix the file.

Run it again.

The tool says PASS.

This teaches people that automation can catch bad inputs before they cause bad reports.

### Part 5 — Exception Triage
Show that Python can rank problems.

It tells people what to check first.

### Part 6 — Control Evidence Pack
Show that Python can create an evidence folder.

This helps with audit/control/review work.

### Part 7 — Launcher menu
Show a simple menu.

The menu lets users run safe sample tools.

### Part 8 — How to start
Tell people:
- start with the sample files;
- use the recommended workflows first;
- ask Connor if something breaks;
- do not run it on sensitive production files until you understand it.

---

## 8. What “safe Python” means

Safe Python does not just mean “the code seems fine.”

Safe means a normal coworker can run it without causing damage.

Safe means:

1. It does not use the internet.
2. It does not call external AI.
3. It does not ask for passwords.
4. It does not change original files.
5. It does not delete files.
6. It does not overwrite source files.
7. It saves outputs in a separate folder.
8. It creates a log.
9. It has sample mode.
10. It gives simple error messages.

That is what you should build.

---

## 9. What the coworker should see first

Do not show them 140 tools first.

That is too many.

Show them 5 to 7 starting workflows.

Good starting workflows:

1. Clean a messy Excel export.
2. Compare two files.
3. Consolidate sheets or files.
4. Find workbook errors and external links.
5. Generate a workbook or executive summary.
6. Run Revenue Leakage Finder on sample data.
7. Run Data Contract Checker on sample data.

After that, they can explore more tools.

---

## 10. How people should get the tools

Use SharePoint.

Do not email attachments around.

Put one official folder or page on SharePoint.

It should have:

```text
Finance Automation Toolkit v1.0
├── 00_START_HERE.pdf
├── Finance_Automation_Toolkit.xlsm
├── Python_Finance_Starter_Pack.zip
├── Sample_Files.zip
├── Quick_Reference_Card.pdf
├── Known_Limitations.pdf
├── Troubleshooting.pdf
└── Release_Notes.pdf
```

People should download from the official place.

That way they know they have the right version.

---

## 11. What you should not build yet

Do not build these right now:

- xlwings Excel buttons;
- external AI tools;
- email automation;
- scheduled automation;
- dashboards on web apps;
- machine learning forecasting;
- anything that needs IT/admin installation.

Those can wait.

Right now, the goal is:

> safe, useful, easy-to-demo tools.

---

## 12. What you should do next

Before writing more code:

1. Pick Revenue Leakage Finder as the main Video 4 demo.
2. Make one official Video 4, not 4a and 4b.
3. Rename the Python menu to something simple like **Finance Automation Launcher**.
4. Write a short Python safety file.
5. Make the release package focus on 5 to 7 workflows, not 140 tools.

That is the plan in simple terms.
