# Step-by-Step Guide — Build FinanceTools.xlsm
## Creating the Excel Workbook with the Finance Tools Button

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Date:** 2026-04-29
**Time required:** ~15 minutes
**Skill level:** No VBA experience needed — every step is written out completely

---

## What you are building

A single Excel file called `FinanceTools.xlsm`. It contains one button labeled
**"Finance Tools."** When a coworker clicks it, a command-line window opens with
the numbered Finance Tools menu. That's it — the whole workbook is just the button.

The Python scripts do all the actual work. The workbook is just the launcher.

---

## Before you start — what you need

- [ ] Excel installed and open (any version from 2016 onward works)
- [ ] The file `modFinanceToolsLauncher.bas` on your machine
      Path: `C:\Users\connor.atlee\RecTrial\UniversalToolkit\python\ZeroInstall\modFinanceToolsLauncher.bas`
- [ ] About 15 minutes of uninterrupted time

You do **not** need Python installed to complete this guide.
You do **not** need to understand VBA code.

---

## PART 1 — Create and save the workbook

### Step 1 — Open Excel
Open Microsoft Excel. It will show the Start screen or a blank workbook.

### Step 2 — Create a new blank workbook
If you see the Start screen: click **Blank workbook**.
If Excel opened directly to a workbook: you're already good.

You should see a plain spreadsheet — Sheet1, column A, row 1 in the top-left cell.

### Step 3 — Save the file as a macro-enabled workbook
This is the most important step. Excel has two workbook formats:
- `.xlsx` — standard Excel file. **Cannot run macros. Do not use this.**
- `.xlsm` — macro-enabled Excel file. **This is what you need.**

Do this now:
1. Press **Ctrl + S** (or click File → Save As)
2. Choose where to save it. For now, save it to your Desktop or any easy-to-find location.
   *(You will move it to the final SharePoint zip folder later.)*
3. In the **"Save as type"** dropdown, click it and select:
   **Excel Macro-Enabled Workbook (*.xlsm)**
   *(It will be near the top of the list — look for the one with "Macro-Enabled" in the name)*
4. In the **File name** box, type exactly: `FinanceTools`
5. Click **Save**

The title bar at the top of Excel should now show **FinanceTools.xlsm**.

**If a yellow warning bar appears** saying "Macros have been disabled" — that is normal.
You will handle that in a later step.

---

## PART 2 — Enable the Developer tab

The Developer tab is where you access the VBA editor and button tools.
It is hidden by default. You only need to do this once.

### Step 4 — Check if the Developer tab is already visible
Look at the ribbon tabs across the top of Excel:
**Home | Insert | Page Layout | Formulas | Data | Review | View**

If you see **Developer** at the end of that list — skip to Step 7.
If you do NOT see it — continue with Steps 5 and 6.

### Step 5 — Open Excel Options
Click **File** (top-left corner of Excel) → then click **Options** at the very bottom
of the left sidebar.

The "Excel Options" window opens.

### Step 6 — Turn on the Developer tab
1. In the left panel of the Options window, click **Customize Ribbon**
2. On the right side, you will see a list of tabs with checkboxes
3. Find **Developer** in the list (it should be near the bottom of the right panel)
4. Check the box next to **Developer**
5. Click **OK**

The Developer tab now appears in the ribbon.

---

## PART 3 — Import the VBA code

This is where you bring in the button logic. You are importing a pre-written file —
you do not need to type any code.

### Step 7 — Open the VBA editor
Press **Alt + F11** on your keyboard at the same time.

A new window opens — this is the Visual Basic for Applications (VBA) editor.
It will look unfamiliar. That is normal. You only need two things in this window:
- The menu bar at the top (File, Edit, View, Insert...)
- The large grey/white area in the middle

### Step 8 — Import the .bas file
1. In the VBA editor, click **File** in the menu bar (the VBA editor's own menu, not Excel's)
2. Click **Import File...**
3. A file browser opens. Navigate to:
   `C:\Users\connor.atlee\RecTrial\UniversalToolkit\python\ZeroInstall\`
4. Find the file `modFinanceToolsLauncher.bas` and click it once to select it
5. Click **Open**

### Step 9 — Confirm the import worked
On the left side of the VBA editor, there is a panel called the **Project Explorer**.
*(If you don't see it, press Ctrl + R to show it.)*

You should see something like:
```
VBAProject (FinanceTools.xlsm)
  └── Modules
        └── modFinanceToolsLauncher
```

If `modFinanceToolsLauncher` appears under Modules — the import worked. ✓

If you do not see the Project Explorer panel on the left:
- Click **View** in the VBA editor menu bar → click **Project Explorer**

### Step 10 — Close the VBA editor
Press **Alt + F4** OR click the X in the top-right corner of the VBA editor window.

You are back in the normal Excel window. The VBA code is now saved inside the workbook.

---

## PART 4 — Add the Finance Tools button

Now you will draw a button on the sheet and connect it to the VBA code you just imported.

### Step 11 — Go to the Developer tab
Click the **Developer** tab in the Excel ribbon (the one you turned on in Part 2).

### Step 12 — Open the Insert controls menu
In the Developer ribbon, look for the **Controls** group.
Inside that group, click the **Insert** button.

A small dropdown appears showing two sections:
- **Form Controls** (top section — simple shapes)
- **ActiveX Controls** (bottom section — more complex)

**Use Form Controls only.** Do NOT use ActiveX Controls.

### Step 13 — Select the Button tool
In the **Form Controls** section (top half of the dropdown), click the icon that looks
like a small rectangle with a cursor — it should say **Button (Form Control)** when
you hover over it.

Your cursor will change to a crosshair (+).

### Step 14 — Draw the button on the sheet
Click and drag on the spreadsheet to draw the button.

Suggested placement and size:
- Start around cell **B3** and drag to around cell **D6**
- This gives you a medium-sized button that is easy to see and click

The exact size and position don't matter for now — you can resize and move it later.

**As soon as you release the mouse**, a dialog box called **"Assign Macro"** opens automatically.

### Step 15 — Assign the macro to the button
In the Assign Macro dialog:
1. You will see a list of available macros. Look for **LaunchFinanceTools** in the list.
2. Click **LaunchFinanceTools** once to select it (it highlights in blue).
3. Click **OK**.

The button is now connected. Clicking it will run the Finance Tools launcher.

**If LaunchFinanceTools does not appear in the list:**
- Click Cancel
- Go back to Step 7 and repeat the import — the .bas file may not have imported correctly
- Then right-click the button → Assign Macro → find LaunchFinanceTools in the list

---

## PART 5 — Label and style the button

### Step 16 — Edit the button label
The button currently shows default text like "Button 1". Change it to "Finance Tools":

1. **Right-click** the button
2. Click **Edit Text**
3. The text inside the button becomes editable
4. Select all the existing text (Ctrl + A) and delete it
5. Type: `Finance Tools`
6. Click anywhere outside the button to deselect

### Step 17 — Resize and position the button (optional but recommended)
To move the button:
- **Right-click** the button → it enters selection mode (you'll see handles around the edges)
- Hover over the button border until the cursor becomes a 4-arrow move cursor
- Click and drag the button to where you want it

To resize the button:
- **Right-click** the button → it enters selection mode
- Click and drag any of the small squares (handles) on the corners or edges

To center it nicely, aim for somewhere in the B–D column range, rows 3–6.
The exact position is your choice — it just needs to be easy to find on screen.

### Step 18 — Click anywhere on the sheet to deselect the button
After editing, click any empty cell to exit button-edit mode.
The button should now show "Finance Tools" as a plain clickable button.

---

## PART 6 — Test the button

### Step 19 — Enable macros (if prompted)
If there is a yellow bar near the top of the Excel window saying
**"SECURITY WARNING: Macros have been disabled"** — click **Enable Content**.

This only appears the first time you open a file with macros on this machine.

### Step 20 — Test the button
Click the **Finance Tools** button (single click, like a normal button).

**What you should see:**
A command-line window opens. It will show one of two things:

**Option A — Error message (expected for now):**
```
Finance Tools could not start.
Python not found at:
  [path]\python\python-embedded\python.exe
```
This is correct behavior. The workbook is working perfectly — it just can't find Python
yet because the folder structure (python-embedded\ and scripts\) hasn't been set up next
to the .xlsm. That comes in a later step when you assemble the SharePoint zip package.

**Option B — The Finance Tools menu:**
```
============================================================
        Finance Tools — Finance Automation Launcher
============================================================
 1. Revenue Leakage Finder
 2. Data Contract Checker
...
```
This means the folder structure is already in place. Excellent — the whole thing works.

**Either outcome means the button is wired correctly.** The button is done.

### Step 21 — Save the workbook
Press **Ctrl + S**.

If Excel asks "Do you want to keep this format?" — click **Keep Current Format**
(to keep the .xlsm macro-enabled format).

---

## PART 7 — Final checks

### Step 22 — Verify the file is saved as .xlsm
Look at the title bar. It should show: `FinanceTools.xlsm`

If it shows `FinanceTools.xlsx` — you saved in the wrong format.
Go to File → Save As → change the file type to Excel Macro-Enabled Workbook (*.xlsm) → Save.

### Step 23 — Close and reopen to confirm macros persist
1. Close the workbook (Ctrl + W or File → Close)
2. Reopen it from wherever you saved it
3. Click **Enable Content** if the yellow bar appears
4. Click the Finance Tools button again — confirm it still launches the window

If it works after reopen — the VBA is saved inside the file correctly. ✓

---

## What's next after this guide

The workbook is built. The button works. The remaining steps before coworkers can use it:

| Step | What it is | Who does it |
|---|---|---|
| Assemble the zip package | Put FinanceTools.xlsm + python-embedded\ + scripts\ + samples\ in one folder | Connor + Claude |
| Test zero-install path | Confirm the button finds bundled Python and the menu opens | Connor |
| Styling (optional) | Add iPipeline colors, logo, or a brief instruction line above the button | Optional |
| Upload to SharePoint | Put the zip in the pilot SharePoint folder | Connor |

---

## Troubleshooting

**"The macro may not be available in this workbook" error when assigning:**
The .bas file did not import correctly. Go back to Step 7 and repeat the import.
Make sure you are importing from the VBA editor's File menu, not Excel's File menu.

**Button clicks but nothing happens:**
Macros may still be disabled. Look for the yellow "Macros have been disabled" bar and
click Enable Content. If no bar is visible, go to File → Info → Enable Content.

**"LaunchFinanceTools" not in the Assign Macro list:**
Open the VBA editor (Alt + F11) and check that modFinanceToolsLauncher appears under
Modules in the Project Explorer. If it is missing, repeat the import (Step 8).

**VBA editor shows a compile error when you open it:**
Click Debug → it will highlight the problem line. Send a screenshot to Claude for help.

**The .bas file location:**
`C:\Users\connor.atlee\RecTrial\UniversalToolkit\python\ZeroInstall\modFinanceToolsLauncher.bas`

---

*End of guide. Version 1.0 — 2026-04-29.*
*Once the button works, let Claude know and we will move to assembling the SharePoint zip.*
