# Build Video Demo Workbook from Existing Code

## Short Answer
Yes — you can build a **new Excel `.xlsm` demo file** and load your existing VBA code into it for the video demo.

This guide gives a clean, repeatable process you can run before recording so the file looks polished and stable.

---

## What You Need
1. Microsoft Excel (desktop)
2. A copy of this repo
3. VBA source modules from `SourceCode/vba/`
4. Your planned demo dataset

---

## Recommended Build Strategy
Use a **fresh workbook each time** you prepare a major demo recording:
- reduces hidden workbook corruption risk,
- keeps old test artifacts out of the recording,
- makes the demo reproducible.

---

## Step-by-Step (No Coding Required)

### Step 1 — Create the base workbook
1. Open Excel.
2. Create a new blank workbook.
3. Save it immediately as a macro-enabled file (`.xlsm`), for example:
   - `KBT_PnL_Demo_Video_v1.xlsm`

### Step 2 — Enable Developer access
1. Go to **File > Options > Customize Ribbon**.
2. Check **Developer**.
3. Click **OK**.

### Step 3 — Import all VBA modules
1. Press **Alt + F11** to open the VBA Editor.
2. In VBA Editor, right-click the project for your new workbook.
3. Choose **Import File...**.
4. Import each `.bas` module from `SourceCode/vba/`.
5. Import any UserForm/code-behind assets you use for the Command Center.

### Step 4 — Verify references
1. In VBA Editor, go to **Tools > References**.
2. Confirm required references are checked (if your modules require any).
3. Resolve any **MISSING** references before testing.

### Step 5 — Add required worksheets
1. Create all worksheets expected by the macros.
2. Use names expected by the code (exact spelling/case where relevant).
3. Add sample data to key input tabs.

### Step 6 — Configure demo defaults
1. Open any config module (for example `modConfig_*` style modules).
2. Confirm paths and environment-specific constants are safe for demo machine use.
3. Remove or replace machine-specific local paths before recording.

### Step 7 — Compile and smoke test
1. In VBA Editor choose **Debug > Compile VBAProject**.
2. Fix any compile errors.
3. Run your first 5 core demo actions in order.
4. Confirm outputs (dashboard/report/charts) appear correctly.

### Step 8 — Freeze recording build
1. Save as a final recording copy, for example:
   - `KBT_PnL_Demo_Video_RECORDING.xlsm`
2. Keep one backup copy untouched.
3. Record using only the frozen recording copy.

---

## Optional: Fast Iteration Pattern
For repeated video practice:
- Keep one **template workbook** with all modules imported.
- Before each rehearsal, duplicate template and rename with timestamp.
- Run a short preflight checklist (compile + 3 critical commands).

---

## Pre-Recording Checklist (Recommended)
- Workbook opens without macro security prompts blocking flow
- Command Center opens cleanly
- No debug popups or compile errors
- Input data already loaded for quick pacing
- Output tabs/charts render correctly at 100% zoom
- Any audio/animation extras are optional and failure-safe

---

## If You Want, Next
You can create a follow-up guide called:
`VIDEO_DEMO_RUN_OF_SHOW.md`
with exact spoken script, click-by-click timing, and backup plan if a command fails live.
