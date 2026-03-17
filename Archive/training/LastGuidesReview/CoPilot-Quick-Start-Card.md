# CoPilot Prompt Guide — Quick Start Card

## What Do You Want To Do?

Use this card to jump straight to the right prompt in the CoPilot Prompt Guide v2.0.

---

### "I want Copilot to analyze MY file and tell me what to fix or automate"

**Use the Recommended All-in-One Prompt** (page 3 of the guide)

**Steps:**

1. Open M365 Copilot (or your AI tool of choice)
2. Upload your Excel file
3. Copy and paste the **Recommended All-in-One Prompt** from the guide
4. Copilot will tell you:
   - What the file contains
   - What could be improved
   - What VBA macros or Python scripts it recommends building
   - Step-by-step instructions to implement each recommendation

**Alternative:** If you want an even more thorough analysis, use the **Mega Prompt** at the end of Chapter 2. It runs a full end-to-end workflow covering purpose, assumptions, errors, fixes, and improvements all in one shot.

---

### "I found a macro in the demo file I like — how do I make it work on MY file?"

**Use Section C: Make Code Fit Your Workbook** (Chapter 2, prompts C1 through C4)

**Steps:**

1. Open the demo file in Excel
2. Press Alt+F11 to open the VBA Editor
3. Find the module you want (see the **VBA Module Reference List** for what each module does)
4. Right-click the module and choose **Export File** to save it as a .bas file
5. Open M365 Copilot
6. Upload TWO things:
   - Your Excel file (the one you want the code to work on)
   - The .bas file you just exported
7. Copy and paste the prompt that matches your situation:

| Your Situation | Use This Prompt |
|---|---|
| My sheets have different names than the code expects | **C2** — Sheet names are different |
| My columns are in a different order or have different headers | **C3** — Columns are different |
| I am not sure what is different — just make it work | **C1** — Map workbook to code requirements |
| My file has multiple tables or sections on one sheet | **C4** — Multiple tables/sections |

8. Copilot will rewrite the code to match YOUR file and give you step-by-step instructions to import and run it

---

### "The code I got from the demo file is throwing an error on my file"

**Use Section B: Error and Bug Fix Prompts** (Chapter 2, prompts B1 through B4)

| Your Situation | Use This Prompt |
|---|---|
| VBA error popup when I run the macro | **B1** — Fix a VBA error |
| Python script crashes or shows an error | **B2** — Fix a Python error |
| Code runs but gives wrong numbers or output | **B3** — Fix wrong output |
| Code works on the demo file but not mine | **B4** — Works on one file but not another |

**Steps:**

1. Open M365 Copilot
2. Upload your Excel file AND the code file (.bas or .py)
3. Also paste the exact error message you saw (copy it from the popup or terminal)
4. Copy and paste the matching prompt from the table above
5. Copilot will find the root cause, fix the code, and give you step-by-step instructions

---

### "I want to learn what a macro does before I use it"

**Use Section G: Learning Prompts** (Chapter 2, prompts G1 through G3)

| Your Goal | Use This Prompt |
|---|---|
| Explain what the code does in plain English | **G1** — Explain code like I am new |
| Learn how to safely edit parts of the code myself | **G2** — Teach me to safely edit it |
| Get a mini tutorial with examples | **G3** — Create a mini tutorial |

---

### "I want to convert a VBA macro to Python (or vice versa)"

**Use Section H: Conversion Prompts** (Chapter 2, prompts H1 through H3)

| Your Goal | Use This Prompt |
|---|---|
| Convert VBA to Python | **H1** |
| Convert Python to VBA | **H2** |
| Not sure which is better for my task | **H3** — Recommend VBA, Python, or both |

---

## Tips for Best Results

1. **Always upload your actual file** — Copilot gives much better answers when it can see your real data and structure
2. **Copy the full prompt** — Do not shorten or paraphrase the prompts. They are written specifically to get Copilot to give you complete, step-by-step answers
3. **One task at a time** — If you need multiple things, run each prompt separately for best results
4. **Save your work first** — Always save a backup copy of your file before running any new macro
5. **If the answer is confusing** — Use prompt G1 ("Explain like I am new") on whatever code Copilot gave you
