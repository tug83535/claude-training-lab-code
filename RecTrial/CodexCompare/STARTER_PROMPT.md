# Starter Prompts — Copy/Paste Reference

This file holds the exact prompts the user will paste into Codex at each stage. Codex itself can read this file for context — but the prompts are triggered by the user in codex.openai.com, one task at a time.

---

## STAGE 1 — Brief Intake + Project Plan

Copy-paste the block below into Codex:

```
You are about to build a world-class Finance & Accounting automation
demo project from scratch for iPipeline. Everything you need is in this
repository — do not assume anything beyond what's written in the
brief files.

STEP 1 — READ, in this exact order, in full:
  1. README.md
  2. CONTEXT.md
  3. CONSTRAINTS.md
  4. BRAND.md
  5. PLAN.md (the template you will fill)

STEP 2 — INVENTORY both files in samples/:
  - samples/ExcelDemoFile_adv.xlsm
  - samples/Sample_Quarterly_ReportV2.xlsm
For each: list every sheet, visibility, purpose, data shape, named
ranges, and any existing VBA. Do not modify either file.

STEP 3 — FILL PLAN.md by replacing every <<FILL IN>> block with your
proposed answers. Preserve the section structure. Commit PLAN.md.

STEP 4 — DO NOT WRITE ANY OTHER CODE YET. Wait for the user to reply
with the literal word "approved" before Stage 2.

Rules while you work:
  - Plain English in all user-facing text
  - Ask clarifying questions rather than guessing — list them in
    PLAN.md section 8 (Open Questions)
  - Follow BRAND.md exactly for any visual output
  - Every feature must pass the CONSTRAINTS.md test: "Could a user
    do this in <5 native clicks?" If yes, don't build it.
```

---

## STAGE 2+ — Build Stages

Use this template (edit the stage number + scope each time):

```
Stage N. Re-read README.md, CONTEXT.md, CONSTRAINTS.md, BRAND.md,
and PLAN.md (which has been approved by the user).

Scope for this stage:
<PASTE THE SPECIFIC MODULES / FEATURES / GUIDES / SCRIPTS FROM PLAN.md>

Rules:
  - Build exactly what's in scope. Do not expand scope without asking.
  - Follow PLAN.md — if you need to deviate, stop and ask.
  - Follow BRAND.md for all visual output.
  - Every user-facing string must be plain English.
  - Self-review before committing. Check: Does each file compile / run?
    Does it handle realistic Finance-file edge cases?
  - Commit when finished. Do not move to a future stage.
```

---

## REVISION PROMPTS (when Codex's output needs fixes)

```
The output from Stage N has issues. Do not re-build from scratch.
Fix only the following:

1. <specific issue>
2. <specific issue>
3. <specific issue>

Do not change anything outside of these issues. Commit the fix.
```

---

## SCOPE-CHANGE PROMPT (when PLAN.md needs updating mid-project)

```
Stop. I need to change the plan before the next build stage.

Update PLAN.md with the following changes:
  - <change 1>
  - <change 2>

Do not execute the changes yet — just edit PLAN.md and mark the
affected section as "(REVISED <date>)". Wait for my "approved"
before building under the new plan.
```
