# OPEN ITEMS — 2026-04-27
## Companion to HANDOFF_2026-04-27.md

This file captures EVERY partial analysis, plan, TODO, and detail discussed in this session that did not get finished. Pair this with `HANDOFF_2026-04-27.md` for full pickup context. **Read both files before starting next session.**

---

## 0. WHAT THIS SESSION ACTUALLY DID (so the next session knows what's "fresh" thinking)

This session did NOT write code. It did:
1. Read all 5 files in `RecTrial\Brainstorm\NewCodeResearchExtra\` (5th-pass external review)
2. Synthesized the review against the prior V4 proposal
3. Delivered a 5-pivot recommendation to Connor (4 to lock, 1 hero call left open)
4. Wrote `RecTrial\HANDOFF_2026-04-27.md`
5. Wrote this file (`RecTrial\OPEN_ITEMS_2026-04-27.md`)
6. Committed the handoff to the `April23CLD` branch as `14ea9e0`

**No repo modules changed. No CLAUDE.md / todo.md / lessons.md changes.** Those were updated 2 commits earlier (`383ba2b`) before this session started.

---

## 1. THE 5 V4 DECISIONS — STATE BY DECISION

### Decision 1 — Single V4 vs split 4a + 4b
- **Recommended in this session:** Single V4, 9–12 min, chaptered
- **Reasoning:** External review consensus. One strong video > two okay videos. Recipe shorts can come later as non-canonical extras.
- **Status:** ⏸ awaiting Connor's lock-in
- **Files affected if locked single:**
  - `tasks/todo.md` Video 4 Replanning section needs rewrite (currently shows split 4a+4b as "current direction")
  - `RecTrial\Brainstorm\VIDEO_4_CURRENT_PROPOSAL.md` will be superseded by the new `VIDEO_4_REVISED_PLAN.md`
- **What gets cut if locked single:** No 4a/4b chapter naming, no separate runbooks, single chapter outline

### Decision 2 — Hero: ARR Waterfall vs Revenue Leakage Finder
- **Recommended in this session:** Middle-ground — Revenue Leakage Finder is the narrative hero, ARR waterfall is the closing visual artifact
- **Reasoning:** The two reviewer files don't fully agree:
  - `01_claude_code_review_brief.md` — "Revenue Leakage Finder IS hero, waterfall is demoted to side artifact"
  - `Claude_Review_ReportOther.md` — "Use ARR waterfall AS visual, frame it AS leakage analysis" (the combine option)
- **Status:** ⏸ awaiting Connor's gut call (the one open decision after the other 4 lock)
- **My honest framing:** "Python found money" beats "Python made a chart." Waterfall is the cleanest visual asset already on the table. Doing both — leakage as the story, waterfall as the closing slide — gets both wins.
- **Risk to flag:** Revenue Leakage Finder is only stronger if the synthetic data feels real. Toy numbers ruin it. Realistic contract/billing structure required.

### Decision 3 — SOX Evidence Collector in or out
- **Status:** ⏸ open from prior session (not addressed in this session's pivots)
- **Connor's earlier note:** depends on whether his team owns SOX evidence work
- **Recommendation:** Default OUT for V1. The Control Evidence Pack (already in the V4 candidate list) covers most of the same ground without explicit SOX framing. Add SOX-specific tooling to v2 if Connor's team takes ownership.

### Decision 4 — xlwings in V4 or parked as v2
- **Recommended in this session:** Parked as v2 (this aligns with prior position)
- **Reasoning:** Locked-down corporate laptops may block xlwings install. Don't make V1 depend on it.
- **Status:** ⏸ technically still open but consensus from all rounds of research is parked
- **Reviewer note:** "Build the CLI for now, but document it as the simplest v1. Treat Excel button / GUI / web front end as the adoption version."

### Decision 5 — Deliverable: CLI menu only, or CLI + xlwings
- **Recommended in this session:** CLI menu only for V1, named `finance_automation_launcher.py` (rename from `finance_copilot.py` because "Copilot" implies AI which we don't have)
- **Status:** ⏸ name change awaiting lock-in
- **Reasoning:** Honest naming. Tool has no AI; don't market it as Copilot.

---

## 2. THE 5 PLANNING DOCS — DETAILED SPEC

These are the deliverables for the next session AFTER Connor locks the decisions. None exist yet.

### 2.1 `VIDEO_4_REVIEW_DECISION_MEMO.md`
- **Path:** `RecTrial\Brainstorm\VIDEO_4_REVIEW_DECISION_MEMO.md`
- **Source brief:** `01_claude_code_review_brief.md` Task A
- **Purpose:** Captures locked decisions and what got cut
- **Required sections:**
  - Audience reframe (2,000+ → 50–150 coworkers near-term)
  - Single V4 vs 4a+4b decision and why
  - Hero choice (Revenue Leakage Finder vs middle-ground waterfall-as-leakage)
  - Public surface narrowing (140 → 5–7 supported workflows)
  - Deliverable name change (`finance_copilot.py` → `finance_automation_launcher.py`)
  - **Stale reference list** — places in CLAUDE.md / docs / planning that still imply 2,000+ audience, CFO/CEO target, 4a/4b structure, finance_copilot.py name, 140 tools as adoption surface. Each location should be flagged with file + section so it can be updated.
- **Estimated length:** ~150–250 lines

### 2.2 `VIDEO_4_REVISED_PLAN.md`
- **Path:** `RecTrial\Brainstorm\VIDEO_4_REVISED_PLAN.md`
- **Source brief:** `01_claude_code_review_brief.md` Task B + Section 3
- **Purpose:** Final V4 production plan
- **Required sections:**
  - Final video title and length target
  - Chapter outline (8 chapters per the review brief Section 3)
  - Demo sequence with timing per chapter
  - Estimated build effort per Python script
  - Estimated recording effort
  - Known tradeoffs (what's good about this plan, what we're giving up)
  - Optional recipe shorts roadmap (non-canonical, post-V4)
- **Reference for chapter outline (from `01_claude_code_review_brief.md`):**
  1. Why Python after Excel/VBA — 45 sec
  2. Safety first — 60 sec
  3. Hero: Revenue Leakage Finder — 2.5 to 3.5 min
  4. Data Contract Checker — 90 sec
  5. Exception Triage Engine — 90 sec
  6. Control Evidence Pack — 90 sec
  7. Launcher — 60 sec
  8. How to start — 30 sec
- **Total: ~9–12 min**

### 2.3 `PYTHON_SAFETY.md`
- **Path:** `RecTrial\UniversalToolkit\python\PYTHON_SAFETY.md` (canonical) OR `RecTrial\Brainstorm\PYTHON_SAFETY_SPEC.md` if no Python code yet
- **Source brief:** `01_claude_code_review_brief.md` Task D + `04_release_safety_distribution_checklist.md` section 4
- **Purpose:** Visible, inspectable safety doc — protects Connor when 50–150 coworkers run scripts
- **Required content (14+ rules):**
  1. No internet calls
  2. No external AI/API calls
  3. No credentials, tokens, secrets, or database connections
  4. Standard library only for new v1 scripts where feasible
  5. Input files are read-only
  6. Scripts never overwrite source files
  7. Outputs go to timestamped folders under `/outputs/`
  8. Every run writes a log
  9. Sample mode is available
  10. Clear failure messages shown to user
  11. Detailed stack traces go to log files, not main user experience
  12. Batch/destructive operations require explicit confirmation
  13. Logs avoid storing sensitive raw data
  14. Launcher includes a visible safety disclaimer
- **Format:** plain English, readable by non-developer coworkers and IT/security reviewers
- **Estimated length:** ~80–120 lines

### 2.4 `MINIMUM_DISTRIBUTION_PLAN.md`
- **Path:** `RecTrial\Brainstorm\MINIMUM_DISTRIBUTION_PLAN.md`
- **Source brief:** `01_claude_code_review_brief.md` Task E + `04_release_safety_distribution_checklist.md` sections 3, 7, 8, 9, 10
- **Purpose:** Concrete SharePoint distribution plan for 50–150 coworkers
- **Required sections:**
  - SharePoint folder/page structure (the 8-file package layout)
  - Launch message draft (subject + body, ready to paste)
  - Support expectations and intake process (Connor owns support — must limit load)
  - Pilot plan — 10–20 users (3–5 Finance + 3–5 Accounting + 3–5 Billing/RevOps + 1–2 managers + optional IT/security observer)
  - Pilot success metrics (10 open package, 5 run sample, 3 try real file, top 3 confusing points identified, top 3 bugs fixed/documented, 2 concrete use cases captured)
  - Final release gate (11-checkpoint table from section 10)
- **Estimated length:** ~150–200 lines

### 2.5 `SUPPORTED_WORKFLOWS_V1.md`
- **Path:** `RecTrial\Brainstorm\SUPPORTED_WORKFLOWS_V1.md`
- **Source brief:** `01_claude_code_review_brief.md` Task F + `04_release_safety_distribution_checklist.md` section 2
- **Purpose:** Map the 5–7 starter workflows to existing VBA/Python modules
- **The 7 supported v1 workflows:**
  1. Clean a messy Excel export → Data Sanitizer / Data Cleaning tools
  2. Compare two files → Sheet Compare / Quick Row Compare
  3. Consolidate sheets/files → Consolidate tools / multi-file consolidator
  4. Find workbook issues → Audit tools / external links / errors
  5. Generate workbook summary → Exec Brief / profile workbook
  6. Find possible revenue leakage → Revenue Leakage Finder (NEW Python)
  7. Check file structure → Data Contract Checker (NEW Python)
- **Required content per workflow:**
  - Description (one paragraph plain English)
  - Which existing module(s) the workflow uses
  - Sample file used to demo it
  - Time-to-value (how fast can a coworker see a result)
  - Prerequisites (none for Excel side; Python install steps for Python side)
- **Everything else** (the other ~135 universal tools) gets labeled "advanced / included for exploration / not the first recommended path"
- **Estimated length:** ~120–180 lines

### Suggested production order
1. `VIDEO_4_REVIEW_DECISION_MEMO.md` (locks captured first)
2. `SUPPORTED_WORKFLOWS_V1.md` (concrete list of what's supported)
3. `VIDEO_4_REVISED_PLAN.md` (chapter outline that references the workflows)
4. `PYTHON_SAFETY.md` (the visible promise)
5. `MINIMUM_DISTRIBUTION_PLAN.md` (the operational wrapper)

---

## 3. THE V4 CODE — DECISION NEEDED ON WHO BUILDS IT

The Codex build spec at `02_codex_build_spec.md` (667 lines) is detailed enough to hand to Codex as-is. But Connor needs to decide who actually builds the V4 Python code:

**Option A — Codex builds.** Hand the spec to Codex. Claude Code reviews. Pros: spec is already Codex-shaped. Cons: another round of cherry-picking.

**Option B — Claude Code builds.** Use the spec as input. Pros: stays in one repo. Cons: spec is long; build effort lands here.

**Option C — Both build, parallel.** Hard to merge cleanly. Probably not it.

**My lean:** Option A. The Codex spec is comprehensive and Codex parallel work has been productive (3 batches cherry-picked successfully). Let Codex build the 6 Python scripts; Claude Code reviews and ports anything that comes out clean.

**Status:** ⏸ open. Connor should decide after the 5 planning docs land.

---

## 4. SCRIPTS NEEDED FOR THE LOCKED V4 DIRECTION

If Connor locks the recommended V4 direction, these 6 Python scripts get built (full specs in `02_codex_build_spec.md`):

| # | Script | Path | Status |
|---|---|---|---|
| 1 | `finance_automation_launcher.py` | `RecTrial\UniversalToolkit\python\ZeroInstall\` | NOT STARTED |
| 2 | `revenue_leakage_finder.py` | `RecTrial\UniversalToolkit\python\ZeroInstall\` | NOT STARTED |
| 3 | `data_contract_checker.py` | `RecTrial\UniversalToolkit\python\ZeroInstall\` | NOT STARTED |
| 4 | `exception_triage_engine.py` | `RecTrial\UniversalToolkit\python\ZeroInstall\` | NOT STARTED |
| 5 | `control_evidence_pack.py` | `RecTrial\UniversalToolkit\python\ZeroInstall\` | NOT STARTED |
| 6 | `workbook_dependency_scanner.py` | `RecTrial\UniversalToolkit\python\ZeroInstall\` | NOT STARTED |

**Plus common utilities:**
- `safe_io.py` — read-only loaders, output folder timestamping
- `logging_utils.py` — `run_log.json` + `run_summary.txt` writers
- `sample_data.py` — synthetic data generators
- `report_utils.py` — Excel/CSV writers with consistent formatting

**Plus sample data files** under `samples/`:
- Billing export (synthetic, realistic)
- Contract / expected revenue file
- Customer mapping
- Workbook to scan for dependencies
- Bad-file (failing data contract) for Checker demo
- Good-file (passing data contract) for Checker demo

**Plus smoke test:**
- `smoke_test_video4_python.py` — runs all 6 scripts in sample mode and verifies outputs exist

---

## 5. SYNTHETIC DATA QUALITY — UNRESOLVED CONCERN

The 5th-pass review highlights this risk repeatedly: Revenue Leakage Finder is only stronger than the waterfall hero IF the synthetic data feels real. Toy numbers ruin it.

**What "feels real" means here:**
- Realistic customer names and IDs
- Realistic contract values (range, distribution, contract terms)
- Plausible billing variance patterns (multiple legitimate reasons for variance)
- Edge cases that look like real edge cases — not contrived examples
- Volume — at least a few hundred rows, not 12 toy rows
- Some genuine "leakage" cases scattered in the data, not a single obvious one

**Open question:** Who designs the synthetic data?
- Codex spec sketches it but doesn't fully specify
- Connor knows iPipeline's actual revenue/billing patterns
- A bad synthetic file kills the demo

**Recommendation for next session:** Before the 6 Python scripts get built, lock the sample data design. Connor should describe (in plain English) what realistic billing variance looks like at iPipeline so the synthetic data captures it.

---

## 6. STALE REFERENCES TO FIX (after V4 decisions lock)

These should NOT be edited until Connor locks the decisions, because they may need different edits depending on what gets locked.

### `CLAUDE.md`
- Line ~"## The Project" — says "2,000+ employees and the CFO/CEO" — needs update to "near-term audience: 50–150 coworkers; broader rollout deferred"
- Section "## ⚡ CURRENT WORK" — says "Video 4 ('Python Automation for Finance') — audio clips generated, demo files built, ready to record manually" — STALE. Original V4 was pulled. New V4 in planning.
- "Video 4 replanning" subsection — currently describes split 4a + 4b plan with ARR Waterfall hero. Needs full rewrite once new direction locks.
- All references to `finance_copilot.py` need to change to `finance_automation_launcher.py`

### `Archive/tasks/todo.md`
- "## Video 4 Replanning" section (lines ~34–71) — describes the OLD direction (split 4a+4b, finance_copilot.py, ARR Waterfall hero). Needs full rewrite once new direction locks.
- "V4 open decisions (awaiting Connor)" — needs to reflect the 5 decisions in their current state (4 recommended-to-lock, 1 open hero call)
- "V4 build list" — needs to reflect the 6 new scripts in correct names
- The "Video 4 — Ready to Record" section (lines ~95–101) is stale — that was the original CMD-based plan that got pulled. Should be archived to a "Pulled / Archived" section, not deleted (history matters).

### `Archive/tasks/lessons.md`
- Could add new lessons from this session:
  - **External reviewers don't always agree — extract disagreements explicitly.** Why: 01_claude_code_review_brief.md and Claude_Review_ReportOther.md disagreed on hero framing. The disagreement itself was the most useful signal. How to apply: when synthesizing multiple external inputs, surface where they conflict, not just where they agree.
  - **Don't run a third research round when the second one already pivots the plan.** Why: 14 raw research files + 6 syntheses + 5 fresh review files is enough. Adding more would delay decisions, not improve them. How to apply: after two rounds of research, push for decisions, not more research.
  - **A reviewer recommendation that conflicts with stale CLAUDE.md framing is signal CLAUDE.md is stale, not that the reviewer is wrong.** Why: this session showed the 2,000+ audience number was always near-term aspirational and the reviewer correctly called it out. How to apply: when external review recommends pivots, audit CLAUDE.md for stale framing before defending the prior direction.

### `RecTrial\PROJECT_OVERVIEW.md`
- Was written before the 5th-pass review landed. Likely has stale references. Should be reviewed and updated to point at the locked V4 direction once decisions land.

### Memory (`...\memory\project_status.md`)
- Says V4 is "ready to record manually" with the OLD direction. Stale. Should be updated to point to the handoff and reflect "V4 in planning, awaiting Connor's lock-in on 5 decisions."

---

## 7. THINGS DISCUSSED IN THIS SESSION THAT DIDN'T GO ANYWHERE (BUT WORTH FLAGGING)

### Two reviewer files disagree
- `01_claude_code_review_brief.md` says Revenue Leakage Finder IS the hero
- `Claude_Review_ReportOther.md` says use ARR waterfall AS visual but frame as leakage
- I recommended the middle-ground (waterfall closing artifact). Connor hasn't decided.
- **Don't lose this disagreement** — it's the most useful piece of signal in the review files.

### Codex spec is buildable but long
- 667 lines of spec
- Specifies file layout, common utilities, sample data, smoke test, acceptance criteria per script, build order
- Could go directly to Codex with no further translation
- Open question: does Connor want Codex to build it, or Claude Code, or both?

### The "5–7 workflows" surface narrowing is partly already done
- The Intelligence module is now visible at Command Center position 6 (last session's work)
- That alone narrows the visible doorway to "the static categories"
- The auto-discovery fallback still surfaces all ~140 tools, but only AFTER the static list
- This is good — it matches the reviewer's "narrow doorway, full toolkit still discoverable" recommendation
- **Implication:** less Command Center work needed than the reviewer might assume

### "Don't release 140 tools" was over-stated by the reviewer
- We don't NEED to remove tools — they're in the .xlsm
- We need to STOP LEADING with them
- Documentation and onboarding lead with the 5–7. Advanced surface stays discoverable for power users.
- This is a doc/positioning fix, not a code fix.

### Pilot plan adds new structure we didn't have
- 10–20 user pilot before broader 50–150 release
- Mix: 3–5 Finance + 3–5 Accounting + 3–5 Billing/RevOps + 1–2 managers + optional IT/security observer
- Pilot success metrics defined
- This is a brand new layer. Connor needs to identify the 10–20 people. That's a real-world task, not a Claude Code task.

### Connor's earlier 5-question constraint answers (from prior sessions)
The user previously answered constraint questions:
1. **C** (option C, whatever C was for that question)
2. **Prefer C, but open to B** (option C with fallback B)
3. **Can pip install freely** (Connor has install permissions on his machine)
- These answers are referenced in earlier session notes but the exact questions/answers should be re-verified before assuming they apply to V4 decisions.

### What's NOT in the 5 planning docs but might come up
- **Approved Python packages list** — pandas, openpyxl, pdfplumber, python-docx, thefuzz, numpy, matplotlib, xlwings (parked), stdlib. The Codex spec leans heavily stdlib-only. If Connor's coworkers don't have pip access, that's a hard constraint that drops several packages.
- **iPipeline brand styling for the Python output** — Excel reports generated by Python should match the brand guide at `docs/ipipeline-brand-styling.md`. Codex spec doesn't fully address this.
- **Coworker laptop install reality** — locked-down laptops may block xlwings, may block pip, may not have Python installed at all. If even pip is blocked, the ZeroInstall stdlib-only scripts become the only option.

---

## 8. THINGS THE NEXT SESSION SHOULD ASK CONNOR

In order of importance:

1. **"Did you lock the four V4 decisions?"** (audience 50–150, single video, rename to `finance_automation_launcher.py`, narrow public surface to 5–7 workflows)
2. **"Did you make the hero call?"** (Revenue Leakage Finder pure vs middle-ground waterfall-as-leakage)
3. **"Who builds the V4 Python — Codex, Claude Code, or both?"**
4. **"What does realistic billing variance look like at iPipeline?"** (for synthetic data design — only matters if Revenue Leakage Finder is the hero)
5. **"Do you want SOX Evidence Collector in V1 or parked?"** (decision 3 was never resolved)
6. **"Who are the 10–20 pilot users?"** (real-world task; can wait until docs are written)
7. **"Are coworkers on locked-down laptops without pip access?"** (changes script architecture massively — affects whether stdlib-only is mandatory)

---

## 9. THINGS NOT TO DO IN NEXT SESSION

- ❌ Don't start a third research round
- ❌ Don't write V4 Python code before the 5 planning docs land
- ❌ Don't update CLAUDE.md / todo.md until Connor locks decisions (otherwise they get rewritten twice)
- ❌ Don't propose splitting V4 into 4a + 4b again unless Connor explicitly reopens
- ❌ Don't propose adding new tools to the universal toolkit during V4 planning (surface narrowing is the goal)
- ❌ Don't touch the 8 protected Python scripts (`aging_report.py`, `bank_reconciler.py`, `compare_files.py`, `forecast_rollforward.py`, `fuzzy_lookup.py`, `pdf_extractor.py`, `variance_analysis.py`, `variance_decomposition.py`) before V4 records
- ❌ Don't propose SendKeys for any new VBA dialog automation (Path A pattern only)
- ❌ Don't re-edit modConfig color constants (existing values are load-bearing in working code)
- ❌ Don't re-edit `RecTrial\PROJECT_OVERVIEW.md` until V4 direction locks (snapshot will need rewrite anyway)

---

## 10. THINGS TO MENTION TO CONNOR THAT WERE NOT EXPLICITLY DISCUSSED

These came up in the review but didn't get full air time:

- **The Codex review file (`Claude_Review_ReportOther.md`) suggests a slightly different post-V4 roadmap than we've discussed.** It moves up: dual logging pattern, CONSTRAINTS.md, BRAND.md, RELEASE_READINESS_CHECKLIST.md, TROUBLESHOOTING.md, Workbook Policy Validator, Dependency Impact Preview, Auto-Repair Suggestions. Most of these match Batches 4–5 already deferred. The new ones are Workbook Policy Validator, Dependency Impact Preview, Auto-Repair Suggestions — those are worth tracking as v2 candidates.
- **The reviewer suggests creating a signed Excel add-in if the VBA toolkit is meant to be widely adopted.** This is a real packaging step we haven't planned. For 50–150 coworkers it's probably not worth signing yet, but for any company-wide rollout it would matter. Park as v2.
- **The reviewer asks: "Are there existing conventions for CLI arguments, output folders, and sample files?"** Yes — defined informally across existing scripts. Should be formalized in a `CLI_CONVENTIONS.md` doc once the 5 V4 scripts are built. Add to v2 docs backlog.
- **The reviewer asks: "Which docs are source-of-truth: CLAUDE.md, Archive/tasks/todo.md, or the RecTrial/Brainstorm docs?"** Answer: CLAUDE.md is authoritative for project conventions, todo.md is the running task list, RecTrial/Brainstorm/* is planning artifacts. The PROJECT_OVERVIEW.md is a point-in-time master narrative. This should be documented somewhere obvious — maybe at the top of CLAUDE.md.

---

## 11. FILES THIS SESSION TOUCHED

| File | Action | Notes |
|---|---|---|
| `C:\Users\connor.atlee\RecTrial\HANDOFF_2026-04-27.md` | CREATED | 12-section handoff doc |
| `RecTrial\HANDOFF_2026-04-27.md` (in repo) | COPIED + COMMITTED | Commit `14ea9e0` on `April23CLD` branch |
| `C:\Users\connor.atlee\RecTrial\OPEN_ITEMS_2026-04-27.md` | CREATED | THIS FILE |

That's it. No code changes. No CLAUDE.md / todo.md / lessons.md changes. No memory updates yet (will do separately if Connor approves).

---

## 12. EVERY DETAIL DISCUSSED — RAW DUMP

For maximum fidelity to "I want next chat to know every single detail":

### From the 5th-pass review files
- `README.md` (21 lines) — summary of the 4 handoff files
- `01_claude_code_review_brief.md` (306 lines) — 6 tasks A–F for Claude Code
- `02_codex_build_spec.md` (667 lines) — buildable Python spec for 6 scripts + common utilities + sample data + smoke test
- `03_explain_like_im_5.md` (318 lines) — plain English version
- `04_release_safety_distribution_checklist.md` (234 lines) — release/safety/distribution layer
- `Claude_Review_ReportOther.md` (167 lines) — slightly different framing on hero (combine option)

### Specific recommendations from those files (consolidated)
- Audience pivot: 2,000 → 50–150 coworkers
- Single V4, 9–12 min, chaptered (8 chapters)
- Revenue Leakage Finder as hero (or waterfall-framed-as-leakage)
- Public surface narrows to 5–7 supported workflows
- Rename `finance_copilot.py` → `finance_automation_launcher.py`
- Add `PYTHON_SAFETY.md` with 14 rules
- SharePoint package layout (8 files)
- Pilot plan (10–20 users → 50–150)
- 11-checkpoint final release gate
- Park xlwings, external AI, IT enterprise deployment for v2
- Build `CONSTRAINTS.md`, `BRAND.md`, `RELEASE_READINESS_CHECKLIST.md`, `TROUBLESHOOTING.md` post-V4
- Track adoption and issues with simple spreadsheet
- Connor owns support — must protect his support load

### My delivered recommendation (in chat to Connor)
- Lock 4: audience 50–150 / single V4 / rename launcher / 5–7 workflows
- Open: hero call (pure leakage vs middle-ground waterfall-as-leakage)
- Push back: "don't release 140 tools" is over-stated — keep them in the .xlsm, just don't lead with them
- Push back: Revenue Leakage Finder only stronger if synthetic data feels real
- Push back: 5 docs is right but produce them tight, not another sprawl

### Connor's most recent state when summary was triggered
- Asked me to review the 5 new files in NewCodeResearchExtra
- Said "dont lose track of the last message you had about the video 4 prior" (the 5 V4 decisions checkpoint)
- After my synthesis, decided to start a fresh chat → asked for handoff doc → asked for this open-items doc

### Earlier session context preserved in memory
- 14 raw research files + 6 compiled syntheses already exist (first round)
- 156 ideas inventoried
- Codex parallel build at `tug83535/AP_CodexVersion`
- Batches 1–3 cherry-picked and live; Batches 4–5 deferred
- Branch April23CLD pushed with RecTrial snapshot
- Intelligence module live at Command Center position 6
- Narrative-on-wrong-column bug fixed (commit 8eff337)
- Path A silent wrapper pattern is canonical for dialog automation
- iPipeline brand colors documented in `modUTL_Branding.bas` header (no named constants)
- 8 V4 Python scripts protected from edits until V4 records
- modConfig color constants are load-bearing — don't edit
- LogAction signature bug pattern (13 historical instances)
- Gemini misperceives colors/labels — usually not real bugs

---

**END OF OPEN ITEMS — 2026-04-27**

Read `HANDOFF_2026-04-27.md` first, then this file. Together they should give the next session full picture.
