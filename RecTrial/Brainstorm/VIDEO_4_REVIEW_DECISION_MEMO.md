# Video 4 — Review & Decision Memo

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Date locked:** 2026-04-28
**Supersedes:** `RecTrial\Brainstorm\VIDEO_4_CURRENT_PROPOSAL.md` (split 4a+4b + ARR Waterfall hero + `finance_copilot.py`)
**Trigger:** 5th-pass external review at `RecTrial\Brainstorm\NewCodeResearchExtra\` recommended five pivots. Connor reviewed and locked them on 2026-04-28.

---

## 1. Why this memo exists

The original Video 4 plan (10 CMD-run Python scripts + ElevenLabs audio) was pulled on 2026-04-22. A redesigned proposal landed on 2026-04-23 splitting V4 into 4a + 4b with an ARR Waterfall hero and a `finance_copilot.py` deliverable. Two rounds of external research (14 raw + 6 synthesis files in the first round, 5 files in the second) tested that proposal and recommended a different direction. This memo captures what got locked, what got cut, and every spot in the project's docs that needs a follow-up edit so it stops contradicting the new direction.

---

## 2. The five locked decisions

### Lock 1 — Audience reframe: 50–150 coworkers near-term, NOT 2,000+
**Old framing:** "demo for 2,000+ employees and the CFO/CEO at iPipeline." That number was always near-term aspirational. The 5th-pass review correctly flagged it as too broad for the actual rollout plan.

**New framing:** Real near-term audience is approximately **50–150 coworkers** in Finance, Accounting, and adjacent operations. Broader rollout (and any company-wide CFO/CEO showcase) is deferred — not cancelled, just not the target this V4 has to hit.

**Why it matters:** every doc, every script, every distribution decision was tuned for "everybody." That over-scoped the finale. Resizing the audience downward makes V4 simpler and more defensible — and protects Connor's support load (he owns support personally).

---

### Lock 2 — Single chaptered V4, NOT split into 4a + 4b
**Old framing:** Video 4a "Python Shows You What Excel Can't" (6–7 min, CFO-led) + Video 4b "Your Python Cookbook" (5–6 min, coworker-led).

**New framing:** **One Video 4 of 4 — "Python Automation for Finance," 9–12 minutes, chaptered.** Eight chapters in one video. Optional recipe shorts can ship later as non-canonical extras.

**Why it matters:** the 4a/4b split fragmented the narrative — 4a was an executive flyby, 4b was a recipe sequence, and they didn't naturally connect. One strong video that walks coworkers from "why Python after Excel/VBA?" through a concrete revenue-leakage demo and out to "how do I start?" tells one story end to end. That holds attention better and is easier to defend as a polished, world-class artifact.

---

### Lock 3 — Rename `finance_copilot.py` → `finance_automation_launcher.py`
**Old framing:** `finance_copilot.py` — menu-driven CLI wrapping all 28 existing scripts plus the new V4 scripts.

**New framing:** **`finance_automation_launcher.py`** with the user-facing label **"Finance Automation Launcher."** Function is the same — a simple menu that runs sample workflows.

**Why it matters:** "Copilot" implies AI. The tool has none. Honest naming protects the trust Connor needs from a non-developer audience and from any IT/security review. Coworkers who download "Finance Copilot" and discover there's no AI inside will rightly distrust the rest of the package.

---

### Lock 4 — Lead with 5–7 supported workflows, NOT 140 tools as the adoption surface
**Old framing:** universal toolkit's ~140 VBA tools + 28 Python scripts presented as the adoption surface. Coworkers point-and-shoot from a 140-deep menu.

**New framing:** public documentation, the launch message, the Quick Reference Card, and the SharePoint package all lead with **5–7 supported starter workflows.** The full toolkit stays inside `Sample_Quarterly_ReportV2.xlsm` and stays discoverable via the Command Center's auto-discovery fallback — but it is not the doorway. Everything outside the 5–7 is labeled "advanced / discoverable / not the first recommended path."

**Why it matters:** a 140-deep menu is paralysis, not a feature. Connor owns support for every coworker who tries something and gets stuck. Narrowing the doorway protects support load AND makes the "what should I try first?" question easy to answer.

The seven starter workflows are detailed in `RecTrial\Brainstorm\SUPPORTED_WORKFLOWS_V1.md`. Summary:
1. Clean a messy Excel export
2. Compare two files
3. Consolidate sheets/files
4. Find workbook issues and external links
5. Generate a workbook / executive summary
6. Run Revenue Leakage Finder on sample data (then on your own file)
7. Run Data Contract Checker on sample data (then on your own file)

---

### Lock 5 — Adoption-grade framing (NEW — Connor added this on 2026-04-28)
**This was not in the 5th-pass review.** Connor added it during the lock-in conversation.

**Framing:** The 5–7 supported workflows ship as **adoption-grade tools**, not just demo material. Coworkers are expected to actually use the VBA, Python, and SQL on their own files and in their own workflows. The universal toolkit (modUTL_*) is designed to drop into any workbook coworkers already have. The Python pack runs on coworker machines against their own real (non-sensitive) files.

**Three things this raises the bar on:**
1. **Documentation** — must support self-service adoption, not just demo comprehension. Quick Reference Card, Start Here, Troubleshooting all need to address "I'm now using this on my own file."
2. **Distribution** — the SharePoint package is download-and-use, not download-and-watch. Sample files first; real files next.
3. **Support load** — real adoption produces real bugs, real questions, real requests. Pilot plan and intake process matter more than they would for a watch-only release. Connor owns support — that constraint is now load-bearing for the whole plan.

**Implication:** the v2 Excel-button / xlwings / GUI path matters more under adoption pressure than it did under watch-only framing. Still parked for v1 (locked-down corporate laptops may block xlwings install) — but its priority moves up for v2 once V1 ships.

---

## 3. Hero call — middle-ground (Revenue Leakage Finder + ARR waterfall artifact)

The two reviewer files inside `NewCodeResearchExtra\` did not fully agree:
- `01_claude_code_review_brief.md` recommended Revenue Leakage Finder as the hero, ARR waterfall demoted to side artifact.
- `Claude_Review_ReportOther.md` recommended using ARR waterfall AS the visual but framing it AS leakage analysis (combine).

**Decision:** **middle-ground (b).** Revenue Leakage Finder is the narrative hero — the "Python found money" story. The ARR waterfall ships as the closing visual artifact in the same chapter, framed as the executive-readable summary of the leakage analysis. We get both wins: the strong "Python found a possible billing problem" story AND the clean visual that's easy to screenshot for a CFO conversation later.

**Why this beats either pure choice:**
- Pure leakage hero only works if the synthetic data feels real. Toy numbers ruin it.
- Pure waterfall hero risks "Python made a chart" — visually pretty, narratively flat.
- Middle-ground keeps the strong narrative AND the strong visual. If the synthetic data underperforms, the waterfall still earns its 60–90 seconds of screen time.

**Risk flag — synthetic data quality.** Revenue Leakage Finder is only stronger than the waterfall hero if the sample contracts/billing files feel like real Finance data. This is the single biggest unsolved item before any V4 code is built. Sample data design will be locked in `RecTrial\Brainstorm\VIDEO_4_REVISED_PLAN.md` (next doc in this sprint) — Connor will describe realistic iPipeline billing variance patterns in plain English, and the sample data spec will codify them. **No V4 Python is written until that lock lands.**

---

## 4. What got cut

| Cut | Replaced by | Why |
|---|---|---|
| Split 4a + 4b | Single chaptered V4 | One end-to-end story holds attention. Recipe shorts can ship later as non-canonical. |
| `finance_copilot.py` name | `finance_automation_launcher.py` | "Copilot" implies AI; the tool has none. Honest naming. |
| 140 tools as adoption surface | 5–7 supported workflows | Connor owns support; narrow the doorway. Toolkit stays in the .xlsm but isn't the way coworkers learn the system. |
| "2,000+ employees + CFO/CEO" framing | "50–150 coworkers near-term" | Realistic near-term scope. CFO/CEO showcase deferred, not cancelled. |
| ARR Waterfall as hero | Revenue Leakage Finder as narrative hero, waterfall as closing visual | "Python found money" beats "Python made a chart." Middle-ground keeps both wins. |
| SOX Evidence Collector in V1 | Parked for v2 | Control Evidence Pack covers most of the same ground without explicit SOX framing. If Connor's team takes ownership of SOX work later, it goes in v2. |
| xlwings Excel Button Edition in V1 | Parked for v2 | Locked-down corporate laptops may block install. Don't make V1 depend on it. Adoption pressure (Lock 5) makes this matter more for v2. |

---

## 5. What stayed

- **No external AI / API calls in V1 scripts** (locked in earlier rounds; reaffirmed)
- **Stdlib-only as the safe default** for new V4 scripts. Pandas / openpyxl / pdfplumber / etc. allowed in scripts only after Connor confirms coworkers have pip access. (Connor: yes pip access on his machine. Coworkers: TBD — open item.)
- **iPipeline brand styling** required on every Excel/HTML/PDF output. Authoritative guide at `docs\ipipeline-brand-styling.md`.
- **Path A silent wrapper pattern** is canonical for any dialog-heavy VBA. Never `SendKeys` against modal dialogs.
- **modConfig color constants** are load-bearing in working code. New visual work uses the brand guide; existing modConfig values are not edited.
- **The 8 protected V4 Python scripts** stay frozen until V4 records: `aging_report.py`, `bank_reconciler.py`, `compare_files.py`, `forecast_rollforward.py`, `fuzzy_lookup.py`, `pdf_extractor.py`, `variance_analysis.py`, `variance_decomposition.py`.

---

## 6. Stale-reference table

These five files contain framing that contradicts the locked V4 direction. **Do not edit them yet** — the rest of the planning sprint may surface additional changes. The plan is to edit all of them in one cleanup pass AFTER all 5 docs in this sprint land. That avoids the edit-twice trap.

### 6.1 `c:\Users\connor.atlee\.claude\projects\claude-training-lab-code\CLAUDE.md`

| Line(s) | Current text | Recommended new wording |
|---|---|---|
| 60–64 | "I am building a world-class demo P&L Excel file with VBA macros, SQL, and Python to present to **2,000+ employees and the CFO/CEO** at iPipeline." | "I am building a world-class demo + adoption-grade Finance automation package — Excel/VBA, Python, and SQL — for **50–150 coworkers near-term in Finance, Accounting, and adjacent operations** at iPipeline. Broader rollout and any CFO/CEO showcase are deferred. Coworkers are expected to actually adopt the tools on their own files, not just watch the demo videos." |
| 94 | "**Scenario 1 (Primary — Demo + coworkers):** … This is the plan for the **CFO/CEO demo and general coworker access**." | "**Scenario 1 (Primary — Adoption + coworkers):** Share the finished `.xlsm` directly. All 39 VBA modules + 5 optional add-ins are inside it. Coworkers open the file and use the Command Center. Lead with the 5–7 supported starter workflows; full toolkit stays discoverable. CFO/CEO showcase is a v2 concern." |
| 159 | "Would the CFO/CEO be proud to see this?" | "Would a Finance & Accounting coworker — and eventually the CFO/CEO — be proud to see this?" (keeps the quality bar; reflects the actual primary audience) |
| Section "## ⚡ CURRENT WORK (2026-04-16) — 4-VIDEO DEMO RECORDING PROJECT" (entire section) | Describes V4 as "audio clips generated, demo files built, ready to record manually" using the original CMD-based 10-script plan. | Rewrite to reflect the new V4 direction (single chaptered V4, Revenue Leakage Finder hero, `finance_automation_launcher.py`, 5–7 workflow doorway, adoption-grade framing). Reference this memo + the sprint's other 4 docs as the source of truth. |
| Section "### Video 4 replanning (2026-04-22 → 2026-04-23)" | Describes split 4a+4b plan with ARR Waterfall hero. | Rewrite to reflect locks 1–5 + middle-ground hero. Keep the planning-doc list updated to include the 5 new sprint docs. |

### 6.2 `c:\Users\connor.atlee\.claude\projects\claude-training-lab-code\Archive\tasks\todo.md`

| Line(s) | Current text | Recommended action |
|---|---|---|
| 34–63 (Video 4 Replanning section) | Lists split 4a+4b as "current direction," `finance_copilot.py` as deliverable, ARR Waterfall as hero, V4 open decisions including 4a+4b approval, original 6-script V4 build list. | **Rewrite the entire section** to reflect locks 1–5 + middle-ground hero + adoption-grade framing + the 6 revised script names (`finance_automation_launcher.py`, `revenue_leakage_finder.py`, `data_contract_checker.py`, `exception_triage_engine.py`, `control_evidence_pack.py`, `workbook_dependency_scanner.py`) + the 5 sprint docs as planning artifacts. |
| 95–101 ("Video 4 — Ready to Record (Manual Recording)" subsection) | Lists install steps + recording steps for the OLD CMD-based plan that was pulled. | **Move to a "Pulled / Archived" subsection** — don't delete (history matters). Replace with a stub pointing to the new V4 plan. |
| 256, 259 (Time Saved Calculator + What If Scenario items) | Marketed as "great talking point for **CFO/CEO**" / "speaks the **CFO's language** directly" | Rewrite to "great talking point for managers and finance leadership" / "speaks the language a Finance manager will understand." Tools themselves stay — the framing changes. |

### 6.3 `C:\Users\connor.atlee\RecTrial\AGENTS.md`

| Line(s) | Current text | Recommended new wording |
|---|---|---|
| 4 | "This is a Finance & Accounting automation demo for iPipeline (SaaS, insurance industry). It combines Excel VBA, Python, and SQL to showcase what's possible in Finance workflows. **Audience: 2,000+ employees + CFO/CEO**." | "This is a Finance & Accounting **adoption-grade automation package + demo videos** for iPipeline. It combines Excel VBA, Python, and SQL. **Near-term audience: 50–150 coworkers** in Finance, Accounting, and adjacent operations. Broader rollout and any CFO/CEO showcase are deferred. Coworkers are expected to actually use the tools on their own files." |

### 6.4 `C:\Users\connor.atlee\RecTrial\PROJECT_OVERVIEW.md`

This file is a "point-in-time master narrative." It was written 2026-04-23 — before the 5th-pass review. It needs a near-full rewrite. Key changes:

| Section | Change |
|---|---|
| Line 6 (one-sentence status) | Update to: "Videos 1–3 recorded and shipped. Universal toolkit + zero-install Python pack live. Video 4 replanning complete (locked 2026-04-28); 5-doc planning sprint produced; V4 Python build pending; post-video Batches 4–5 still parked." |
| Line 12 (elevator pitch) | Replace "for 2,000+ iPipeline coworkers + the CFO and CEO" with "for 50–150 iPipeline coworkers near-term in Finance, Accounting, and adjacent operations. Broader rollout deferred." |
| Lines 20–22 (Audience & voice) | Rewrite primary viewers as 50–150 near-term; demote "executive viewers" to "future audience for v2"; add adoption-grade framing — coworkers will actually use the tools, not just watch. |
| Lines 175–208 ("Current Video 4 plan") | Replace with a summary of the locked V4 direction (single chaptered V4, middle-ground hero, `finance_automation_launcher.py`, 6 revised scripts, sample data design lock pending). Reference the 5 sprint docs as canonical. |
| Lines 212–220 ("Open decisions") | Replace with a short list of remaining open items (sample data design, who builds, coworker pip access, pilot user list). |

### 6.5 Memory — `C:\Users\connor.atlee\.claude\projects\c--Users-connor-atlee--claude-projects-claude-training-lab-code\memory\`

| File | Change |
|---|---|
| `MEMORY.md` line 1 | Replace "building CEO/CFO demo" with "building adoption-grade automation package for 50–150 coworkers near-term; CEO/CFO showcase deferred to v2." |
| `user_profile.md` lines 3, 9 (and any others mentioning 2000+ / CFO/CEO) | Same correction — 50–150 coworkers near-term, broader audience deferred. |
| `project_status.md` (entire V4 section) | Replace "5 decisions awaiting Connor's lock-in" with "V4 direction LOCKED 2026-04-28: single chaptered V4, middle-ground hero (Revenue Leakage Finder + ARR waterfall artifact), `finance_automation_launcher.py`, 5–7 workflow doorway, adoption-grade framing. 5-doc planning sprint in progress / complete. V4 Python build pending." |
| `reference_locations.md` line 13, 18, 55 | Update stale notes — `RecTrial\AGENTS.md` audience reference will be corrected; `VIDEO_4_CURRENT_PROPOSAL.md` is now superseded by this memo + the rest of the sprint; todo.md V4 section will be rewritten. |

### 6.6 Other files worth scanning during the cleanup pass

These weren't in scope for this memo but are likely to contain stale framing:
- `RecTrial\Brainstorm\VIDEO_4_CURRENT_PROPOSAL.md` — should be marked **SUPERSEDED** with a pointer to this memo + the sprint's other 4 docs. Don't delete (history matters).
- `RecTrial\Brainstorm\VIDEO_4_DRAFT_IDEAS.md` — initial 17-idea draft. Probably fine as a historical artifact; check that nothing in it reads as current direction.
- Training guides under `training/` and `FinalRoughGuides/` — most predate this conversation and probably reference the old "CFO/CEO demo" framing. Out of scope for V4 sprint cleanup; flag for a separate audit pass post-V4.

---

## 7. Open items (for next session, NOT blockers for this sprint)

| # | Item | Where it gets resolved |
|---|---|---|
| 1 | Sample data design for Revenue Leakage Finder (realistic iPipeline billing variance patterns) | `VIDEO_4_REVISED_PLAN.md` — Connor describes patterns in plain English; doc codifies them |
| 2 | Coworker pip access reality (changes whether stdlib-only is mandatory or just preferred) | `PYTHON_SAFETY.md` — fallback paragraph documents both states; Connor confirms after a quick coworker check |
| 3 | Who builds V4 Python — Codex, Claude Code, or both | After all 5 sprint docs land. Connor decides; Codex spec at `02_codex_build_spec.md` is buildable as-is. |
| 4 | Pilot user list (10–20 specific people) | `MINIMUM_DISTRIBUTION_PLAN.md` documents the role mix (3–5 Finance + 3–5 Accounting + 3–5 Billing/RevOps + 1–2 managers + optional IT/security observer); Connor identifies actual people separately. |

---

## 8. What happens next (production order for the rest of the sprint)

1. ✅ **This memo** — locks captured first.
2. ⏳ `RecTrial\Brainstorm\SUPPORTED_WORKFLOWS_V1.md` — the 5–7 starter workflows mapped to existing modules + adoption-grade "drop into your own file" guidance per workflow.
3. ⏳ `RecTrial\Brainstorm\VIDEO_4_REVISED_PLAN.md` — chapter outline, demo sequence, build/recording effort, **sample data design lock section** (the synthetic data spec for Revenue Leakage Finder).
4. ⏳ `RecTrial\UniversalToolkit\python\PYTHON_SAFETY.md` — 14 visible safety rules + adoption-grade "running on your real files" section + pip access fallback paragraph.
5. ⏳ `RecTrial\Brainstorm\MINIMUM_DISTRIBUTION_PLAN.md` — SharePoint package, launch message, pilot plan, support intake process.

After all 5 docs land:
- Cleanup pass on the stale-reference files in section 6 — edit once, not twice.
- Decide who builds V4 Python (Codex / Claude Code / both).
- V4 Python build begins.

---

**End of memo.**
