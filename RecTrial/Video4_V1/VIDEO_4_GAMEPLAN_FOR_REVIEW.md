# Video 4 — Full Gameplan for AI Review
## Python Automation for Finance — iPipeline

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Date:** 2026-04-28
**Purpose:** Self-contained briefing for an external AI system reviewing the Video 4 plan.
Read this document first. Source planning docs are listed in Section 11 for deeper dives.

---

## 1. Who is Connor and what is this project?

Connor works in Finance & Accounting at iPipeline. He is **not a developer**. He is building
a world-class automation package — Excel VBA, Python, and SQL tools — for Finance &
Accounting staff at iPipeline and recording a 4-video demo series to accompany it.

**Videos 1–3 are complete and shipped.** Video 4 is the final video.

- **Video 1** — "What's Possible" — overview of the full toolkit
- **Video 2** — "Full Demo Walkthrough" — end-to-end demo of the Excel/VBA tools
- **Video 3** — "Universal Tools" — VBA tools that work on any Excel file
- **Video 4** — "Python Automation for Finance" — Python tools for Finance analysts

Connor records narration separately using ElevenLabs (AI voice), then records a screen
capture to match. He edits them together in video editing software.

---

## 2. Near-term audience

**50–150 coworkers** in Finance, Accounting, and adjacent operations at iPipeline.
Non-developers. Excel-literate. Zero Python exposure. They will receive a SharePoint zip
and are expected to actually use the tools on their own work files — not just watch the video.

Broader rollout and any CFO/CEO showcase are deferred to a later date, not cancelled.

---

## 3. The 5 locked decisions (as of 2026-04-28)

### Lock 1 — Single chaptered video, 9–12 minutes
No split into 4a and 4b. One continuous video with 8 chapters. Chapter markers let viewers
jump to any tool without watching from the start.

### Lock 2 — Revenue Leakage Finder is the hero
The central story is: "Python finds hidden billing problems your Excel can't see." The Revenue
Leakage Finder runs against a synthetic billing + contracts dataset and surfaces:
- Customers billed with no matching contract on file
- Contracts expired but still generating invoices
- Billing amounts outside contract terms by >10%

The ARR waterfall chart is a closing visual artifact at the end of Chapter 3 — not a
separate hero in its own right.

### Lock 3 — The deliverable is `finance_automation_launcher.py` (not "finance_copilot.py")
A numbered CLI menu that runs each Python tool. Name was changed because "Copilot"
implies AI and these tools have none. Honest naming matters for coworker trust and IT review.

### Lock 4 — Lead with 5–7 supported starter workflows
The toolkit has ~140 VBA tools and 6+ Python scripts. Coworkers are only shown 5–7
specifically documented workflows as the recommended entry points. Everything else is
discoverable but not the doorway. This protects Connor's support load.

**The 7 supported starter workflows:**
1. Clean a messy Excel export (VBA: modUTL_DataSanitizer)
2. Compare two files (Python: compare_workbooks.py)
3. Consolidate sheets/files (VBA: modUTL_Consolidate)
4. Find workbook issues and external links (VBA: modUTL_Audit + workbook_dependency_scanner.py)
5. Generate a workbook / executive summary (Python: build_exec_summary.py)
6. Run Revenue Leakage Finder on sample data, then your own file
7. Run Data Contract Checker on sample data, then your own file

### Lock 5 — Adoption-grade framing
Coworkers are expected to actually use these tools on their own files. This is not a
watch-only demo. That raises the bar on documentation (self-service), distribution
(download-and-use), and support intake (Connor owns it personally).

---

## 4. What's built and what remains

### Already built and tested (as of 2026-04-28)

**Python scripts — ALL built, smoke test 5/5 PASS:**
All 6 V4 scripts are in `RecTrial\UniversalToolkit\python\ZeroInstall\`.
They use only Python standard library — no pip install required for coworkers.

| Script | What it does |
|---|---|
| `revenue_leakage_finder.py` | Cross-checks billing records vs contracts; finds 5 exception classes; outputs ranked CSV + HTML report + ARR waterfall |
| `data_contract_checker.py` | Validates a CSV file against a schema contract before analysis starts; outputs PASS/FAIL with plain-English error list |
| `exception_triage_engine.py` | Takes the leakage finder output; scores each exception by 4 factors (impact 45%, confidence 30%, recency 15%, repeat 10%); outputs top-10 action list CSV |
| `control_evidence_pack.py` | Scans any output folder; generates SHA-256 hashes for every file; produces a tamper-evident manifest + HTML evidence summary for audit/control use |
| `workbook_dependency_scanner.py` | Scans an Excel file for cross-sheet references, external links, named ranges; outputs a dependency map CSV; uses stdlib zipfile+xml, no openpyxl needed |
| `finance_automation_launcher.py` | The numbered CLI menu: options 1–5 = run each tool in sample mode, option 6 = show safety rules, option 7 = open outputs folder in Explorer, option 8 = exit |

**Common utilities** (shared by all scripts):
- `common/safe_io.py` — read-only file access, timestamped output folder creation
- `common/logging_utils.py` — plain-English run log per execution
- `common/report_utils.py` — HTML report generation (stdlib only)
- `common/sample_data.py` — generates synthetic test data (123 contracts, 336 billing rows, 6 embedded exception classes)

**Sample data** (in `ZeroInstall\samples\`):
- `contracts_sample.csv` — 123 synthetic iPipeline-style contracts
- `billing_sample.csv` — 336 billing rows with intentional issues embedded

**Smoke test:** `smoke_test_video4_python.py` — 5/5 PASS

**Planning docs** (all complete, all in `RecTrial\Brainstorm\`):
- `VIDEO_4_REVIEW_DECISION_MEMO.md` — 5 locks + what got cut + stale-reference table
- `SUPPORTED_WORKFLOWS_V1.md` — 7 starter workflows mapped to existing modules + adoption guidance
- `VIDEO_4_REVISED_PLAN.md` — 8-chapter outline + sample data design lock + build effort estimate
- `PYTHON_SAFETY.md` — 14 safety rules in plain English for non-developer coworkers
- `MINIMUM_DISTRIBUTION_PLAN.md` — SharePoint zip structure, pilot plan, support intake, 11-checkpoint release gate

**Narration script** (complete, v1.1):
- `RecTrial\Video4_V1\VIDEO_4_NARRATION_SCRIPT_v1.md`
- 9 ElevenLabs clips: V4_C01 through V4_C08, with Chapter 3 split into V4_C03a + V4_C03b
- ~1,345 words total, ~10:22 at 130 wpm — solidly in the 9–12 min target range

**Shot list** (complete, v1):
- `RecTrial\Video4_V1\VIDEO_4_SHOT_LIST_v1.md`
- Per-clip screen action guide, before-recording checklist, post-recording checklist

**VBA launcher code** (drafted, not yet imported into Excel):
- `RecTrial\UniversalToolkit\python\ZeroInstall\modFinanceToolsLauncher.bas`
- VBA Shell() sub that launches `finance_automation_launcher.py` via bundled Python
- Awaiting Connor's review before import into FinanceTools.xlsm

### Still remaining before recording

1. Connor reads and adjusts narration script wording to match natural speaking style
2. Connor records ElevenLabs audio (9 clips)
3. Connor reviews `modFinanceToolsLauncher.bas` VBA code and provides feedback
4. Build `FinanceTools.xlsm` — Excel workbook with the Finance Tools button wired to LaunchFinanceTools sub
5. Test zero-install path: bundled Python 3.11 embeddable + scripts running without system Python
6. Assemble SharePoint zip package (FinanceTools.xlsm + python-embedded/ + scripts/ + samples/ + outputs/ + docs/)
7. Record and edit Video 4

---

## 5. The delivery model

**How coworkers run the tools:**

```
SharePoint zip download
  └── Unzip to a folder on their machine
        ├── FinanceTools.xlsm          ← open this
        ├── python\
        │     └── python-embedded\
        │           └── python.exe     ← bundled Python 3.11 (zero install)
        ├── scripts\
        │     └── finance_automation_launcher.py + the 5 tool scripts
        ├── samples\
        │     └── contracts_sample.csv + billing_sample.csv
        └── outputs\                   ← created automatically per run
```

**What coworkers do:**
1. Open FinanceTools.xlsm
2. Click the "Finance Tools" button (one button, no per-tool buttons in V1)
3. A CLI window opens with the numbered menu
4. Type a number, press Enter — tool runs in sample mode
5. Results appear in the outputs/ folder (HTML report + CSV)

There is **no Python installation required** for coworkers. Python 3.11 embeddable ships
in the zip alongside the scripts. The VBA button builds the path to python.exe dynamically
using `ThisWorkbook.Path` so it works regardless of where the user unzips the folder.

**VBA pattern used:**
```vba
pyExe    = ThisWorkbook.Path & "\python\python-embedded\python.exe"
pyScript = ThisWorkbook.Path & "\scripts\finance_automation_launcher.py"
Shell "cmd.exe /k " & Chr(34) & pyExe & Chr(34) & " " & Chr(34) & pyScript & Chr(34), vbNormalFocus
```

---

## 6. The 8-chapter structure with timing

| Chapter | Title | Target time | What's on screen |
|---|---|---|---|
| C01 | Why Python after Excel and VBA? | 45 sec | Static slide — two columns: "Excel/VBA is for..." vs "Python adds..." |
| C02 | Safety first | 60 sec | PYTHON_SAFETY.md in Notepad (scroll) → outputs/ folder in Explorer |
| C03a | Revenue Leakage Finder — setup | 55 sec | Excel button → CLI menu → option 1 → processing begins |
| C03b | Revenue Leakage Finder — results | 1 min 53 sec | HTML report in browser → ARR waterfall → exceptions_ranked.csv in Excel |
| C04 | Data Contract Checker | 1 min 26 sec | CLI menu → FAIL output → fix in Notepad → PASS output |
| C05 | Exception Triage Engine | 1 min 26 sec | CLI menu → scored terminal output → top_10_action_list.csv in Excel |
| C06 | Control Evidence Pack | 1 min 26 sec | CLI menu → file list + SHA-256 hashes → evidence_summary.html in browser |
| C07 | Finance Automation Launcher | 60 sec | Excel button → full menu → option 7 (opens Explorer) → option 8 (exit) |
| C08 | How to start | 32 sec | Static text card with 4 rules |
| **Total** | | **~10 min 22 sec** | |

**Note on Chapter 3 split:** C03a is stable (setup and context — rarely needs re-recording).
C03b contains specific numbers (123 contracts, 336 rows, 38 exceptions) that only need
re-recording if the sample data changes. The split lets Connor re-record just the affected half.

---

## 7. Open flags for review

### Flag 1 — stdlib waterfall vs matplotlib (LOW — does not block recording)
The ARR waterfall in the HTML report is a pure CSS bar chart (stdlib only). If Connor later
confirms coworkers have pip access, this could be upgraded to a proper matplotlib figure.
The narration line ("This is the ARR waterfall...") does not need to change — only the visual.
This flag is marked in V4_C03b of the narration script.

### Flag 2 — coworker pip access (TBD — Connor's real-world task)
All 6 scripts are stdlib-only by design. If Connor confirms pip access, pandas/openpyxl
could be added in a v1.1 update for cleaner CSV handling and richer Excel output.
Recording proceeds with stdlib-only. This is a post-V4 decision.

### Flag 3 — zero-install path not yet live-tested
The bundled Python 3.11 embeddable approach has not yet been tested on a real iPipeline
coworker machine. This needs to happen before the SharePoint package goes out, not before
recording. The scripts work correctly under system Python (smoke test 5/5 PASS).

### Flag 4 — FinanceTools.xlsm not yet built
The Excel workbook with the Finance Tools button does not exist yet. The VBA code
(`modFinanceToolsLauncher.bas`) is drafted but not imported into a workbook. This workbook
needs to be built, tested, and included in the SharePoint zip before distribution.

---

## 8. What this video is NOT

- Not a developer tutorial — no Python syntax shown, no `pip install`, no IDE
- Not a CFO/CEO pitch — plain Finance & Accounting coworker audience
- Not a 140-tool showcase — only the 5–7 starter workflows are the doorway
- Not a split 4a+4b structure — one video only
- Not using "Finance Copilot" naming — no AI implied
- Not using matplotlib (yet) — all charts are stdlib HTML/CSS
- xlwings is parked for v2 (locked-down corporate laptops may block xlwings install)
- SOX Evidence Collector is out of V1 scope (Control Evidence Pack covers the ground)
- Video 5 ("Getting Started" — how to download and set up) is planned but separate from V4

---

## 9. Safety rules in effect during recording

14 rules in `RecTrial\UniversalToolkit\python\PYTHON_SAFETY.md`. Key ones:
1. Scripts run entirely on the local machine — no internet, no external calls
2. Input files are **never modified** — all scripts open files read-only
3. Every run creates a new timestamped folder in outputs/ — nothing is ever overwritten
4. All output goes to outputs/YYYYMMDD_HHMMSS_toolname/ — input files never touched
5. When something goes wrong, the user sees plain English, not a Python stack trace

During the video demo: Connor runs all scripts using sample data only. No real customer
names, no real billing data, no real contract data appears on screen.

---

## 10. Pilot plan (post-recording)

Before broad rollout, Connor will pilot with 10–20 users:
- 3–5 Finance team members
- 3–5 Accounting team members
- 3–5 Billing/RevOps team members
- 1–2 managers
- Optional: 1–2 IT/security observers

Pilot goal: confirm zero-install path works on real iPipeline laptops, collect real bug reports,
confirm the 5–7 starter workflows are usable without Connor's help.

Full pilot plan and 11-checkpoint release gate in `RecTrial\Brainstorm\MINIMUM_DISTRIBUTION_PLAN.md`.

---

## 11. Source planning docs (for deeper review)

All in `RecTrial\Brainstorm\` unless noted:

| Doc | What it covers |
|---|---|
| `VIDEO_4_REVIEW_DECISION_MEMO.md` | All 5 locked decisions with full reasoning + what got cut |
| `VIDEO_4_REVISED_PLAN.md` | 8-chapter detail, sample data design lock, build effort estimate |
| `SUPPORTED_WORKFLOWS_V1.md` | 7 starter workflows with module paths, adoption guidance, prerequisites |
| `MINIMUM_DISTRIBUTION_PLAN.md` | SharePoint zip structure, launch message, pilot plan, 11-checkpoint release gate |
| `PYTHON_SAFETY.md` (in `python\`) | 14 safety rules + adoption-grade "running on your real files" section |
| `RecTrial\Video4_V1\VIDEO_4_NARRATION_SCRIPT_v1.md` | Full word-for-word narration for ElevenLabs, 9 clips |
| `RecTrial\Video4_V1\VIDEO_4_SHOT_LIST_v1.md` | Per-clip screen recording guide, before/after checklists |
| `RecTrial\UniversalToolkit\python\ZeroInstall\README_VIDEO4_PYTHON.md` | Python scripts README for coworkers |

---

## 12. Questions this AI reviewer might want to address

Suggested review angles — adjust as needed for the specific review task:

1. **Narrative coherence** — Does the 8-chapter arc tell one clear story from problem to solution to "how do I start"? Are there gaps or jumps?
2. **Audience fit** — Is the language and pacing right for a non-developer Finance coworker? Is anything too technical without enough context?
3. **Demo believability** — Does the Revenue Leakage Finder demo feel realistic for an iPipeline Finance analyst? Does the sample data size (123 contracts, 336 rows) feel real or toy?
4. **Safety/trust** — Is the safety chapter (Chapter 2) strong enough to overcome a non-developer's hesitation about running Python scripts?
5. **Adoption pathway** — Is there a clear, low-friction "how do I try this on my own file?" moment? Does the video land on a call to action?
6. **Narration quality** — Does the narration script read naturally at a conversational pace? Any lines that feel stiff or overly technical?
7. **Shot list completeness** — Are the on-screen actions clear enough that someone recording for the first time could follow them without ambiguity?
8. **Delivery model risks** — Any risks with the bundled Python / VBA Shell() / zero-install approach that should be flagged before distribution?
9. **Open items** — Are the four flagged open items (matplotlib, pip access, live test, xlsm build) appropriately scoped as post-recording vs pre-recording?
10. **What's missing** — Is there anything a Finance & Accounting coworker audience needs that isn't covered in any of the 8 chapters?

---

*End of gameplan document. Version 1.0 — 2026-04-28.*
*For questions contact Connor Atlee — Finance & Accounting, iPipeline.*
