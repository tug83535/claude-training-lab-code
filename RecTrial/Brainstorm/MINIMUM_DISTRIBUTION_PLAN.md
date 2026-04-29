# Minimum Distribution Plan — Finance Automation Toolkit v1.0

**Owner:** Connor Atlee — Finance & Accounting, iPipeline
**Date:** 2026-04-28
**Audience:** 50–150 coworkers in Finance, Accounting, and adjacent operations (near-term)
**Delivery model:** One self-contained zip on SharePoint. Download, unzip, open Excel, click button. No install, no terminal, no IT involvement for normal use.

---

## 1. How this release works — the key principle

> Release a small, supported doorway into a large toolkit.

The full toolkit contains ~140 VBA tools and 28 Python scripts. That is not what's being released. What's being released is:

- 7 supported starter workflows (defined in `SUPPORTED_WORKFLOWS_V1.md`)
- One Excel workbook with buttons for each workflow
- Python running locally via Excel buttons — no Command Prompt, no setup
- Sample data for every workflow so coworkers can try before touching real files
- Clear docs so Connor's support load stays manageable

Everything else in the toolkit is still inside `FinanceTools.xlsm` and the Command Center — discoverable for power users — but not the opening pitch.

---

## 2. Package structure — what's inside the zip

One zip file on SharePoint. Coworkers download it, unzip to a permanent folder on their machine, and open Excel. That's the entire setup.

```
FinanceTools_v1.0.zip
└── FinanceTools_v1.0\
    ├── FinanceTools.xlsm               ← open this first after reading Start Here
    ├── 00_START_HERE.pdf               ← read this before opening anything else
    ├── Quick_Reference_Card_v1.0.pdf   ← one-page cheat sheet, print or keep open
    ├── python\
    │   ├── python-embedded\            ← bundled Python 3.11 (no install required)
    │   └── scripts\                    ← all Python automation scripts
    ├── samples\                        ← synthetic data for sample mode (all 7 workflows)
    ├── outputs\                        ← empty placeholder; all run outputs appear here
    └── docs\
        ├── PYTHON_SAFETY.md            ← full safety rules (also shown in the workbook)
        ├── Known_Limitations_v1.0.pdf
        ├── Troubleshooting_v1.0.pdf
        └── Release_Notes_v1.0.pdf
```

**Why one zip:** coworkers get one thing to download, one thing to unzip, and one workbook to open. There is no separate Python download, no separate sample files download, no separate docs download. Everything is in the right place the moment they unzip.

**Where to unzip:** anywhere except a OneDrive-synced folder or a network drive. A local folder like `C:\FinanceTools\` or `C:\Users\yourname\Documents\FinanceTools\` works reliably. The toolkit uses relative paths so it works wherever it lives — the one exception is live-synced OneDrive folders, which can cause file-locking conflicts while Python scripts are writing output.

**Note for Connor's machine:** Connor's Desktop is OneDrive-redirected to `C:\Users\connor.atlee\OneDrive - iPipeline\Desktop\`. When testing, unzip to a non-synced location. When writing instructions for coworkers, recommend `Documents\FinanceTools\` as the default unzip target.

---

## 3. SharePoint page structure

One page or folder. Not a complex site — one place, one download.

**Page name:** Finance Automation Toolkit v1.0

**Page sections:**
1. **What this is** (2–3 sentences): Excel and Python automation tools for Finance & Accounting. Start with sample files. 7 supported workflows.
2. **Watch first** (embedded or linked): Video 4 — Python Automation for Finance (when recorded)
3. **Download** (one prominent button): `FinanceTools_v1.0.zip`
4. **Getting started** (3-line bulleted steps): Download → Unzip to Documents\FinanceTools\ → Open 00_START_HERE.pdf
5. **Questions / issues**: Contact Connor Atlee

**Not on this page:** links to individual scripts, VBA module documentation, the full 140-tool list, or any "advanced" content. That stays discoverable inside the workbook.

**Version control:** When v1.1 or v2.0 ships, update the zip and the page version label. Keep the old zip available in an `Archive/` subfolder on the same SharePoint page so coworkers who are mid-use aren't broken.

---

## 4. Launch message — ready to send

**Subject:**
```
Finance Automation Toolkit v1.0 — Excel + Python automation for Finance, available now
```

**Body:**
```
Hi team,

I've put together a Finance Automation Toolkit (v1.0) with 7 practical workflows
for Excel and Python automation. It's now available on SharePoint.

[SharePoint link]

What's included:
- Excel tools: clean messy exports, compare files, consolidate sheets, find workbook
  issues, generate a workbook summary
- Python tools: Revenue Leakage Finder (finds billing gaps vs. contracts), Data
  Contract Checker (validates file structure before analysis)

How to start:
1. Download and unzip FinanceTools_v1.0.zip to your Documents folder
2. Read 00_START_HERE.pdf (2 pages — worth the 3 minutes)
3. Open FinanceTools.xlsm and click "Run Sample" on any workflow
4. Once you've run a sample, click "Run on Your File" and pick your own file

Everything runs on your local machine. No internet, no external AI, no credentials,
no changes to your input files. Outputs go to an outputs\ folder inside the toolkit.

This is v1, so please start with the supported workflows (listed in the Quick
Reference Card). Send questions or issues directly to me.

Thanks,
Connor
```

---

## 5. Self-service onboarding sequence

This is the step-by-step path every coworker should follow on their first use. The `00_START_HERE.pdf` should walk through exactly these steps.

**Step 1 — Download and unzip**
- Download `FinanceTools_v1.0.zip` from the SharePoint page
- Unzip to `C:\Users\yourname\Documents\FinanceTools\` (or similar non-OneDrive location)
- Do not open from inside the zip — always run from the unzipped folder

**Step 2 — Read Start Here first (3 minutes)**
- Open `00_START_HERE.pdf` before anything else
- It tells you: what's in the toolkit, which workflow to try first, what the Excel buttons do, how to find your output, and how to get help

**Step 3 — Enable macros when prompted**
- Open `FinanceTools.xlsm`
- When Excel shows the security bar at the top ("Macros have been disabled"), click **Enable Content**
- This is required — the Excel buttons are macros. The toolkit cannot run without them.
- If your IT settings prevent enabling macros entirely, contact Connor.

**Step 4 — Run a sample first, always**
- In the workbook, find the workflow you want to try and click **Run Sample**
- Sample mode uses pre-built synthetic data — it does not touch any real files
- Watch what happens: Python runs, output folder appears, results are ready
- Open the output folder and review what was produced

**Step 5 — Run on your own file**
- Once you've seen sample mode work, click **Run on Your File** for the same workflow
- A file browse dialog appears — navigate to your file and select it
- Do not use sensitive production files on your first try. Use a copy of a report or an export from a non-critical period.
- Output appears in the same `outputs\` folder with a new timestamp

**Step 6 — Review the output**
- Open the output folder (`FinanceTools_v1.0\outputs\YYYYMMDD_HHMMSS_toolname\`)
- Review the HTML report and/or CSV file
- Check `run_summary.txt` if anything looks unexpected — it explains what the script did

**Step 7 — Questions or issues**
- Contact Connor directly (not a shared inbox)
- Include: which workflow, what you clicked, a screenshot of the error or unexpected result, whether you were using sample or real data

---

## 6. Support intake process

Connor owns support. The release must limit avoidable load. The following process applies from day one.

**What coworkers send when reporting an issue:**
1. Tool name (which workflow / which button)
2. Screenshot of the error message or unexpected result
3. Copy of `run_summary.txt` from the output folder (if Python workflow)
4. Whether they were running sample mode or real file
5. Where they unzipped the toolkit (in case it's a path issue)

**What Connor tracks** — maintain a simple log (a spreadsheet is fine):

| Date | User/Team | Workflow | Issue type | Status | Fix needed? | Notes |
|---|---|---|---|---|---|---|

Issue types: `Bug`, `User confusion`, `Enhancement request`, `Unsupported use case`, `IT/security flag`

**Response time target:** within 3 business days for bugs; within 1 week for questions. Set this expectation in the Start Here doc so coworkers aren't waiting for a same-day response.

**Intake limit:** if more than 5 issues/week arrive during pilot, pause the broader rollout and triage before continuing. The pilot phase exists specifically to catch and absorb this load before it hits 50–150 people.

---

## 7. Pilot plan — 10–20 users before broader release

Do not send the toolkit to all 50–150 people on day one. Run a 2-week pilot with a targeted group first.

**Pilot cohort size:** 10–20 people (minimum 10 to get meaningful signal)

**Role mix:**

| Role group | Count | Purpose |
|---|---|---|
| Finance users | 3–5 | Core audience for the Excel workflows |
| Accounting users | 3–5 | Core audience, especially for data cleaning and compare tools |
| Billing / RevOps users | 3–5 | Primary audience for Revenue Leakage Finder |
| Managers (Finance or Accounting) | 1–2 | Validate that the output format works for leadership |
| IT or security observer (optional) | 0–1 | Early flag on endpoint scanner / bundled python.exe concern |

**Connor's action:** identify the specific names for each slot. The role groups are the target — filling them with real people is a real-world task that only Connor can do.

**Pilot briefing message (send to the cohort before they download):**
```
Hi [name],

I'm doing a soft launch of the Finance Automation Toolkit with a small group before
sharing it more broadly. Would you be willing to try it out and give me feedback?

All you need to do:
1. Download and unzip the package (I'll send the link separately)
2. Try at least one workflow using sample mode
3. If it goes well, try it on a real non-sensitive file
4. Send me any issues or confusion you hit — even small things

I'm specifically looking for: things that are confusing, things that break, and
any workflows that don't match how you actually work.

Takes 15–30 minutes. Much appreciated.

Connor
```

**Pilot duration:** 2 weeks from link-send to feedback collection

**Pilot success metrics** — before the broader rollout, confirm:

| Metric | Target | Status |
|---|---|---|
| People who open the package | At least 10 of the cohort | |
| People who run a sample workflow | At least 5 | |
| People who try on a real file | At least 3 | |
| Top confusing points identified | At least 3 specific items | |
| Bugs found and fixed or documented | Top 3 resolved | |
| Concrete use cases captured | At least 2 ("I used it to...") | |

If the pilot completes and all 6 metrics are hit, proceed to broader rollout. If fewer than 3 people complete a sample run, investigate before expanding — something in the onboarding is broken.

---

## 8. IT dependency — bundled python.exe (launch blocker, not code blocker)

**The issue:** `FinanceTools_v1.0.zip` contains `python.exe` inside the `python\python-embedded\` folder. Some enterprise endpoint scanners automatically flag `.exe` files distributed via SharePoint zip packages.

**What this means:** if iPipeline's endpoint scanner flags the package, coworkers who download and try to run it may be blocked or the python.exe may be quarantined silently, causing all Python buttons to fail with no useful error message.

**This is Connor's responsibility to resolve before the pilot launches** — not a code problem, not a design change. The code is correct regardless. What's needed:

- [ ] Confirm with IT whether the endpoint scanner scans SharePoint download zips for executable content
- [ ] If yes: request either a scanner exception for this specific package, or an alternative distribution channel (e.g., a specific approved shared drive, or IT-managed deployment)
- [ ] If no issue: proceed
- [ ] Document the outcome in the support log

**Fallback if the scanner is a hard blocker:** distribute `python-embedded\` and scripts separately from `FinanceTools.xlsm`, have IT stage the Python folder in a known location, and update `ThisWorkbook.Path` path logic to point there. This is less clean but workable. Flag to Connor before building this fallback since it changes the path assumptions in the VBA Shell() calls.

---

## 9. Final release gate — do not release to 50–150 until all pass

| Gate | Pass? |
|---|---|
| Video 4 is recorded and linked on the SharePoint page | |
| Revenue Leakage Finder sample mode works end-to-end (Excel button → output) | |
| Data Contract Checker sample mode works end-to-end | |
| All Excel buttons invoke bundled Python correctly (not system Python by default) | |
| Browse dialog tested and works for files outside the toolkit folder | |
| Python safety doc (`PYTHON_SAFETY.md`) is current and accurate | |
| `00_START_HERE.pdf` written, reviewed, and included in the zip | |
| `Quick_Reference_Card_v1.0.pdf` written, reviewed, and included in the zip | |
| `Known_Limitations_v1.0.pdf` exists (what does NOT work, what's not supported) | |
| IT endpoint scanner check completed — outcome documented | |
| Pilot (10–20 users) completed and all 6 success metrics met | |

All 11 must pass before broader rollout. The gate is binary — pass or not yet.

---

## 10. What v2 looks like (post-V1 release, not a V1 blocker)

With Excel buttons in V1, the remaining friction points that V2 should address:

| V2 item | Why it matters |
|---|---|
| Signed Excel add-in (`.xlam`) for company-wide deployment | For broader rollout beyond 50–150, IT may require a signed add-in rather than a loose `.xlsm` macro file |
| Live data exchange (xlwings or equivalent) | V1 uses file I/O — Python reads CSV, writes CSV, Excel reads the output. V2 could exchange data in-memory for faster, tighter integration. Currently parked due to xlwings install uncertainty, but the Shell() finding means this is now a design choice, not a hard constraint. |
| More supported workflows | As pilot feedback arrives, promote additional workflows from "advanced" to "supported starter" based on actual coworker usage patterns |
| IT-managed deployment | If the toolkit gets adopted beyond Finance & Accounting into other departments, IT-managed packaging and update distribution becomes necessary |
| Governance docs | `CONSTRAINTS.md`, `BRAND.md`, `RELEASE_READINESS_CHECKLIST.md`, `TROUBLESHOOTING.md` — already in the Codex Batch 5 backlog |

---

**End of distribution plan.**
