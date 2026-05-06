# Report 4 — Release, Safety, and Distribution Checklist
## iPipeline Finance Automation Toolkit v1.0

**Purpose:** This is the extra file I would add because the project’s biggest gap is not code. It is release readiness.

Use this checklist before sharing the toolkit with 50–150 coworkers.

---

## 1. Release principle

Do not release “140 tools.”

Release a small, supported starter package.

Recommended public framing:

> Finance Automation Toolkit v1.0 helps coworkers safely try a small set of Excel and Python automation workflows using sample files first.

---

## 2. Supported v1 workflows

The v1 release should officially support only these workflows:

| # | Workflow | Primary tool/script | Support status |
|---:|---|---|---|
| 1 | Clean a messy Excel export | Data Sanitizer / Data Cleaning tools | Supported starter workflow |
| 2 | Compare two files | Sheet Compare / Quick Row Compare | Supported starter workflow |
| 3 | Consolidate sheets/files | Consolidate tools / multi-file consolidator | Supported starter workflow |
| 4 | Find workbook issues | Audit tools / external links / errors | Supported starter workflow |
| 5 | Generate workbook summary | Exec Brief / profile workbook | Supported starter workflow |
| 6 | Find possible revenue leakage | Revenue Leakage Finder | Supported Python sample workflow |
| 7 | Check file structure | Data Contract Checker | Supported Python sample workflow |

Everything else:
- included for exploration;
- not the first recommended path;
- not guaranteed as a beginner workflow.

---

## 3. Package structure

Recommended SharePoint package:

```text
Finance Automation Toolkit v1.0
├── 00_START_HERE.pdf
├── Finance_Automation_Toolkit_v1.0.xlsm
├── Python_Finance_Starter_Pack_v1.0.zip
├── Sample_Files_v1.0.zip
├── Quick_Reference_Card_v1.0.pdf
├── Known_Limitations_v1.0.pdf
├── Troubleshooting_v1.0.pdf
├── Release_Notes_v1.0.pdf
└── Optional_Advanced_Tools/
```

---

## 4. Python safety checklist

Before release, confirm every Video 4 Python script satisfies this:

| Requirement | Pass? |
|---|---|
| No internet/API calls |  |
| No external AI calls |  |
| No credentials, tokens, or secrets |  |
| No database connections |  |
| No source-file overwrites |  |
| No file deletion |  |
| Input files are read-only |  |
| Outputs go to `/outputs/` |  |
| Each run uses timestamped output folder |  |
| Each run writes `run_log.json` |  |
| Each run writes `run_summary.txt` |  |
| Sample mode exists |  |
| Clear user-facing error messages exist |  |
| Known limitations are documented |  |
| Smoke test passes |  |

---

## 5. Excel/VBA safety checklist

Before release, confirm:

| Requirement | Pass? |
|---|---|
| Destructive tools create backup or preview first |  |
| Beginner tools are clearly labeled |  |
| Advanced/risky tools are not pushed as first-run tools |  |
| Command Center duplicate labels are cleaned up |  |
| Current version number is visible |  |
| Sample workbook works on a clean machine |  |
| Macros launch from SharePoint/download location as expected |  |
| User can recover if something goes wrong |  |
| Run receipt exists for material tools where feasible |  |
| Known limitations are documented |  |

---

## 6. Documentation checklist

Create or update these before release:

| Document | Purpose | Required? |
|---|---|---|
| `00_START_HERE.pdf` | First file users open | Yes |
| `Quick_Reference_Card.pdf` | One-page cheat sheet | Yes |
| `Known_Limitations.pdf` | Protects trust | Yes |
| `Troubleshooting.pdf` | Reduces support burden | Yes |
| `Release_Notes.pdf` | Version clarity | Yes |
| `PYTHON_SAFETY.md` | Explains Python safety | Yes |
| `SUPPORTED_WORKFLOWS_V1.md` | Narrows support scope | Yes |
| `README_VIDEO4_PYTHON.md` | Python starter instructions | Yes |

---

## 7. Launch message draft

Subject:
```text
Finance Automation Toolkit v1.0 — safe sample workflows for Excel + Python
```

Body:
```text
Hi all,

I put together a Finance Automation Toolkit v1.0 for people who want to try the workflows shown in the demo series.

Start here:
[SharePoint link]

Recommended first workflows:
1. Clean a messy Excel export
2. Compare two files
3. Consolidate sheets/files
4. Find workbook issues
5. Generate a workbook summary
6. Run the sample Revenue Leakage Finder
7. Run the sample Data Contract Checker

Please start with the sample files first.

The Python scripts are local-only: they do not call the internet, do not use external AI, do not ask for credentials, and do not modify source input files. Outputs are written to a separate outputs folder.

This is a v1 toolkit, so use the supported starter workflows first and send issues/questions to me.

Thanks,
Connor
```

---

## 8. Support model

Because Connor owns support, the release must limit avoidable support load.

Minimum support process:

1. Users send:
   - tool name;
   - screenshot;
   - error message;
   - run log if Python;
   - whether they used sample or real file.

2. Connor tracks:
   - bug;
   - user confusion;
   - enhancement request;
   - unsupported use case.

3. Use a simple tracker:

| Date | User/team | Tool | Issue type | Status | Fix needed? | Notes |
|---|---|---|---|---|---|---|

---

## 9. Pilot plan

Before sharing with all 50–150 people, test with 10–20.

Pilot mix:
- 3–5 Finance users;
- 3–5 Accounting users;
- 3–5 Billing/RevOps users;
- 1–2 managers;
- optional IT/security observer.

Pilot success metrics:
- at least 10 people open the package;
- at least 5 run a sample workflow;
- at least 3 try a real non-sensitive file;
- top 3 confusing points are identified;
- top 3 bugs are fixed or documented;
- at least 2 concrete use cases are captured.

---

## 10. Final release gate

Do not release to the broader 50–150 group until:

| Gate | Pass? |
|---|---|
| Video 4 direction is locked |  |
| Revenue Leakage Finder sample works |  |
| Launcher sample mode works |  |
| Python safety doc exists |  |
| Excel starter workflows are clear |  |
| Command Center is not visibly messy |  |
| SharePoint folder/page exists |  |
| Start Here doc exists |  |
| Known limitations doc exists |  |
| Support intake process exists |  |
| v1 package has versioned filenames |  |

---

## 11. Recommendation

The release should feel smaller than the build.

That is not a weakness.

It is how you make the project usable.

You built a large toolkit. Now release a small, safe doorway into it.
