# MobileCLDCode — Full Code Catalog (For Another Claude Chat)

## Meta: How to use this document

You are a Claude instance arriving cold. A previous Claude session built a
folder called `MobileCLDCode/` inside the iPipeline "APCLDmerge" training lab
repo. This document, plus the seven per-category docs in `docs/`, is the
canonical reference for what's in that folder.

**Read order:**
1. This file (master context + index).
2. The per-category doc(s) relevant to whatever the user is asking about —
   they live at `MobileCLDCode/docs/0X_*_Catalog.md`.
3. Only then open the actual code files. Each per-category doc summarizes
   every file in enough depth that you usually won't need to open them.

**What you should NOT assume:**
- That these files are already deployed. They are reference code.
- That the main APCLDmerge project depends on any of this. It doesn't.
  MobileCLDCode is a separate "what's possible" showcase, not part of the
  production demo workbook.

---

## 1. Parent project context (APCLDmerge)

- **Company:** iPipeline, a ~2,000 person software-for-insurance business.
- **User:** Connor (non-developer). Works in Finance & Accounting and writes
  training material. Uses VBA, SQL, and Python at a practical "I can read it"
  level but relies on Claude to write new code.
- **Parent project:** "APCLDmerge" — a world-class P&L Excel demo workbook
  with 39 VBA modules, SQL pipeline, 14 Python scripts, training guides,
  and video walkthroughs. Target audience: the entire iPipeline workforce
  including the CFO/CEO.
- **Current work in the parent project (as of 2026-04-16):** recording
  4 demo videos using an Excel "Director" macro that automates the recording.
  That is a different track entirely from MobileCLDCode.
- **iPipeline brand:** Primary `#0B4779` blue, Navy `#112E51`, Arial family
  fonts only. The MobileCLDCode files respect this where they produce
  styled output.

---

## 2. What MobileCLDCode is

A sibling folder to the main demo. Its purpose is to answer the question
"what *else* can code do for a large software business?" It shows
coworkers that automation goes beyond one workbook — it reaches into
Outlook, Teams, JIRA, SharePoint, SQL Server, AWS, PDFs, log files,
Active Directory, PowerShell, Office Scripts, and Power Automate.

**The hard constraint the previous Claude session held to:**
Every file here does something that modern OneDrive for Business or
Microsoft 365 Excel *cannot do natively*. No "here's how to sort a
table". Every example is only solvable with code.

---

## 3. The 7 categories (42 files + README + 8 doc files)

| # | Folder | Files | Theme |
|---|---|---|---|
| 01 | `01_VBA/` | 11 | Excel reaching out to external systems (Outlook, Teams, JIRA, SharePoint REST, SQL Server, Slack, file system) |
| 02 | `02_Python/` | 13 + `requirements.txt` | Heavy-duty analytics, NLP, API integration, PDF/log parsing, ML scoring |
| 03 | `03_SQL/` | 6 | Warehouse-side analytics that Excel cannot do at scale |
| 04 | `04_PowerShell/` | 4 | IT & infrastructure automation (AD, SharePoint admin, certs, provisioning) |
| 05 | `05_OfficeScripts/` | 3 | Excel-on-the-web + Power Automate triggers |
| 06 | `06_PowerAutomate/` | 2 | Cross-service orchestration flow templates |
| 07 | `07_PowerQuery/` | 2 | Paginated REST + multi-folder ingestion inside Excel/Power BI |

---

## 4. File tree

```
MobileCLDCode/
├── README.md                  ← short-form user-facing overview
├── CODE_CATALOG.md            ← this file (deep index for another Claude)
├── docs/
│   ├── 01_VBA_Catalog.md
│   ├── 02_Python_Catalog.md
│   ├── 03_SQL_Catalog.md
│   ├── 04_PowerShell_Catalog.md
│   ├── 05_OfficeScripts_Catalog.md
│   ├── 06_PowerAutomate_Catalog.md
│   └── 07_PowerQuery_Catalog.md
│
├── 01_VBA/
│   ├── modMailMerge_WithAttachments.bas
│   ├── modTeamsNotifier.bas
│   ├── modJiraBridge.bas
│   ├── modInvoicePDFGenerator.bas
│   ├── modSharePointSync.bas
│   ├── modSQLServerRunner.bas
│   ├── modRenewalAlertEngine.bas
│   ├── modMultiWorkbookDiff.bas
│   ├── modFolderOrganizer.bas
│   ├── modCalendarAppointmentBuilder.bas
│   └── modSlackNotifier.bas
│
├── 02_Python/
│   ├── requirements.txt
│   ├── saas_arr_waterfall.py
│   ├── customer_churn_risk_scorer.py
│   ├── license_utilization_analyzer.py
│   ├── aws_cost_optimizer.py
│   ├── contract_pdf_extractor.py
│   ├── support_ticket_triage.py
│   ├── api_slo_tracker.py
│   ├── revenue_recognition_engine.py
│   ├── jira_weekly_digest.py
│   ├── cohort_retention_analyzer.py
│   ├── email_to_structured_data.py
│   ├── sox_evidence_collector.py
│   └── git_developer_metrics.py
│
├── 03_SQL/
│   ├── saas_metrics_suite.sql
│   ├── customer_360_view.sql
│   ├── sales_pipeline_velocity.sql
│   ├── data_quality_audit.sql
│   ├── revenue_recognition_schedule.sql
│   └── slow_query_tuner.sql
│
├── 04_PowerShell/
│   ├── AD_Inactive_User_Audit.ps1
│   ├── SharePoint_Site_Storage_Audit.ps1
│   ├── SSL_Cert_Expiry_Monitor.ps1
│   └── New_Hire_Account_Provisioner.ps1
│
├── 05_OfficeScripts/
│   ├── TeamsWebhookOnThreshold.ts
│   ├── DailyMetricsExport.ts
│   └── BulkFormatDataImport.ts
│
├── 06_PowerAutomate/
│   ├── PurchaseApproval_Flow.json
│   └── NewHire_Onboarding_Flow.json
│
└── 07_PowerQuery/
    ├── REST_API_Pagination.pq
    └── Multi_Folder_Merge.pq
```

---

## 5. Design conventions used across every file

Any Claude working on this folder should preserve these:

1. **Header block.** Every file begins with three labeled sections:
   - `PURPOSE` — one paragraph on what the code does.
   - `WHY THIS IS NOT NATIVE` — why Excel / OneDrive / Office can't already
     do this. This is the sales pitch for the training audience.
   - `USE CASE` — a specific, concrete scenario at a large software company.
2. **Secrets by named range (VBA) / env var (Python, PowerShell).**
   No tokens hardcoded. VBA reads named ranges like `TeamsWebhookUrl`,
   `JiraApiToken`, `SQLConnString`. Python reads `os.environ[...]`.
3. **Excel output uses openpyxl.** Python scripts emit `.xlsx` as the default
   output format (board-ready). Headers are bold; iPipeline blue `#0B4779`
   where a colored header band is applied.
4. **No external-repo dependencies.** Every file is standalone. If it needs
   something, it imports from the standard library or a pinned package
   listed in `requirements.txt` (Python) or installable from PSGallery
   (PowerShell).
5. **Brand colors are referenced by RGB triplets.** Primary blue is
   `RGB(11, 71, 121)` or `#0B4779`. The main project's VBA uses slightly
   different legacy color constants — do not copy those here.
6. **No fabricated APIs.** JIRA Cloud REST v3, SharePoint REST + Graph,
   Teams Incoming Webhooks, Slack Webhook + Bot, AWS CUR, pg_stat_statements,
   and SQL Server DMVs are all real. Endpoints, auth headers, and shapes
   match their current public docs.

---

## 6. What you should do when the user asks for changes

- **If they want to add a new file:** Match the header block + conventions.
  Put it in the category folder that fits, update both the README.md and
  the relevant `docs/0X_*_Catalog.md`, and verify the main `CODE_CATALOG.md`
  still reflects the file tree in Section 4.
- **If they want to remove a file:** Remove from disk, update both docs,
  and update this catalog.
- **If they want to refactor a file:** Keep the header block. Keep the
  no-hardcoded-secrets rule. Preserve the use-case story — that's what the
  training audience will read.
- **If they want to reuse this pattern at a different company:** Search for
  "iPipeline" and brand hex codes, swap them, and update the file tree if
  you add more categories.

---

## 7. Git state

- **Branch MobileCLDCode lives on:** `claude/mobilecldcode-business-automation-zC3Jt`
- **Initial commit that added the folder:** `32d3927` — "Add MobileCLDCode:
  automation ideas for a large software business"
- **Repo URL:** `tug83535/claude-training-lab-code`
- **Parent project's active branch (different!):** `claude/resume-ipipeline-demo-qKRHn`

---

## 8. Where to go next

Open whichever per-category doc matches the task:

- **User is editing a VBA module →** `docs/01_VBA_Catalog.md`
- **User is editing a Python script →** `docs/02_Python_Catalog.md`
- **User is editing a SQL query →** `docs/03_SQL_Catalog.md`
- **User is editing a PowerShell script →** `docs/04_PowerShell_Catalog.md`
- **User is editing an Office Script →** `docs/05_OfficeScripts_Catalog.md`
- **User is editing a Power Automate flow →** `docs/06_PowerAutomate_Catalog.md`
- **User is editing a Power Query →** `docs/07_PowerQuery_Catalog.md`

Each per-category doc gives you, for every file:
- File path + size scale
- Purpose, why-not-native, use case
- Required sheet layouts / CSV schemas / env vars / named ranges
- Entry points (public subs / `main()` / CREATE VIEW statements)
- Key internal helpers
- Gotchas / known limitations
- Dependencies
- How it integrates with other files in the folder
