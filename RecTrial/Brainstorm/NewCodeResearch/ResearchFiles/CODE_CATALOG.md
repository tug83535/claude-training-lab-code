# MobileCLDCode — Full Code Catalog (Single-Document Reference)

## Meta: How another Claude chat should read this

You are a Claude instance arriving cold. A previous Claude session built
`MobileCLDCode/` inside the iPipeline APCLDmerge training lab repo. This
single document is the complete reference for that folder — every file, its
purpose, its inputs, its entry points, its gotchas, and how it integrates
with everything else.

**Rules of engagement:**
- Trust this doc. It was written by the same session that built the code.
- Before opening any code file, check its section in here first. You
  usually won't need to open the file.
- If the user asks to edit, add, or remove a file, keep this document in
  sync afterwards. It is the source of truth for what's in the folder.
- All 42 code files sit under `MobileCLDCode/` in the repo. The README.md
  at the root of that folder is the short-form coworker-facing pitch; this
  document is the long-form technical reference.

**Repo facts:**
- Repo: `tug83535/claude-training-lab-code`
- Branch MobileCLDCode lives on: `claude/mobilecldcode-business-automation-zC3Jt`
- Initial commit that created the folder: `32d3927`
- Parent project's active branch (different, do not touch): `claude/resume-ipipeline-demo-qKRHn`
- Working directory on the filesystem: `/home/user/claude-training-lab-code/`

---

## 1. Parent project context you must understand

**Company.** iPipeline — ~2,000-person software-for-insurance business in
the US. Typical enterprise SaaS profile: multi-tenant app, customer success
team, sales ops, a finance org that closes books monthly, a platform eng
team running AWS, and a compliance function that handles SOX + SOC 2.

**User.** Connor Atlee, not a developer. Works in Finance & Accounting
and writes training material. Reads VBA/SQL/Python at a working level but
relies on Claude to author new code. Non-technical audience for all
training material. The voice of every piece of output is plain-English.

**Parent project (APCLDmerge).** A P&L Excel demo workbook with 39 VBA
modules, a SQL data pipeline (SQLite 3), 14 Python scripts, training
guides, and a 4-video walkthrough. Target audience: the entire iPipeline
workforce including CFO/CEO. Everything is held to a "world-class"
standard — this is what every file in this repo aspires to.

**Current parent-project work (as of 2026-04-22).** Recording 4 demo videos
using an Excel VBA "Director" macro. Videos 1-2 done; 3 mid-debug after
Path A silent-wrapper refactor; 4 ready to record manually. MobileCLDCode
is a *separate track* and has nothing to do with video recording.

**iPipeline brand.**
- Primary blue `#0B4779` (aka `RGB(11, 71, 121)`)
- Secondary Navy `#112E51`, Innovation Blue `#4B9BCB`
- Accents Lime `#BFF18C`, Aqua `#2BCCD3`
- Neutrals Arctic White `#F9F9F9`, Charcoal `#161616`
- Fonts Arial family only
- Note: the parent project's legacy VBA modConfig uses slightly different
  color constants. Don't touch those. Any NEW styled output in
  MobileCLDCode uses the official brand hex codes above.

---

## 2. What MobileCLDCode is (and isn't)

**Is.** A sibling folder to the main APCLDmerge demo. Its purpose is to
answer "what else can code do for a large software business?" It
demonstrates automation beyond a single workbook — reaching into Outlook,
Teams, JIRA, SharePoint, SQL Server, AWS, PDFs, log files, Active
Directory, PowerShell, Office Scripts, and Power Automate.

**Isn't.** Production code. It's reference / training / showcase code.
It's also not part of the APCLDmerge demo workbook. Nothing in the main
demo depends on anything here.

**The hard constraint the previous session held to.** Every file here
does something modern OneDrive for Business / M365 Excel **cannot do
natively**. No trivial "how to sort a table". Every example requires code.

---

## 3. Design conventions used in every file (preserve when editing)

1. **Header block.** Every file opens with three labeled sections:
   `PURPOSE`, `WHY THIS IS NOT NATIVE`, `USE CASE`.
2. **No hardcoded secrets.** VBA reads named ranges. Python reads env
   vars. PowerShell takes parameters. If you see a hardcoded token in a
   refactor, flag it.
3. **Excel output (Python) uses openpyxl.** Default extension `.xlsx`.
   Board-ready. Styled headers: iPipeline blue band, white bold text.
4. **Standalone.** Each file pulls only from stdlib + a pinned package
   list (`02_Python/requirements.txt`) or PSGallery modules (PowerShell).
5. **Brand colors by RGB triplet.** `RGB(11, 71, 121)` or `#0B4779`.
6. **Real APIs only.** JIRA Cloud REST v3, SharePoint REST + Graph,
   Teams Incoming Webhooks, Slack Webhook + Bot, AWS CUR, pg_stat_statements,
   SQL Server DMVs. Endpoints and shapes match current public docs.
7. **Header comments explain why, not what.** The reader is non-technical
   Finance or Ops, not a developer.

---

## 4. File tree

```
MobileCLDCode/
├── README.md                   ← coworker-facing overview
├── CODE_CATALOG.md             ← THIS FILE (deep reference, single doc)
│
├── 01_VBA/                     (11 modules)
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
├── 02_Python/                  (13 scripts + requirements)
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
├── 03_SQL/                     (6 scripts)
│   ├── saas_metrics_suite.sql
│   ├── customer_360_view.sql
│   ├── sales_pipeline_velocity.sql
│   ├── data_quality_audit.sql
│   ├── revenue_recognition_schedule.sql
│   └── slow_query_tuner.sql
│
├── 04_PowerShell/              (4 scripts)
│   ├── AD_Inactive_User_Audit.ps1
│   ├── SharePoint_Site_Storage_Audit.ps1
│   ├── SSL_Cert_Expiry_Monitor.ps1
│   └── New_Hire_Account_Provisioner.ps1
│
├── 05_OfficeScripts/           (3 scripts)
│   ├── TeamsWebhookOnThreshold.ts
│   ├── DailyMetricsExport.ts
│   └── BulkFormatDataImport.ts
│
├── 06_PowerAutomate/           (2 flow JSON templates)
│   ├── PurchaseApproval_Flow.json
│   └── NewHire_Onboarding_Flow.json
│
└── 07_PowerQuery/              (2 M files)
    ├── REST_API_Pagination.pq
    └── Multi_Folder_Merge.pq
```

Total: 42 code files + README + this catalog.

<!-- APPEND MARK 1 -->
