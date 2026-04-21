# MobileCLDCode — Automation Ideas for a Large Software Business

**What this folder is**
A companion to the main APCLDmerge P&L demo. The main repo shows Finance & Accounting
what VBA / Python / SQL can do inside a single workbook. This folder shows what the
same languages — plus PowerShell, Office Scripts, Power Automate, and Power Query —
can do **across the whole business** at a large software company like iPipeline.

**The rule I held myself to**
Every example in this folder does something that modern OneDrive for Business,
Excel (desktop or web), or Microsoft 365 cannot natively do. No "here's how to
sort a table in VBA". Every file solves a real problem where code is the
only reasonable answer.

---

## The 8 categories

| Folder | What it shows | Problem it solves |
|---|---|---|
| `01_VBA/` | Excel reaching out to the rest of the world | Excel cannot post to Teams, send Outlook mail merges with per-row attachments, create JIRA tickets, sync SharePoint lists, or run SQL Server queries on its own |
| `02_Python/` | Heavy-duty analytics and NLP | Python does what Excel can't: parse gigabyte log files, read PDFs, score customers with ML, compute rev rec at scale, and cross-reference APIs |
| `03_SQL/` | Warehouse-side analytics | Real metrics live in Snowflake/BigQuery/Postgres. SQL returns a board-ready number in seconds against 500M rows |
| `04_PowerShell/` | IT & infrastructure automation | AD audits, SharePoint storage scans, SSL expiry monitors, new-hire provisioning |
| `05_OfficeScripts/` | Excel-on-the-web automation | Run from the browser, trigger from Power Automate, post to Teams — no desktop install |
| `06_PowerAutomate/` | Cross-service orchestration | One flow, many services (SharePoint + Outlook + Teams + Planner + Azure AD) |
| `07_PowerQuery/` | Data ingestion at the edge of Excel | Paginated REST APIs, multi-folder merges with schema drift |

---

## 01_VBA — What's in each file

| File | One-liner |
|---|---|
| `modMailMerge_WithAttachments.bas` | Outlook mail merge with **a different PDF per recipient** (Word mail merge can't do this) |
| `modTeamsNotifier.bas` | Post Adaptive Cards to a Teams channel on threshold breach |
| `modJiraBridge.bas` | Create JIRA tickets in bulk from Excel rows + fetch ticket data back into a sheet |
| `modInvoicePDFGenerator.bas` | Generate one branded PDF invoice per customer row, save to per-customer folders, email them |
| `modSharePointSync.bas` | Two-way sync between Excel and a SharePoint list via REST API (the modern connector is read-only) |
| `modSQLServerRunner.bas` | Run a library of parameterized SQL queries and drop each result set on its own sheet |
| `modRenewalAlertEngine.bas` | Scan contract renewal dates, send per-owner alerts, escalate to manager inside 30 days |
| `modMultiWorkbookDiff.bas` | Cell-by-cell diff across N workbooks at once (native Compare is 2-files-only, legacy, and half-broken) |
| `modFolderOrganizer.bas` | Scan folder tree, let you write rename/move rules in Excel, apply them in bulk |
| `modCalendarAppointmentBuilder.bas` | Create 200 Outlook meetings or appointments from Excel rows in one click |
| `modSlackNotifier.bas` | Post Block Kit messages / tables to Slack channels from any macro |

---

## 02_Python — What's in each file

| File | One-liner |
|---|---|
| `saas_arr_waterfall.py` | MRR/ARR waterfall with all 5 SaaS movements (New, Expansion, Contraction, Churn, Ending) + NRR/GRR/Quick Ratio |
| `customer_churn_risk_scorer.py` | Risk-score every customer 0-100 with scikit-learn, explain each score with its top 3 drivers |
| `license_utilization_analyzer.py` | Find unused SaaS seats across every tool the company uses, quantify annualized waste |
| `aws_cost_optimizer.py` | Parse AWS CUR CSVs to find idle EC2, orphan EBS, oversized RDS, idle NAT gateways, RI/SP underutilization |
| `contract_pdf_extractor.py` | Bulk-extract structured terms (effective date, term, auto-renew, ACV, SLA, DPA) from 1000s of PDF contracts |
| `support_ticket_triage.py` | Auto-classify + sentiment + theme-cluster every support ticket; Mon-morning CS digest |
| `api_slo_tracker.py` | Parse log files, compute p50/p95/p99 + rolling 30-day error budget per endpoint |
| `revenue_recognition_engine.py` | ASC 606 recognition schedule + deferred revenue rollforward + commission amortization + exceptions |
| `jira_weekly_digest.py` | Cross-project JIRA digest: velocity, age of open issues, blockers, stale tickets |
| `cohort_retention_analyzer.py` | Logo + dollar retention by cohort, triangular retention heatmap |
| `email_to_structured_data.py` | Read a shared mailbox, extract invoice #, PO, amount, due date from bodies + PDF attachments |
| `sox_evidence_collector.py` | Auto-compile SOX/SOC 2 evidence quarterly: change tickets, deploys, terminations → access revocation |
| `git_developer_metrics.py` | Cross-repo engineering productivity: commit cadence, bus factor, weekly activity, cycle time |

---

## 03_SQL — What's in each file

| File | One-liner |
|---|---|
| `saas_metrics_suite.sql` | MRR/ARR waterfall, Rule of 40, Magic Number, CAC Payback, Burn Multiple — all in one file |
| `customer_360_view.sql` | Unified one-row-per-customer view joining CRM + billing + support + product usage + CS signals + pipeline |
| `sales_pipeline_velocity.sql` | Stage funnel, time-in-stage, stuck deals, weighted commit forecast |
| `data_quality_audit.sql` | Orphan FKs, unexpected NULLs, duplicate keys, frozen tables, calendar gaps, outlier amounts, line-sum mismatches |
| `revenue_recognition_schedule.sql` | Pure-SQL version of the rev rec engine: runs in the warehouse, powers BI dashboards |
| `slow_query_tuner.sql` | DMV + pg_stat_statements library: top-N slow queries, missing indexes, unused indexes, blocking, bloat |

---

## 04_PowerShell — What's in each file

| File | One-liner |
|---|---|
| `AD_Inactive_User_Audit.ps1` | Quarterly audit of users with no recent login — offboarding-ready Excel report |
| `SharePoint_Site_Storage_Audit.ps1` | Company-wide SharePoint + OneDrive storage audit, orphan detection, stale sites |
| `SSL_Cert_Expiry_Monitor.ps1` | Scheduled TLS cert expiry monitor across every public endpoint, alert via email + Teams |
| `New_Hire_Account_Provisioner.ps1` | Bulk AD account creation + group membership + welcome letter generation from a weekly HR CSV |

---

## 05_OfficeScripts — What's in each file

| File | One-liner |
|---|---|
| `TeamsWebhookOnThreshold.ts` | Excel-web script: read a "Watchers" sheet, post a Teams card when any threshold breaches |
| `DailyMetricsExport.ts` | Export a styled range to CSV content a Power Automate flow can save timestamped to SharePoint |
| `BulkFormatDataImport.ts` | Smart-format freshly-pasted data (dates, currency, percent) based on header names, add a styled table |

---

## 06_PowerAutomate — What's in each file

| File | One-liner |
|---|---|
| `PurchaseApproval_Flow.json` | Importable flow: SharePoint form → auto-approve <$500 / manager <$5K / CFO otherwise → SharePoint update + Outlook + Teams |
| `NewHire_Onboarding_Flow.json` | Importable flow: HR form triggers OneDrive folder + packet copy + Planner tasks + calendar invites + Teams welcome + AD membership + tracker update |

---

## 07_PowerQuery — What's in each file

| File | One-liner |
|---|---|
| `REST_API_Pagination.pq` | Generic M function: consume a paginated REST API with bearer token, flatten every page into one table |
| `Multi_Folder_Merge.pq` | Walk N folders, read every matching file, tolerate schema drift, tag each row with source file |

---

## How to use these at a large software business

- **Not all of these need to run in prod.** Some are one-time scripts (the SOX evidence collector, the AWS waste finder). Others are scheduled (SSL expiry, SLO tracker, JIRA digest). A few are event-driven (the Power Automate flows).
- **Each file starts with a `PURPOSE`, `WHY THIS IS NOT NATIVE`, and `USE CASE` block.** That doubles as the training material a coworker needs to understand why it exists.
- **Everything is standalone.** No hidden dependencies on the main iPipeline demo workbook. Any of these can be lifted into a different company tomorrow.

---

## Dependency notes

- **Python scripts:** `pip install -r 02_Python/requirements.txt`
- **PowerShell scripts:** Require RSAT (AD), PnP.PowerShell (SharePoint), and optionally ImportExcel (pretty Excel output). All via PSGallery.
- **VBA modules:** Require Windows Outlook, and either an open workbook with specific sheet layouts or named ranges for auth tokens.
- **Office Scripts:** Require an Excel-web file in SharePoint / OneDrive + a Microsoft 365 Business / Enterprise license.
- **Power Automate flows:** Require at minimum the Office 365, SharePoint, and Teams connectors; some require the premium Approvals connector.
- **SQL:** Written in PostgreSQL / Snowflake-flavored SQL with SQL Server alternatives where needed. See in-file comments.

---

## One more thing

The main APCLDmerge project (parent folder) is the Finance & Accounting demo —
62 Command Center actions inside one Excel workbook. This folder is the
**"what else is possible"** showcase. Coworkers can watch the video for the
main project, then pick one of these as their next learning step.

*Built for iPipeline — 2026-04-21*
