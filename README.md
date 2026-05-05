# APCLDmerge — P&L Automation Toolkit

> A world-class demo Profit & Loss Excel file with VBA macros, SQL scripts, and Python tools built for Finance & Accounting at iPipeline. Designed for presentation to 2,000+ employees and senior leadership (CFO/CEO).

---

## 📌 What This Project Is

This repository contains everything needed to build, run, and maintain an automated P&L reporting system. The toolkit combines:

- **Excel + VBA** — A polished, macro-driven P&L workbook for month-end close, variance analysis, dashboards, PDF exports, and more
- **SQL** — Scripts for data staging, transformation, validation, and reporting enhancements
- **Python** — Scripts for forecasting, AP matching, month-end automation, email reporting, dashboards, and CLI tools
- **Docs** — Step-by-step setup guides, training materials, operations runbooks, and quick-start references

---

## 📁 Folder Structure

```
claude-training-lab-code/
│
├── excel/                        ← Excel workbook files (.xlsx)
│   └── KeystoneBenefitTech_PL_Model.xlsx
│
├── vba/                          ← VBA modules (.bas files) — import into Excel
│   ├── modConfig_v2.1.bas
│   ├── modDashboard_v2.1.bas
│   ├── modDataQuality_v2.1.bas
│   ├── modFormBuilder_v2.1.bas
│   ├── modMasterMenu_v2.1.bas
│   ├── modMonthlyTabGenerator_v2.1.bas
│   ├── modNavigation_v2.1.bas
│   ├── modPDFExport_v2.1.bas
│   ├── modPerformance_v2.1.bas
│   ├── modReconciliation_v2.1.bas
│   ├── modSearch_v2.1.bas
│   ├── modVarianceAnalysis_v2.1.bas
│   └── frmCommandCenter_code.txt  ← UserForm code-behind (manual paste)
│
├── sql/                          ← SQL scripts
│   ├── staging.sql
│   ├── transformations.sql
│   ├── validations.sql
│   └── pnl_enhancements.sql
│
├── python/                       ← Python automation scripts
│   ├── pnl_runner.py             ← Main entry point (run this first)
│   ├── pnl_config.py
│   ├── pnl_dashboard.py
│   ├── pnl_month_end.py
│   ├── pnl_forecast.py
│   ├── pnl_allocation_simulator.py
│   ├── pnl_ap_matcher.py
│   ├── pnl_cli.py
│   ├── pnl_email_report.py
│   ├── pnl_snapshot.py
│   ├── pnl_tests.py
│   └── requirements.txt          ← Python package list
│
├── docs/
│   ├── setup/                    ← How to set up the workbook from scratch
│   │   ├── QUICK_START.md
│   │   ├── IMPLEMENTATION_GUIDE.md
│   │   ├── START_TO_FINISH_GUIDE.md
│   │   ├── WORKBOOK_SETUP_NOTES.md
│   │   └── KBT_File_Map.pdf
│   ├── day-to-day/               ← Guides for everyday use
│   │   ├── OPERATIONS_RUNBOOK.md
│   │   ├── SANITIZATION_PLAYBOOK.md
│   │   └── USER_TRAINING_GUIDE.md
│   ├── overview/                 ← High-level project docs
│   │   ├── EXECUTIVE_SUMMARY.md
│   │   └── ARCHITECTURE_DIAGRAM.md
│   └── ai-tools/                 ← VBA macro reference libraries (AI-generated)
│       ├── GPT.md
│       ├── Gemini.md
│       └── Perlex.md
│
├── training/                     ← Training materials for coworkers
│   └── README.md
│
├── qa/                           ← QA tracking, test plans, and validation reports
│   ├── CHANGELOG.md
│   ├── TEST_PLAN.md
│   ├── VALIDATION_REPORT.md
│   ├── INTEGRATION_TEST_GUIDE.md
│   ├── ISSUE_CLOSURE.md
│   └── logging_template.csv
│
├── tasks/                        ← Session management (internal use)
│   ├── todo.md                   ← Running task list
│   └── lessons.md                ← Lessons learned log
│
├── CLAUDE.md                     ← Instructions for the AI assistant
└── README.md                     ← This file
```

---

## 🚀 Where to Start

**First time here? Go to:**
👉 [`docs/setup/QUICK_START.md`](docs/setup/QUICK_START.md) — Get up and running in 10 minutes

**Setting up the Excel workbook?**
👉 [`docs/setup/IMPLEMENTATION_GUIDE.md`](docs/setup/IMPLEMENTATION_GUIDE.md) — Full step-by-step workbook setup

**Learning how to use the tool day-to-day?**
👉 [`docs/day-to-day/USER_TRAINING_GUIDE.md`](docs/day-to-day/USER_TRAINING_GUIDE.md) — All 50 commands explained in plain English

**Running the Python tools?**
👉 [`python/pnl_runner.py`](python/pnl_runner.py) — The single entry point for all Python commands

---

## 🧰 Current Version

| Area        | Version | Last Updated |
|-------------|---------|--------------|
| VBA Modules | v2.1.0  | 2026-02-20   |
| Python Scripts | v2.1.0 | 2026-02-20  |
| Documentation | v2.1   | 2026-02-20   |

See [`qa/CHANGELOG.md`](qa/CHANGELOG.md) for the full version history.

---

## 📋 Project Status

See [`tasks/todo.md`](tasks/todo.md) for the current task list and what's coming next.

---

## 🏢 About This Project

Built for the Finance & Accounting team at **iPipeline**. All guides and training materials are written in plain English — no technical background required.

Questions? Contact the project owner or review the lessons log at [`tasks/lessons.md`](tasks/lessons.md).