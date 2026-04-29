# New Idea List Based on RecTrial

## Finance Copilot Launcher (One-Click Menu)

- Simple: Put all the tools behind one easy menu so people click a number instead of remembering script names.
- Technical: Build a Python CLI (`finance_copilot.py`) that reads a config file (JSON/YAML), lists tools by category, validates inputs, and runs selected scripts with friendly prompts.
- Category: Universal
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: RecTrial already has many scripts; this turns a “tool pile” into one guided user experience.
- Build Difficulty: Easy

## Reconciliation Control Tower Dashboard

- Simple: A scoreboard that shows which reconciliations are done, stuck, or risky.
- Technical: Use Power Query to combine reconciliation output files, load into Power BI, and build DAX KPIs (open items, aged exceptions, unresolved dollar amount, completion %).
- Category: Universal
- Best Tool: Power BI / DAX
- Demo Value: High
- Why It Fits Rectrial: RecTrial already has reconciliation scripts and exception outputs that can feed a control dashboard.
- Build Difficulty: Medium

## Data Contract Checker for Incoming Files

- Simple: A gatekeeper that checks if a file is “in the right shape” before processing.
- Technical: Python validates required columns, datatypes, date formats, duplicates, and null thresholds; outputs a PASS/FAIL report sheet and a remediation checklist.
- Category: Universal
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: Extends existing sanitizer/quality tools and prevents bad data from entering downstream workflows.
- Build Difficulty: Easy

## Exception Triage Scoring Engine

- Simple: Auto-rank the most important problems first so teams fix the biggest issues before small ones.
- Technical: Use Python or DAX scoring model: `priority = impact * confidence * recency_weight`; generate top-20 action list and owner assignment export.
- Category: Universal
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: Matches existing variance/reconciliation focus and improves “what to do first” decisions.
- Build Difficulty: Medium

## Variance Narrative Builder (Finance-Friendly Text)

- Simple: Turn big number changes into plain-English bullet points for leaders.
- Technical: Python template engine + rule logic (materiality thresholds, YoY/MoM logic, trend direction labels) to produce narrative text and “talking points” sheets.
- Category: Universal
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: Builds directly on variance modules and reporting outputs already present.
- Build Difficulty: Medium

## Workbook Dependency Risk Map

- Simple: Show which sheets/formulas depend on each other so you can see where one change could break many things.
- Technical: Parse formulas with Python/openpyxl, build dependency graph, output HTML network view plus a high-risk node list.
- Category: File-Dependent
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: Fits the existing “show what Excel alone can’t easily show” storyline.
- Build Difficulty: Medium

## ARR/MRR Waterfall Builder

- Simple: Turn subscription changes into a clean “start-to-end revenue waterfall” chart.
- Technical: Python ingests subscription transaction data, computes New/Expansion/Contraction/Churn bridges, writes waterfall table, and renders branded chart (matplotlib/plotly).
- Category: File-Dependent
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: Aligns with current Video 4 finance storytelling around SaaS metrics.
- Build Difficulty: Medium

## Monthly Close Evidence Pack Generator

- Simple: One button creates an audit-ready folder with reports, checks, hashes, and a manifest.
- Technical: Python zips source outputs + metadata, creates checksum manifest (SHA-256), writes a control summary document, and timestamps package.
- Category: Universal
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: Supports governance and audit-readiness themes already present in docs and toolset.
- Build Difficulty: Medium

## Power Query “Universal Ingestion Template”

- Simple: A reusable import template that cleans many messy files the same way every month.
- Technical: Build parameterized Power Query functions for folder import, column standardization, type enforcement, and exception tagging; output clean model tables.
- Category: Universal
- Best Tool: Power Query
- Demo Value: Medium
- Why It Fits Rectrial: Complements VBA/Python sanitizers with a no-code option for Excel-first users.
- Build Difficulty: Medium

## SQL Reconciliation View Generator

- Simple: Compare two data sources in SQL and instantly see matched, missing, and mismatched records.
- Technical: Use SQL templates with configurable keys/amount/date tolerances; output three views (`matched`, `missing`, `mismatch`) and summary counts by reason code.
- Category: Universal
- Best Tool: SQL
- Demo Value: Medium
- Why It Fits Rectrial: Extends existing SQL staging/validation patterns into a reusable reconciliation framework.
- Build Difficulty: Medium

## VBA “Quick Demo Mode” Macro

- Simple: A shortcut button that runs a curated set of 5–8 impressive tools in sequence.
- Technical: Add a VBA orchestrator subroutine that calls existing toolkit procedures with safe defaults and logs each step to an audit sheet.
- Category: File-Dependent
- Best Tool: VBA
- Demo Value: High
- Why It Fits Rectrial: Leverages rich VBA inventory and improves live demo reliability for business audiences.
- Build Difficulty: Easy

## Cross-File Master Data Mapper

- Simple: Teach the system that “Cust ID,” “Customer_ID,” and “Client Number” mean the same thing.
- Technical: Maintain mapping dictionary table, apply fuzzy matching + manual override layer, then standardize columns before consolidation/reconciliation.
- Category: Universal
- Best Tool: Python
- Demo Value: Medium
- Why It Fits Rectrial: Natural extension of existing mapping/cleaning/fuzzy tools in UniversalToolkit.
- Build Difficulty: Medium

## Forecast Backtest Scorecard

- Simple: Compare last month’s forecast vs what really happened and grade forecast quality.
- Technical: Python computes MAPE, bias, and error bands by account/entity; Power BI displays trend of forecast accuracy over time.
- Category: File-Dependent
- Best Tool: Python + Power BI / DAX
- Demo Value: High
- Why It Fits Rectrial: Strengthens forecast modules with measurable model accountability.
- Build Difficulty: Medium

## AP Duplicate and Near-Duplicate Detector (Explainable)

- Simple: Catch likely duplicate invoices and explain why each one was flagged.
- Technical: Python rule engine + fuzzy scoring on vendor, invoice number, date, amount; output confidence score and reason tags per match pair.
- Category: Universal
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: Directly aligns with existing finance validation and reconciliation objectives.
- Build Difficulty: Medium

## Executive Brief Pack in 3 Formats

- Simple: Create the same executive summary as Excel tab, PDF, and PowerPoint notes.
- Technical: Python composes a single narrative dataset and exports to multiple outputs (`openpyxl`, `python-docx`/PDF workflow, PPT template injection).
- Category: File-Dependent
- Best Tool: Python
- Demo Value: High
- Why It Fits Rectrial: Builds on existing exec-brief and reporting scripts while improving executive communication.
- Build Difficulty: Hard
