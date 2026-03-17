# Universal Excel & Python Toolkit — Complete Tool Reference
## 120 Tools Across 20 Categories

---

## Category 1: Data Cleaning & Sanitization (Tools 01–10)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 01 | Universal Whitespace Cleaner | VBA | Removes leading, trailing, and duplicate internal spaces including non-breaking spaces and tabs from all text cells |
| 02 | Blank Row & Column Remover | VBA | Identifies and removes entirely blank rows and/or columns within the used range |
| 03 | Non-Printable Character Stripper | VBA | Finds and removes invisible control characters (ASCII 0–31, 127–159) with preview before changes |
| 04 | Text Case Standardizer | VBA | Converts text to UPPER, lower, Title, or Sentence case with abbreviation exception handling |
| 05 | Smart Find & Replace | VBA | Multi-pair find-replace in a single order-safe pass with preview, formula protection, and no chaining |
| 06 | Date Format Unifier | Python | Converts mixed date formats in a column to one consistent format, flagging ambiguous dates for review |
| 07 | Encoding & Character Set Fixer | Python | Auto-detects and converts CSV file encoding to clean UTF-8, fixing garbled characters |
| 08 | HTML & Tag Stripper | VBA | Removes HTML/XML tags and converts HTML entities to plain text in cells from web-system exports |
| 09 | Duplicate Space & Punctuation Cleaner | VBA | Collapses duplicate internal spaces and normalizes punctuation inconsistencies |
| 10 | Column Content Profiler | Python | Profiles every column in a file: type distribution, min/max, unique counts, null rates, quality score |

## Category 2: Number & Format Standardization (Tools 11–17)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 11 | Text-to-Number Converter | VBA | Converts numbers stored as text to true numeric values, handling currency symbols, commas, parentheses |
| 12 | Currency Format Standardizer | VBA | Standardizes mixed currency display formats (US, European, UK) to one consistent format |
| 13 | Number Precision Standardizer | VBA | Sets consistent decimal precision and optionally rounds underlying values, flagging display vs value mismatches |
| 14 | Percentage Format Normalizer | VBA | Detects and standardizes the decimal (0.15) vs whole-number (15) percentage convention |
| 15 | Phone Number Formatter | VBA | Standardizes phone numbers to a consistent format, flagging unrecognized patterns for review |
| 16 | Postal Code & ZIP Formatter | VBA | Restores leading zeros on ZIP codes, standardizes postal code formats, converts column to text |
| 17 | Unit & Measurement Standardizer | Python | Detects mixed unit labels in cells ("5 hours", "300 minutes") and converts to a common unit |

## Category 3: Duplicate & Error Detection (Tools 18–24)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 18 | Exact Duplicate Finder & Highlighter | VBA | Finds exact duplicate rows based on user-selected key columns, highlights and marks originals vs duplicates |
| 19 | Near-Duplicate Detector | Python | Fuzzy string matching to find likely-same entities entered differently ("Microsoft Corp" vs "Microsoft Corporation") |
| 20 | Cross-Column Duplicate Checker | VBA | Finds values that appear in multiple columns when they should be mutually exclusive |
| 21 | Unique Value Extractor | VBA | Extracts sorted unique values from any column with optional frequency counts |
| 22 | Blank Cell Locator & Reporter | VBA | Reports every blank cell by column with severity classification and percentage completeness |
| 23 | Data Type Mismatch Detector | VBA | Flags cells where the data type differs from the column's dominant type |
| 24 | Formula Error Finder | VBA | Finds all formula errors (#N/A, #REF!, etc.) with plain-English explanations and impact ranking |

## Category 4: Audit & Validation (Tools 25–32)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 25 | Formula Consistency Checker | VBA | Flags rows where a formula was replaced by a hardcoded value or where the formula pattern breaks |
| 26 | Data Completeness Scorecard | Python | Grades a dataset on 5 quality dimensions (A through F) with a formatted one-page scorecard |
| 27 | Column Type Validator | VBA | Validates cells against user-defined type rules (integer, date, pattern match, number range, valid list) |
| 28 | Broken Reference Auditor | VBA | Scans all formulas for #REF! errors, broken external links, missing sheet references, circular references |
| 29 | Data Boundary Detector | VBA | Reports the true data boundary vs Excel's UsedRange, identifies data islands and inflated ranges |
| 30 | Header Validator | VBA | Validates column headers against a required schema with fuzzy near-match suggestions for missing headers |
| 31 | SQL-Style Query Tool | Python | Runs SQL SELECT queries against any Excel/CSV file using DuckDB with zero database setup |
| 32 | Workbook Dependency Mapper | VBA | Maps sheet-to-sheet dependencies, ranks critical cells, lists external workbook references |

## Category 5: Finance & Accounting Specific (Tools 33–40)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 33 | Account Code Format Checker | VBA | Validates account codes against a configurable format pattern, flags invalid codes with near-match suggestions |
| 34 | Journal Entry Validator | VBA | Pre-submission validation: debits = credits, required fields populated, dates in posting period |
| 35 | Aging Bucket Calculator | VBA | Calculates days aged and assigns configurable aging buckets (Current, 1-30, 31-60, etc.) with summary table |
| 36 | Variance Analyzer | VBA | Calculates actual vs budget variance with absolute, percentage, and favorable/unfavorable classification |
| 37 | Period-End Close Checklist Validator | VBA | Validates close checklist completion: all tasks done, approvals in place, exceptions documented |
| 38 | Intercompany Transaction Identifier | Python | Flags likely intercompany transactions using entity name matching and IC account code ranges |
| 39 | Running Total & Subtotal Validator | VBA | Recalculates subtotals from detail rows and flags any that don't match the displayed value |
| 40 | Balance Sheet Tie-Out Checker | VBA | Verifies Assets = Liabilities + Equity and checks configurable subtotal tie-out points |

## Category 6: Reporting & Summarization (Tools 41–47)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 41 | Dynamic Summary Table Generator | VBA | Builds grouped summary tables (SUM, COUNT, AVG) as clean static output — no pivot table dependency |
| 42 | Multi-Sheet Rollup Consolidator | VBA | Stacks data from multiple same-structure sheets into one master sheet with source tracking |
| 43 | Exception-Only Report Builder | VBA | Extracts only rows meeting user-defined criteria into a standalone report sheet |
| 44 | Period-Over-Period Comparison | Python | Compares two versions of a report, shows every changed value with old/new/delta |
| 45 | Conditional Narrative Generator | Python | Reads a summary table and generates plain-English executive summary bullet points |
| 46 | Top-N / Bottom-N Ranker | VBA | Extracts top N and bottom N records by any numeric column into a formatted report |
| 47 | Cross-Tab Summary Builder | Python | Builds a two-dimensional cross-tabulation table from flat data with row/column totals |

## Category 7: Workbook & Sheet Management (Tools 48–53)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 48 | Bulk Sheet Renamer | VBA | Renames multiple sheet tabs at once using a mapping table or pattern rule |
| 49 | Hidden Sheet Revealer & Auditor | VBA | Lists all hidden/very-hidden sheets with content summaries, offers to unhide |
| 50 | Formula-to-Value Converter | VBA | Replaces formulas with values (all or by type) with automatic backup |
| 51 | Sheet Splitter by Value | Python | Splits a dataset into separate sheets or files based on unique values in a column |
| 52 | Workbook Metadata Reporter | VBA | One-page summary of any workbook: sheet counts, row counts, formula counts, external links, file size |
| 53 | Sheet Structure Comparer | VBA | Compares column headers and layout of two sheets, reports every structural difference |

## Category 8: Cross-File Operations (Tools 54–60)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 54 | Multi-File Consolidator | Python | Reads all Excel/CSV files in a folder and consolidates into one master file with source tracking |
| 55 | Folder-Wide Search & Replace | Python | Find-replace across every Excel file in a folder with backup and change log |
| 56 | Cross-Workbook Lookup | Python | VLOOKUP-style matching between two separate files without opening both, reports unmatched records |
| 57 | File Comparison & Diff Tool | Python | Cell-by-cell comparison of two file versions with color-coded diff report |
| 58 | Folder Inventory Scanner | Python | Catalogs every Excel/CSV in a folder: file name, sheets, row counts, headers, last modified |
| 59 | Batch Header Standardizer | Python | Renames column headers across all files in a folder to match a standard mapping |
| 60 | Master-Detail Linker | Python | Joins master and detail files, flags orphaned details and inactive master records |

## Category 9: Data Transformation & Reshaping (Tools 61–67)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 61 | Wide-to-Long Reshaper | Python | Unpivots wide tables (months as columns) into long format (one row per observation) for analytics |
| 62 | Long-to-Wide Pivot Builder | Python | Converts long-format data into wide-format reports (reverse of Tool 61) |
| 63 | Column Splitter | VBA | Splits one column into multiple columns based on a delimiter (non-destructive) |
| 64 | Column Merger | VBA | Combines multiple columns into one with configurable separator, handles blanks intelligently |
| 65 | Row-to-Column Transposer | VBA | Transposes ranges with proper header handling to a new sheet (preserves original) |
| 66 | Nested Data Flattener | Python | Fills down parent values into child rows so every row is a complete record |
| 67 | Multi-Row Record Merger | Python | Merges multi-row legacy records into single rows based on a key column |

## Category 10: Power BI & CSV Prep (Tools 68–72)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 68 | Power BI-Ready Formatter | Python | One-pass transformation: removes merged cells, fixes headers, removes blanks/subtotals, standardizes types |
| 69 | CSV Cleaner & Encoding Standardizer | Python | Fixes encoding, delimiter, quoting, line endings, and BOM in one pass |
| 70 | Column Type Enforcer | Python | Forces column values to a specified type, logging every failed conversion |
| 71 | Lookup Table Generator | VBA | Extracts unique values into a formatted lookup table with ID, value, description, and status columns |
| 72 | Data Model Relationship Mapper | Python | Discovers potential join relationships between sheets by analyzing value overlap |

## Category 11: PDF & Export (Tools 73–76)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 73 | Excel-to-PDF Batch Converter | Python | Converts all Excel files in a folder to PDF with configurable settings |
| 74 | PDF Table Extractor | Python | Reads tables from PDFs into clean Excel data, handles multi-page tables |
| 75 | Print Area Optimizer | VBA | Auto-detects data range and configures all print settings for clean output in one click |
| 76 | Multi-Sheet PDF Packager | VBA | Exports selected sheets as one combined PDF with auto-generated table of contents |

## Category 12: Modern / High-Impact (Tools 77–82)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 77 | Two-File Reconciliation Engine | Python | Three-way recon: matched, File A only, File B only, matched with differences — with tolerance |
| 78 | Anomaly & Outlier Detector | Python | Statistical outlier detection (IQR or Z-Score) with distribution chart and flagged value report |
| 79 | Smart Data Merge | Python | Merges two datasets with conflict detection — surfaces both values for human resolution |
| 80 | Change Log Generator | VBA | Snapshots a sheet and on re-run produces a detailed log of every cell added, modified, or deleted |
| 81 | Template Compliance Checker | VBA | Validates a workbook against a required template: correct sheets, headers, named ranges, structure |
| 82 | File Health Scorecard | Python | Comprehensive health assessment across 6 dimensions with letter grade (A–F) and recommendations |

## Category 13: Reconciliation & Matching (Tools 83–87)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 83 | Three-Way Invoice Matcher | Python | Matches POs, receiving reports, and invoices — flags price/quantity variances and unmatched items |
| 84 | Bank Statement-to-GL Reconciler | Python | Matches bank transactions to GL entries with date proximity, produces standard bank rec output |
| 85 | Subledger-to-GL Tie-Out Tool | Python | Compares subledger detail totals to GL balances, drills into transactions causing differences |
| 86 | Vendor Statement Reconciler | Python | Matches vendor statement against internal AP records, identifies discrepancies |
| 87 | Cash Application Matcher | Python | Matches payments to open invoices using amount, reference, and combination matching |

## Category 14: Budget, Forecast & FP&A (Tools 88–93)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 88 | Forecast Accuracy Calculator | Python | Calculates MAPE, bias, and tracking signal comparing forecasts to actuals over multiple periods |
| 89 | Budget Reallocation Modeler | VBA | Distributes a budget adjustment across line items (proportional, equal, weighted, protect-and-spread) |
| 90 | Waterfall Bridge Data Builder | Python | Structures data for waterfall/bridge charts with auto-generated chart |
| 91 | Rolling Forecast Shifter | VBA | Shifts a rolling forecast forward by one period, replacing expired forecast with actuals |
| 92 | Headcount-to-Expense Calculator | Python | Converts a hiring plan into monthly salary and benefits expense projections with proration |
| 93 | Scenario Comparison Table Builder | VBA | Produces a formatted base-vs-scenarios table with deltas, percentages, and favorable/unfavorable flags |

## Category 15: Billing, Revenue & AR (Tools 94–98)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 94 | Usage-Based Billing Calculator | Python | Applies tiered/volume pricing rules to raw usage data, produces per-customer billing schedule |
| 95 | Revenue Schedule Builder | Python | Builds monthly revenue recognition schedules from contract data (straight-line or usage-based) |
| 96 | Customer Concentration Analyzer | Python | Pareto analysis, HHI calculation, and top-N customer percentage for concentration risk reporting |
| 97 | DSO / DPO / DIO Calculator | Python | Calculates Days Sales Outstanding, Days Payable Outstanding, Days Inventory Outstanding with trend |
| 98 | Deferred Revenue Rollforward Builder | Python | Builds Beginning + Bookings - Recognized = Ending schedule by customer/product by month |

## Category 16: HR, Payroll & People Data (Tools 99–103)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 99 | Org Chart Data Validator | Python | Checks for valid managers, circular chains, orphaned records, single root, department consistency |
| 100 | Headcount Reconciler | Python | Compares headcount across 2-3 systems (HRIS, Payroll, Finance), identifies discrepancies |
| 101 | Termination Date Cross-Checker | VBA | Finds terminated employees still in active benefit, access, or project lists |
| 102 | Compensation Band Validator | Python | Checks every salary against defined band min/max, flags outliers with deviation percentage |
| 103 | PTO & Leave Balance Auditor | Python | Validates PTO accruals against policy rules, checks carry-over limits and math accuracy |

## Category 17: Compliance & Audit Trail (Tools 104–108)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 104 | Segregation of Duties Checker | Python | Checks user-role matrix for SoD violations based on configurable conflict rules |
| 105 | Audit Sample Selector | Python | Statistically valid sampling (random, stratified, monetary unit) with methodology documentation |
| 106 | Policy Threshold Screener | VBA | Flags transactions exceeding policy thresholds and checks for required control evidence |
| 107 | Year-Over-Year Trend Analyzer | Python | Multi-year comparison flagging unusual YoY changes for audit analytical procedures |
| 108 | Data Extraction Log Builder | VBA | Auto-logs who extracted data, when, from where, with what filters — SOX audit trail |

## Category 18: Operations & Project Data (Tools 109–113)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 109 | SLA Compliance Calculator | Python | Calculates SLA compliance rates using business-hours math, identifies every breach |
| 110 | Project Cost Tracker Consolidator | Python | Consolidates project budgets into a portfolio view with RAG status per project |
| 111 | Resource Utilization Calculator | Python | Calculates billable vs non-billable utilization rates per person/team/project |
| 112 | Milestone Date Slippage Tracker | VBA | Compares planned vs actual milestone dates, calculates cumulative slip and trend |
| 113 | Vendor Scorecard Builder | Python | Weighted multi-dimension vendor scoring with tier classification (A/B/C/D) |

## Category 19: Data Migration & System Prep (Tools 114–117)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 114 | Import File Formatter | Python | Reformats any source file to match a target system's import template with validation |
| 115 | Code Mapping Translator | Python | Applies old-to-new code translation across an entire dataset with audit trail |
| 116 | Data Migration Validation Suite | Python | Full pre/post-migration checks: record counts, sums, nulls, referential integrity, format compliance |
| 117 | Legacy Data Cleaner | Python | Combined cleaning pipeline (encoding + whitespace + characters + dates + numbers + case + dedup) in one run |

## Category 20: Smart Analysis & Insights (Tools 118–120)
| # | Name | Language | Purpose |
|---|------|----------|---------|
| 118 | Automatic Data Dictionary Generator | Python | Auto-generates a formatted data dictionary from any file: types, ranges, samples, descriptions |
| 119 | Spend Categorization Engine | Python | Assigns expense categories to transaction descriptions using keyword matching with learning |
| 120 | Trend & Seasonality Detector | Python | Detects trend, seasonality, anomalies, and structural breaks in time-series data with plain-English summary |
