# Prompt for Claude Opus — Build DemoKit Industries Excel File

Copy everything below this line and paste it into a new Claude Opus chat session.

---

## YOUR TASK

Build me a world-class, professional Excel file (.xlsx) called **DemoKit_Industries_Q1_2026.xlsx**.

This file represents the **"DemoKit Industries — Q1 2026 Financial Close Package"** — a fictional mid-size company's quarterly close package. It must look like a real file that a Finance team at a Fortune 500 company would produce.

**This file will be used in a live demo in front of 2,000+ employees and the CFO/CEO.** Every detail matters. It must be polished, realistic, and impressive.

### PURPOSE

This file is a **demo file for a Universal Toolkit** — a set of 78 VBA tools that clean, audit, format, and analyze Excel files. The file needs to be **intentionally seeded with common real-world data problems** so that each tool has something to find and fix during the demo. The problems must be subtle and realistic — the kind of things that actually happen in Finance teams, not obvious joke data.

---

## BRANDING (iPipeline Corporate Brand)

**DO NOT apply branding to this file.** Leave it unbranded on purpose — one of the tools (ApplyiPipelineBranding) will brand it live during the demo. The file should look like a **generic, unbranded corporate spreadsheet** — professional structure but no special colors or formatting beyond basic Excel defaults.

- Use default Excel fonts (Calibri is fine — the branding tool will switch to Arial)
- Use default colors — no custom fill colors on headers
- Bold the header rows but do NOT color them
- Keep it clean and structured but intentionally plain

---

## FILE STRUCTURE — 6 SHEETS

### Sheet 1: "Trial Balance"

**Purpose:** Classic trial balance report — the backbone of any quarterly close.

**Columns:** A=Account Number | B=Account Description | C=Debit | D=Credit

**Data — ~80 rows:**
- Use realistic 4-digit GL account numbers following standard chart of accounts:
  - 1000-1999: Assets (Cash, AR, Inventory, Prepaid, Fixed Assets, Accum Depreciation)
  - 2000-2999: Liabilities (AP, Accrued Liabilities, Notes Payable, Deferred Revenue)
  - 3000-3999: Equity (Common Stock, Retained Earnings, APIC)
  - 4000-4999: Revenue (Product Revenue, Service Revenue, Licensing Revenue, Other Income)
  - 5000-5999: COGS (Direct Materials, Direct Labor, Manufacturing Overhead, Freight)
  - 6000-6999: Operating Expenses (Salaries, Rent, Utilities, Insurance, Depreciation, Marketing, Travel, Professional Fees, Software Licenses, Office Supplies)
  - 7000-7999: Other Expense (Interest Expense, Loss on Disposal)
- Use realistic dollar amounts — Assets in the $50K-$5M range, Revenue lines $500K-$3M, Expenses $10K-$800K
- Total Debits and Total Credits should be on the last row

**Seeded Problems (7):**
1. **Total Debits ≠ Total Credits** — off by exactly $0.03 (TrialBalanceChecker will catch this)
2. **3 cells with text-stored numbers** — type an apostrophe before the number in cells C15, D28, and C45 so they appear as text (ConvertTextToNumbers / DataSanitizer will fix)
3. **2 cells with external link formulas** — in cells C10 and D10, put formulas like `='[OldWorkbook.xlsx]Sheet1'!B10` and `='[Q4_2025_TB.xlsx]Data'!C22` (ExternalLinkFinder will catch these)
4. **1 cell with #REF! error** — cell C62 should show #REF! (WorkbookErrorScanner will find it)
5. **1 cell with floating-point tail** — cell C30 should be 125000.0000000001 instead of 125000 (FixFloatingPointTails will fix)
6. **2 cells with leading spaces** — Account Description cells B20 and B55 should have 2-3 leading spaces before the text (RemoveLeadingTrailingSpaces will fix)
7. **1 named range that points to a deleted reference** — create a Named Range called "PriorYearTB" that points to a non-existent sheet like `=OldSheet!A1:D80` (NamedRangeAuditor will catch this)

---

### Sheet 2: "AP Aging Detail"

**Purpose:** Accounts Payable aging report — every vendor invoice with payment status.

**Columns:** A=Vendor ID | B=Vendor Name | C=Invoice # | D=Invoice Date | E=Due Date | F=Amount | G=PO Number | H=Status

**Data — ~100 rows:**
- Use 25-30 unique realistic vendor names (office supplies, IT services, consulting, utilities, building maintenance, insurance, legal, marketing agencies, software companies, etc.)
- Example vendors: Vertex Office Solutions, Meridian IT Services, ClearView Consulting Group, Apex Building Maintenance, DataStream Software Inc., Wellington Legal Partners, BrightPath Marketing, NorthStar Insurance Corp., etc.
- Vendor IDs: V-1001 through V-1030 format
- Invoice numbers: INV-20260001 through INV-20260100 format
- Dates should span Jan 1, 2026 through Mar 31, 2026
- Amounts should range from $250 to $175,000
- PO numbers: PO-2026-0001 format (leave ~10 blank for variety)
- Status values: "Paid", "Open", "Past Due", "Disputed" (mix realistically — ~60% Paid, 20% Open, 15% Past Due, 5% Disputed)

**Seeded Problems (8):**
1. **3 duplicate invoices** — same Vendor + same Amount + date within 3 days of each other. Make them subtle (not consecutive rows). DuplicateInvoiceDetector + ExactDuplicateFinder will catch these
2. **Mixed date formats** — Most dates in M/D/YYYY format, but seed ~8 dates in different formats: "Jan 15, 2026" (3 cells), "2026-01-22" (3 cells), "15-Jan-2026" (2 cells). DateFormatStandardizer will normalize these
3. **5 blank Vendor IDs** — leave column A empty on rows 12, 34, 56, 78, 91. GenerateUniqueCustomerIDs will fill these
4. **2 negative amounts** — cells F44 and F87 should be -$1,250.00 and -$3,400.00 (credit memos mixed into AP — FindNegativeAmounts will flag)
5. **3 suspiciously round numbers** — $50,000.00, $100,000.00, $25,000.00 exactly (FindSuspiciousRoundNumbers will flag)
6. **2 cells with non-breaking spaces** — in Vendor Name cells B22 and B65, use a non-breaking space (char 160) between words instead of regular space. UniversalWhitespaceCleaner will fix
7. **1 #N/A error** — cell F55 shows #N/A instead of a dollar amount (ReplaceErrorValues will fix)
8. **Amount column has no number formatting** — raw numbers with no commas or dollar signs (NumberFormatStandardizer / CurrencyFormatStandardizer will fix)

---

### Sheet 3: "GL Journal Entries"

**Purpose:** General ledger journal entries posted during Q1 — the messiest sheet on purpose.

**Columns:** A=JE Number | B=Date | C=Account # | D=Account Description | E=Debit | F=Credit | G=Memo

**Data — ~60 rows across 12-15 journal entries:**
- Journal entry numbers: JE-2026-001 through JE-2026-015
- Each JE has 2-6 lines (debits and credits)
- Entries should represent realistic transactions: payroll accrual, rent expense, revenue recognition, depreciation, prepaid amortization, AP payment run, bank reconciliation adjustment, intercompany transfer, bad debt write-off, insurance expense allocation
- Most JEs should balance (total debits = total credits per JE)
- Memos should be realistic: "Q1 payroll accrual - all departments", "March rent - HQ building", "Revenue recognition - Project Alpha", etc.

**Seeded Problems (7):**
1. **Merged cells in JE Number column** — merge cells A2:A5 for JE-2026-001, A6:A9 for JE-2026-002, etc. for the first 5 JEs. UnmergeAndFillDown will fix these
2. **2 unbalanced journal entries** — JE-2026-008 should be off by $500 (debit side heavy), JE-2026-013 should be off by $0.50. JournalEntryValidator will flag these
3. **8 blank rows** — scattered between journal entry groups (rows 10, 20, 28, 35, 40, 48, 53, 58). DeleteBlankRows will clean these up
4. **3 cells with #DIV/0! errors** — in the Debit or Credit columns. WorkbookErrorScanner will find them
5. **Phantom hyperlinks** — add hyperlinks to 4-5 cells in the Memo column that link to nowhere/fake URLs. PhantomHyperlinkPurger will remove them
6. **2 cells with extra whitespace** — double spaces in Memo text. RemoveLeadingTrailingSpaces will fix
7. **No freeze panes** — header row should NOT be frozen (FreezeTopRowAllSheets will fix this across all sheets)

---

### Sheet 4: "Budget vs Actual"

**Purpose:** Department budget comparison — the CFO's favorite report.

**Columns:** A=Department | B=Category | C=Budget Q1 | D=Actual Q1 | E=$ Variance | F=% Variance

**Data — ~30 rows:**
- 8-10 departments: Finance, Marketing, Sales, Engineering, HR, Operations, Legal, IT, Executive, Customer Support
- Categories per department: Salaries, Benefits, Travel, Software, Supplies, Professional Services, Other
- Budget amounts: $50K-$2M range
- Actuals: most within 5-10% of budget, but 3-4 departments significantly over/under
- Variance formulas: E = D - C, F = (D - C) / C

**Seeded Problems (6):**
1. **3 inconsistent variance formulas** — rows 8, 17, and 25 should have hardcoded values instead of formulas in column E (like someone manually typed a number). InconsistentFormulasAuditor / FormulaConsistencyChecker will catch these
2. **No number formatting** — Budget and Actual columns show raw numbers without commas (NumberFormatStandardizer will fix)
3. **No negative highlighting** — negative variances look the same as positive (HighlightNegativesRed will fix)
4. **2 text-stored numbers** — cells C12 and D12 are text, not numbers (ConvertTextToNumbers will fix)
5. **No conditional formatting at all** — completely plain (perfect for FluxAnalysis to add threshold highlighting)
6. **Header row is not frozen** — no freeze panes applied

---

### Sheet 5: "Vendor Master"

**Purpose:** Master vendor directory — classic dirty data scenario.

**Columns:** A=Vendor ID | B=Company Name | C=Contact Name | D=Address | E=City | F=State | G=Phone | H=Email

**Data — ~50 rows:**
- Use the same vendors from the AP Aging sheet plus 20-25 more
- Realistic addresses across different US states
- Phone numbers in mixed formats
- Email addresses at company domains

**Seeded Problems (9):**
1. **Leading/trailing spaces** — 6 cells across Company Name and Contact Name columns have invisible leading or trailing spaces. RemoveLeadingTrailingSpaces will fix
2. **Non-breaking spaces** — 3 cells have char(160) instead of regular spaces. UniversalWhitespaceCleaner will fix
3. **Non-printable characters** — 2 cells in Address column have a hidden control character (like char(7) or char(0)). NonPrintableCharStripper will fix
4. **Mixed case** — Company names inconsistently cased: "VERTEX OFFICE SOLUTIONS" (row 5), "vertex office solutions" (row 30), "Vertex Office Solutions" (row 42). TextCaseStandardizer will normalize
5. **Duplicate vendors** — 3 vendors appear twice with slightly different names: "DataStream Software" vs "Datastream Software Inc.", "NorthStar Insurance" vs "North Star Insurance Corp", "BrightPath Marketing" vs "Bright Path Marketing LLC". HighlightDuplicateRows / ExactDuplicateFinder will flag
6. **5 blank Vendor IDs** — GenerateUniqueCustomerIDs will fill
7. **Phone format inconsistency** — mix of (555) 123-4567, 555-123-4567, 5551234567, 555.123.4567
8. **2 cells with double spaces** inside text ("John  Smith" with two spaces)
9. **3 embedded hyperlinks** in the Email column that aren't needed — PhantomHyperlinkPurger will clean

---

### Sheet 6: "Monthly P&L"

**Purpose:** Income statement summary by month — the executive sheet that ties everything together.

**Columns:** A=Line Item | B=January | C=February | D=March | E=Q1 Total

**Data — ~35 rows organized as:**
- **Revenue section** (5-6 lines): Product Revenue, Service Revenue, Licensing Revenue, Other Income, Total Revenue
- **COGS section** (4-5 lines): Direct Materials, Direct Labor, Manufacturing Overhead, Freight, Total COGS
- **Gross Profit** (formula: Total Revenue - Total COGS)
- **Operating Expenses section** (8-10 lines): Salaries & Benefits, Rent & Facilities, Marketing & Advertising, Technology & Software, Travel & Entertainment, Professional Fees, Depreciation, Insurance, Office & Admin, Total Operating Expenses
- **Operating Income** (formula: Gross Profit - Total OpEx)
- **Other Income/Expense** (2-3 lines): Interest Expense, Other, Total Other
- **Net Income Before Tax**
- **Tax Provision**
- **Net Income**
- Revenue should be in the $2M-$5M/month range
- Show realistic month-over-month growth (Jan < Feb < Mar slightly)
- Q1 Total column = sum of Jan + Feb + Mar

**Seeded Problems (5):**
1. **3 floating-point tails** — cells B8, C15, and D22 should have values like 2450000.0000000003, 187500.0000000001, 95000.00000000002. FixFloatingPointTails will clean these
2. **2 formula cells mixed with hardcoded** — the Q1 Total column should be formulas (=SUM) for most rows, but rows 12 and 20 should be hardcoded values instead of formulas. FormulaToValueHardcoder can demo here
3. **No formatting** — no dollar signs, no commas, no parentheses for negatives (NumberFormatStandardizer + FinancialNumberFormattingSuite will format)
4. **Interest Expense should be negative** but displayed as a positive number without parentheses (HighlightNegativesRed will flag)
5. **No print headers/footers set** — PrintHeaderFooterStandardizer will add "DemoKit Industries — Confidential" headers

---

## DATA QUALITY REQUIREMENTS

1. **All numbers must be realistic** — nothing that looks obviously fake. Study real Fortune 500 financials for scale
2. **The Trial Balance must actually balance** (except for the intentional $0.03 discrepancy) — debits genuinely equal credits across all accounts minus $0.03
3. **Journal entries must balance** (except the 2 intentional imbalances)
4. **Budget vs Actual variances must make logical sense** — Marketing overspent on a campaign, IT underspent because a project was delayed, etc.
5. **The Monthly P&L must flow logically** — Revenue > COGS > Gross Profit > OpEx > Net Income. Margins should be realistic (40-60% gross margin, 10-20% net margin)
6. **Vendor names must sound like real companies** — not joke names
7. **All dates must be in Q1 2026** (January 1 - March 31, 2026)

---

## FORMATTING REQUIREMENTS

1. **Keep formatting minimal and generic** — this is intentional. The branding tool transforms it live
2. **Bold header rows** on every sheet — but no fill colors
3. **Column widths** should be auto-fitted so all data is visible
4. **Wrap text** on columns with longer content (Memo, Address, Description)
5. **Total/subtotal rows** should be bold with a top border (single line) — standard accounting format
6. **Right-align** all number columns
7. **Left-align** all text columns
8. **Add a thin border grid** on all data (light gray borders) — keeps it clean but plain
9. **Sheet tab colors**: do NOT set any tab colors (the branding tool will do this)
10. **No filters applied** to any sheet (ResetAllFilters demo)
11. **No freeze panes** on any sheet (FreezeTopRowAllSheets demo)

---

## WHAT MAKES THIS WORLD-CLASS

- A Finance professional should look at this and think "this looks exactly like a file I'd get from my team"
- The data tells a coherent story — DemoKit Industries had a solid Q1 with revenue growing each month, a few departments overspent, and the AP team has some cleanup to do
- The seeded problems are **subtle** — they look like genuine mistakes, not planted errors
- The structure is clean and professional even without branding
- Every sheet connects — the vendors in AP Aging match the Vendor Master, the GL accounts in Journal Entries match the Trial Balance, the P&L ties to the budget categories

---

## FINAL CHECKLIST BEFORE DELIVERING

- [ ] 6 sheets total, named exactly as specified
- [ ] ~80 rows Trial Balance, ~100 rows AP Aging, ~60 rows GL JEs, ~30 rows Budget vs Actual, ~50 rows Vendor Master, ~35 rows Monthly P&L
- [ ] All seeded problems present (7+8+7+6+9+5 = 42 total problems)
- [ ] Trial Balance debits/credits off by exactly $0.03
- [ ] 2 unbalanced journal entries (off by $500 and $0.50)
- [ ] 3 duplicate invoices in AP Aging
- [ ] Mixed date formats in AP Aging
- [ ] Merged cells in GL JEs column A
- [ ] Inconsistent formulas in Budget vs Actual
- [ ] Floating-point tails in Monthly P&L
- [ ] Text-stored numbers across 3 sheets
- [ ] External link formulas in Trial Balance
- [ ] Dirty vendor data (spaces, case, duplicates, non-printables)
- [ ] Generic formatting (bold headers, no colors, no branding)
- [ ] All data is realistic and internally consistent
- [ ] No freeze panes on any sheet
- [ ] Named Range "PriorYearTB" pointing to non-existent sheet

**Build this file and make it perfect. This is going in front of 2,000 people and the CEO.**
