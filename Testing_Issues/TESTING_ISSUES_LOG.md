# KBT P&L Toolkit v2.1.0 — Testing Issues Log

> **Project:** Keystone BenefitTech P&L Demo
> **Tester:** Connor Atlee
> **Repo:** AP_ProjectVer2_Testing
> **Branch:** `claude/review-code-testing-s4dsQ`
> **Last Updated:** 2026-03-02

---

## How to Use This Log

Every testing issue — whether it turns out to be a real bug, a false positive,
or a documentation problem — gets logged here. This gives the next Claude account
(or future tester) a full picture of what has already been investigated so nothing
gets re-investigated from scratch.

**Status values:**
- `RESOLVED — NOT A BUG` — Investigated; the code is correct; test or documentation was wrong
- `RESOLVED — FIXED` — Real bug; fix was applied and pushed
- `OPEN` — Real issue; not yet fixed
- `DEFERRED` — Known issue; accepted as-is for now

---

## ISSUE-T1-001 — T1.06: Outdated module names in TEST_PLAN.md

| Field | Details |
|---|---|
| **Test ID** | T1.06 |
| **Date Found** | 2026-03-01 |
| **Found By** | Connor Atlee (manual review during testing) |
| **Status** | RESOLVED — NOT A BUG |

### What Happened
The original TEST_PLAN.md T1.06 import test listed these module names in the
`python -c "import ..."` command:
- `pnl_cleaner`
- `pnl_loader`
- `pnl_allocator`
- `pnl_variance`
- `pnl_report`

None of these files exist in the `python/` folder. They are old names from an
earlier draft of the toolkit that were renamed during development.

### Actual File Names (correct)
The 14 real Python files are:
`build_charts.py`, `pnl_allocation_simulator.py`, `pnl_ap_matcher.py`,
`pnl_cli.py`, `pnl_config.py`, `pnl_dashboard.py`, `pnl_email_report.py`,
`pnl_forecast.py`, `pnl_monte_carlo.py`, `pnl_month_end.py`, `pnl_runner.py`,
`pnl_snapshot.py`, `pnl_tests.py`, `redesign_pl_model.py`

### Resolution
TEST_PLAN.md was updated with the correct file names. The Python scripts
themselves are complete and functional. T1.06 result: **PASS**.

### Impact on Downstream Tests
None. The scripts are fine. Only the test plan documentation was outdated.

---

## ISSUE-T1-002 — T1.05: Source Excel file not found warning

| Field | Details |
|---|---|
| **Test ID** | T1.05 |
| **Date Found** | 2026-03-01 |
| **Found By** | Connor Atlee |
| **Status** | RESOLVED — NOT A BUG |

### What Happened
Running `python pnl_config.py` printed the full config summary successfully
(all shares summed to 1.00, version 2.1.0 confirmed), but also printed a warning:

```
⚠ Source file not found: ExcelDemoFile_adv.xlsm
```

### Root Cause
`pnl_config.py` checks for the source Excel file at the path defined in
`SOURCE_FILE`. The test was run from the repo folder, not the folder containing
the Excel workbook. Additionally, the Excel file had been renamed during the
redesign session (it now ends in `_TEST.xlsm` during testing).

### Resolution
This is expected behavior. `pnl_config.py` gracefully reports the missing file
and continues — it does not crash. The warning is cosmetic. T1.05 result: **PASS**.

### Impact on Downstream Tests
None. The config module loads correctly. Any downstream test that actually
needs to read the Excel file should be run from the folder where the workbook lives.

---

## ISSUE-T1-003 — T1.07: UTF-8 scan false positive (most important issue this session)

| Field | Details |
|---|---|
| **Test ID** | T1.07 |
| **Date Found** | 2026-03-01 |
| **Found By** | Connor Atlee (scan ran on all 14 files; all 14 flagged) |
| **Investigated By** | Claude (2026-03-02) |
| **Status** | RESOLVED — NOT A BUG |

### What Happened
T1.07 ("Python UTF-8 clean") ran a scan for non-ASCII bytes across all 14
Python files. All 14 files were flagged as containing non-ASCII bytes. The tester
ran the scan using two different methods and got the same result both times.

The original test plan pass criteria said: "Zero mojibake codepoints"

The failing files:
```
build_charts.py, pnl_allocation_simulator.py, pnl_ap_matcher.py,
pnl_cli.py, pnl_config.py, pnl_dashboard.py, pnl_email_report.py,
pnl_forecast.py, pnl_monte_carlo.py, pnl_month_end.py, pnl_runner.py,
pnl_snapshot.py, pnl_tests.py, redesign_pl_model.py
```

### Investigation (2026-03-02)

All 14 files were scanned using Python's built-in `open(...).decode('utf-8')`.
**Result: All 14 files decoded as valid UTF-8 without a single UnicodeDecodeError.**

The non-ASCII characters present in the files are all intentional and correctly
encoded. Full inventory of what was found:

| Character | Unicode | Type | Used In |
|---|---|---|---|
| `—` | U+2014 | Em dash | Titles, section headers, comments (all files) |
| `→` | U+2192 | Right arrow | Direction/flow indicators |
| `↑` `↓` | U+2191/U+2193 | Up/down arrows | Change direction (allocation_simulator) |
| `↔` | U+2194 | Left-right arrow | Matched pairs (ap_matcher) |
| `✓` | U+2713 | Check mark | Pass indicators (config, cli, email_report) |
| `✗` | U+2717 | Cross mark | Fail indicators (config, cli) |
| `⚠` | U+26A0 | Warning sign | Warning messages (multiple files) |
| `▲` `▼` | U+25B2/U+25BC | Up/down triangles | Revenue direction (email_report) |
| `▶` | U+25B6 | Right triangle | Section markers (ap_matcher, cli) |
| `─` | U+2500 | Box drawing (horiz.) | Visual section dividers (ap_matcher, forecast) |
| `═` `║` `╔` `╗` `╚` `╝` | U+2550 etc. | Box drawing | Banner borders (pnl_runner) |
| `Δ` | U+0394 | Greek capital delta | Change/delta labels (allocation_simulator) |
| `α` | U+03B1 | Greek alpha | Confidence level (monte_carlo) |
| `±` | U+00B1 | Plus-minus sign | Uncertainty ranges (monte_carlo) |
| `×` | U+00D7 | Multiplication sign | "Department × Product" labels (dashboard, snapshot) |
| `📊` | U+1F4CA | Bar chart emoji | Streamlit page icon (dashboard) |
| `█` `◇` `○` | Various | Block/shapes | Text-based bar charts (monte_carlo) |

**None of the above are mojibake.** Mojibake would look like: `Ã©` for é,
`â€"` for —, `Ã¢` for â. Those sequences do NOT appear anywhere in any file.

### Why the Scan Showed a Failure
The scan tool used checked for the presence of *any* non-ASCII byte. This is
too strict. A correctly written UTF-8 scan should only flag garbled Latin-1
misread sequences (mojibake), not legitimate Unicode characters.

### Resolution
- All 14 Python files: confirmed clean, no changes needed
- TEST_PLAN.md updated: T1.07 and T4.01 pass criteria clarified to specify
  "mojibake patterns only" not "any non-ASCII byte"
- T1.07 result: **PASS**

### What Mojibake Would Actually Look Like (for future reference)
If a file had real encoding corruption, you would see strings like:
- `â€"` instead of `—` (em dash misread as Latin-1)
- `Ã©` instead of `é` (e-acute misread as Latin-1)
- `Ã¢` instead of `â` (a-circumflex misread)
- `â€œ` / `â€` instead of `"` / `"` (smart quotes misread)

None of these patterns exist in any file.

### Impact on Downstream Tests
**T1.08 and T2 are unblocked.** T1.07 is not a real failure and does not
affect any downstream test. Proceed with testing from T1.08.

---

## T1 Test Summary Table

| Test ID | Result | Notes |
|---|---|---|
| T1.01 | PASS | VBA compiles with zero errors |
| T1.02 | PASS | All 32 modules present in Project Explorer |
| T1.03 | PASS | Option Explicit in all modules |
| T1.04 | PASS | ?APP_VERSION = "2.1.0" |
| T1.05 | PASS | pnl_config.py loads cleanly; source file warning is cosmetic |
| T1.06 | PASS | All 14 Python scripts import successfully |
| T1.07 | PASS | All 14 files are valid UTF-8; non-ASCII chars are intentional Unicode |
| T1.08 | NOT YET RUN | Next: `pip install -r requirements.txt` |

---

## Next Steps for Next Account / Tester

1. Resume at **T1.08**: run `pip install -r requirements.txt` from the `python/` folder
2. Verify all packages install without errors
3. Proceed through **T2** (Foundation Issues) — see qa/TEST_PLAN.md for procedures
4. Log all results in `qa/TEST_RESULTS.md` (create it if it doesn't exist)
5. Add any new issues to this file using the same format above

---

*Log maintained by Claude on behalf of Connor Atlee — iPipeline Finance & Accounting*
