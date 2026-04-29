# AI Agent Instructions — RecTrial (Finance Automation Demo)

## Project Overview
This is a Finance & Accounting **adoption-grade automation package + demo videos** for iPipeline (SaaS, insurance industry). It combines Excel VBA, Python, and SQL. **Near-term audience: 50–150 coworkers** in Finance, Accounting, and adjacent operations. Broader rollout and any CFO/CEO showcase are deferred. Coworkers are expected to actually use the tools on their own files.

Two-prong architecture:
- **Prong 1 — Universal Toolkit:** File-agnostic VBA macros + Python utilities. No hardcoded sheet names. Works on any workbook.
- **Prong 2 — File-Specific Demo:** Rich features tied to `DemoFile/ExcelDemoFile_adv.xlsm` (16-sheet P&L workbook).

## Key Docs (Read Before Changing Anything)
- [CodexCompare/CONSTRAINTS.md](CodexCompare/CONSTRAINTS.md) — Banned features + threshold to justify any new feature
- [CodexCompare/BRAND.md](CodexCompare/BRAND.md) — Colors, fonts, formatting rules (non-negotiable)
- [CodexCompare/CONTEXT.md](CodexCompare/CONTEXT.md) — Audience, company, success criteria
- [CodexCompare/PLAN.md](CodexCompare/PLAN.md) — Architecture decisions + file inventory
- [CodexCompare/PROJECT_TODO.md](CodexCompare/PROJECT_TODO.md) — Priority queue
- [Guide/MASTER_RECORDING_GUIDE.md](Guide/MASTER_RECORDING_GUIDE.md) — Recording workflow (OBS + Director macro)

## Folder Map
| Folder | Purpose |
|--------|---------|
| `DemoVBA/` | 39 active VBA modules (v2.1) — the working project |
| `DemoPython/` | 13 active Python scripts (P&L workflows, forecasting, dashboard) |
| `DemoFile/` | Primary Excel workbook (do not modify directly) |
| `CodexCompare/` | Reference architecture — parallel build from scratch |
| `CodexCompare/vba/` | 13 reference VBA modules (6 universal, 7 demo-specific) |
| `CodexCompare/python/` | 8 reference Python scripts (4 universal, 4 demo-specific) |
| `CodexCompare/guides/` | 11 training/handoff docs |
| `CodexCompare/tests/` | Smoke checks + unit tests |
| `AudioClips/VideoN/` | MP3 narration files for Director macro |
| `VideoScripts/` | Narration markdown scripts |
| `Guide/` | Recording orchestration docs |

## Build & Test Commands (run from `CodexCompare/`)
```bash
make smoke       # Quick validation (stage2_smoke_check.py)
make unit        # pytest on test_python_utilities.py
make py-compile  # Validate syntax on all .py files
make check       # Full check: py-compile + unit + inventory
```

## Naming Conventions
- **VBA modules:** `mod<Feature>_v2.1.bas` (active) or `modUTL_<Feature>.bas` (universal toolkit reference)
- **Python:** `pnl_<feature>.py` (DemoPython) or `<purpose>_<workbook>.py` (CodexCompare)
- **Output sheets:** `UTL_*` (universal), `VER_*` (version snapshot), `BKP_*` (backup)
- **Command Center action IDs:** Numeric (1–65+), must be registered in `modCommandCenter` or `modFormBuilder`

## Code Quality Rules (Always Apply)
1. **Non-destructive:** Write to new sheets, never overwrite source data
2. **Logging:** Every action → `VBA_AuditLog` (timestamp, user, module, procedure, status)
3. **Defensive guards:** Check for merged cells, blank rows, error formulas before processing
4. **Prong 1 test:** Universal tools must work on BOTH sample files (ExcelDemoFile + SampleFile)
5. **Progress feedback:** Long operations use `modProgressBar`; every feature shows a success/fail message
6. **No banned patterns:** See [CONSTRAINTS.md](CodexCompare/CONSTRAINTS.md) — no native Excel sort/filter/pivot, no OneDrive features

## Branding (Non-Negotiable)
| Element | Spec |
|---------|------|
| Font | Arial only |
| Primary | `#0B4779` (iPipeline Blue) |
| Secondary | `#112E51` (Navy) |
| Accent | `#4B9BCB` (Innovation Blue) |
| Pass/Favorable | `#BFF18C` (Lime Green) |
| Info | `#2BCCD3` (Aqua) |
| Currency | `$#,##0` |
| Percentages | `0.0%` |
| Dates | `mmm-yy` or `yyyy-mm-dd` |

## Python Dependencies (DemoPython)
Install via: `pip install -r DemoPython/requirements.txt`
Key packages: pandas, openpyxl, statsmodels, scikit-learn, plotly, streamlit
