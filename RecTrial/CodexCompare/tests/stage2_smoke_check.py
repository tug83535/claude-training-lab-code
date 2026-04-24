#!/usr/bin/env python3
"""Stage smoke checks for universal VBA and Python artifacts."""

from __future__ import annotations

from pathlib import Path
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]

REQUIRED_VBA_FILES = {
    "vba/universal/modUTL_Core.bas": [
        "Public Function UTL_GetTargetSheets",
        "Public Sub UTL_LogAction",
        "Public Function UTL_DetectDataRange",
    ],
    "vba/universal/modUTL_DataSanitizer.bas": [
        "Public Sub RunFullSanitize",
        "Public Sub PreviewSanitizeChanges",
    ],
    "vba/universal/modUTL_CommandCenter.bas": [
        "Public Sub BuildCommandCenter",
        "Public Sub Run_CommandCenter_Sanitize",
    ],
    "vba/universal/modUTL_CompareConsolidate.bas": [
        "Public Sub CompareActiveSheetToSheet",
        "Public Sub ConsolidateVisibleSheetsByHeader",
    ],
    "vba/universal/modUTL_Intelligence.bas": [
        "Public Sub MaterialityClassifierActiveSheet",
        "Public Sub GenerateExceptionNarrativesActiveSheet",
    ],
    "vba/universal/modUTL_OutputPack.bas": [
        "Public Sub BuildExecutiveOnePagerFromActiveSheet",
        "Public Sub ExportExecutivePackPDF",
    ],
}


REQUIRED_DEMO_VBA_FILES = {
    "vba/demo/modDemo_Config.bas": [
        "Public Function DemoWorkbookReady",
        "Public Sub DemoValidateWorkbookOrStop",
    ],
    "vba/demo/modDemo_AuditTrail.bas": [
        "Public Sub DemoLog",
        "Public Function DemoGetOrCreateAuditSheet",
    ],
    "vba/demo/modDemo_ReconciliationEngine.bas": [
        "Public Sub RunDemoReconciliation",
        "Private Sub RunRevenueTieOut",
    ],
    "vba/demo/modDemo_VarianceNarrative.bas": [
        "Public Sub GenerateDemoVarianceNarrative",
        "Private Function BuildVarianceNarrative",
    ],
    "vba/demo/modDemo_ExecBriefPack.bas": [
        "Public Sub BuildDemoExecutiveBriefPack",
        "Private Sub BuildKpiSection",
    ],
    "vba/demo/modDemo_WhatIfScenario.bas": [
        "Public Sub RunDemoWhatIfScenarios",
        "Private Function EvaluateScenario",
    ],
    "vba/demo/modDemo_CommandCenter.bas": [
        "Public Sub BuildDemoCommandCenter",
        "Private Sub AddDemoButton",
    ],
}


REQUIRED_SQL_FILES = {
    "sql/universal/template_gl_extract.sql": ["SELECT", "FROM finance.gl_transactions"],
    "sql/universal/template_revenue_extract.sql": ["SELECT", "FROM finance.revenue_transactions"],
    "sql/demo/demo_pnl_reconciliation_view.sql": ["CREATE OR ALTER VIEW", "demo.vw_pnl_reconciliation_source"],
    "sql/demo/demo_variance_fact.sql": ["WITH base AS", "delta_value"],
}

REQUIRED_PYTHON_SCRIPTS = [
    "python/universal/profile_workbook.py",
    "python/universal/sanitize_dataset.py",
    "python/universal/compare_workbooks.py",
    "python/universal/build_exec_summary.py",
    "python/demo/pnl_data_extract.py",
    "python/demo/variance_classifier.py",
    "python/demo/scenario_runner.py",
    "python/demo/export_brief_package.py",
]

SAMPLE_FILES = [
    ROOT / "samples/ExcelDemoFile_adv.xlsm",
    ROOT / "samples/Sample_Quarterly_ReportV2.xlsm",
]

CATALOG_PATH = ROOT / "guides/universal-tool-catalog.md"
MINIMUM_TOOL_TARGET = 155
COPILOT_GUIDE_PATH = ROOT / "guides/copilot-prompt-guide.md"
UNIVERSAL_GUIDE_PATH = ROOT / "guides/universal-toolkit-user-guide.md"
DEMO_GUIDE_PATH = ROOT / "guides/demo-walkthrough-guide.md"
BRAND_REF_PATH = ROOT / "guides/brand-styling-reference.md"
TROUBLESHOOTING_PATH = ROOT / "guides/troubleshooting-reference.md"
RELEASE_CHECKLIST_PATH = ROOT / "guides/release-readiness-checklist.md"
ARCHITECTURE_GUIDE_PATH = ROOT / "guides/architecture-overview.md"
GIT_PUSH_GUIDE_PATH = ROOT / "guides/git-branch-push-quickstart.md"
CLAUDE_HANDOFF_GUIDE_PATH = ROOT / "guides/claude-handoff-deep-analysis.md"
CLAUDE_PROMPT_GUIDE_PATH = ROOT / "guides/claude-review-prompt.md"
PROJECT_TODO_PATH = ROOT / "PROJECT_TODO.md"
CHANGELOG_PATH = ROOT / "CHANGELOG.md"
VIDEO_1_PATH = ROOT / "videos/video-1-executive-hook.md"
VIDEO_2_PATH = ROOT / "videos/video-2-demo-workbook-deep-dive.md"
VIDEO_3_PATH = ROOT / "videos/video-3-universal-toolkit-in-action.md"
VIDEO_4_PATH = ROOT / "videos/video-4-python-sql-integration.md"
VIDEO_5_PATH = ROOT / "videos/video-5-copilot-adaptation-lab.md"
README_PATH = ROOT / "README.md"
SMOKE_SCRIPT_PATH = ROOT / "scripts/run_stage_smoke.sh"
SMOKE_WORKFLOW_PATH = ROOT / ".github/workflows/smoke-check.yml"
MAKEFILE_PATH = ROOT / "Makefile"
CONTRIBUTING_PATH = ROOT / "CONTRIBUTING.md"
BOOTSTRAP_SCRIPT_PATH = ROOT / "scripts/bootstrap_demo_workspace.py"
CODE_INVENTORY_SCRIPT_PATH = ROOT / "scripts/update_code_inventory.py"
CODE_INVENTORY_PATH = ROOT / "CODE_INVENTORY.md"
PY_UNIT_TEST_PATH = ROOT / "tests/test_python_utilities.py"

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    sys.exit(1)


def check_vba_files() -> None:
    for rel, markers in REQUIRED_VBA_FILES.items():
        path = ROOT / rel
        if not path.exists():
            fail(f"Missing required file: {rel}")

        content = path.read_text(encoding="utf-8")
        for marker in markers:
            if marker not in content:
                fail(f"Marker '{marker}' not found in {rel}")

        if "Option Explicit" not in content:
            fail(f"Option Explicit missing in {rel}")


def check_demo_vba_files() -> None:
    for rel, markers in REQUIRED_DEMO_VBA_FILES.items():
        path = ROOT / rel
        if not path.exists():
            fail(f"Missing required demo VBA file: {rel}")

        content = path.read_text(encoding="utf-8")
        for marker in markers:
            if marker not in content:
                fail(f"Marker '{marker}' not found in {rel}")

        if "Option Explicit" not in content:
            fail(f"Option Explicit missing in {rel}")


def check_sql_files() -> None:
    for rel, markers in REQUIRED_SQL_FILES.items():
        path = ROOT / rel
        if not path.exists():
            fail(f"Missing required SQL file: {rel}")

        content = path.read_text(encoding="utf-8")
        for marker in markers:
            if marker not in content:
                fail(f"SQL file missing marker '{marker}': {rel}")


def check_python_scripts() -> None:
    for rel in REQUIRED_PYTHON_SCRIPTS:
        path = ROOT / rel
        if not path.exists():
            fail(f"Missing required Python script: {rel}")

        result = subprocess.run(
            [sys.executable, str(path), "--help"],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            fail(f"Help command failed for {rel}: {result.stderr.strip()}")


def workbook_sheet_count(path: Path) -> int:
    with zipfile.ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets = workbook.find("main:sheets", NS)
        if sheets is None:
            return 0
        return len(list(sheets))


def has_vba_project(path: Path) -> bool:
    with zipfile.ZipFile(path) as zf:
        return "xl/vbaProject.bin" in zf.namelist()


def check_samples() -> None:
    for path in SAMPLE_FILES:
        if not path.exists():
            fail(f"Sample file not found: {path}")

        count = workbook_sheet_count(path)
        if count == 0:
            fail(f"Workbook has no sheets: {path.name}")

        if not has_vba_project(path):
            fail(f"Workbook missing VBA project: {path.name}")

        print(f"OK: {path.name} -> {count} sheet(s), VBA project present")


def check_no_sample_mutation_hint() -> None:
    for path in SAMPLE_FILES:
        if path.stat().st_size < 10_000:
            fail(f"Unexpectedly small sample workbook (possible corruption): {path.name}")


def check_tool_catalog() -> None:
    if not CATALOG_PATH.exists():
        fail("Missing universal tool catalog markdown file")

    content = CATALOG_PATH.read_text(encoding="utf-8")
    tool_numbers = re.findall(r"^\d+\.\s", content, flags=re.MULTILINE)
    count = len(tool_numbers)

    if count < MINIMUM_TOOL_TARGET:
        fail(f"Tool catalog contains {count} tools; expected at least {MINIMUM_TOOL_TARGET}")

    print(f"OK: universal tool catalog contains {count} tools")



def check_copilot_guide() -> None:
    if not COPILOT_GUIDE_PATH.exists():
        fail("Missing guides/copilot-prompt-guide.md")

    content = COPILOT_GUIDE_PATH.read_text(encoding="utf-8")
    required_sections = [
        "## 1) Who this guide is for",
        "## 4) Starter prompt template (copy/paste)",
        "## 6) Worked examples by feature",
        "## 8) Troubleshooting when CoPilot gets it wrong",
    ]

    for section in required_sections:
        if section not in content:
            fail(f"CoPilot guide missing required section: {section}")


def check_additional_guides() -> None:
    required_guides = {
        UNIVERSAL_GUIDE_PATH: ["# Universal Toolkit User Guide", "## 4) Quickstart (15 minutes)", "## 7) Troubleshooting"],
        DEMO_GUIDE_PATH: ["# Demo Walkthrough Guide (P&L Workbook)", "## 2) Demo sequence", "## 3) Expected outputs"],
        BRAND_REF_PATH: ["# Brand Styling Reference (Operational)", "## 1) Color palette", "## 6) Banned items"],
        TROUBLESHOOTING_PATH: ["# Troubleshooting Reference", "## 1) Macro does not run", "## 6) Escalation template"],
        RELEASE_CHECKLIST_PATH: ["# Release Readiness Checklist", "## 1) Workbook readiness", "## 7) Final sign-off"],
        ARCHITECTURE_GUIDE_PATH: ["# Architecture Overview", "## 1) Two-prong model", "## 5) Validation strategy"],
        GIT_PUSH_GUIDE_PATH: ["# Git Branch + Push Quickstart (Beginner-Safe)", "## 3) Push your branch", "## 5) Common errors and fixes"],
        CLAUDE_HANDOFF_GUIDE_PATH: ["# Claude Code Handoff — Deep Analysis Package", "## 4) What to Compare Against Claude-built Version", "## 7) Suggested Claude Review Deliverables"],
        CLAUDE_PROMPT_GUIDE_PATH: ["# Claude Code Comparison Prompt (Copy/Paste)", "## Required outputs", "## Deliverable style"],
    }

    for guide_path, sections in required_guides.items():
        if not guide_path.exists():
            fail(f"Missing guide: {guide_path}")

        content = guide_path.read_text(encoding="utf-8")
        for section in sections:
            if section not in content:
                fail(f"Guide missing required section '{section}': {guide_path}")


def check_video_scripts() -> None:
    required_videos = {
        VIDEO_1_PATH: ["## Timestamped Outline", "## Full Narration Script", "## Closing CTA"],
        VIDEO_2_PATH: ["## Timestamped Outline", "## Full Narration Script", "## Closing CTA"],
        VIDEO_3_PATH: ["## Timestamped Outline", "## Full Narration Script", "## Closing CTA"],
        VIDEO_4_PATH: ["## Timestamped Outline", "## Full Narration Script", "## Closing CTA"],
        VIDEO_5_PATH: ["## Timestamped Outline", "## Full Narration Script", "## Closing CTA"],
    }

    for video_path, sections in required_videos.items():
        if not video_path.exists():
            fail(f"Missing video script: {video_path}")

        content = video_path.read_text(encoding="utf-8")
        for section in sections:
            if section not in content:
                fail(f"Video script missing required section '{section}': {video_path}")


def check_readme_status_block() -> None:
    if not README_PATH.exists():
        fail("Missing README.md")

    content = README_PATH.read_text(encoding="utf-8")
    required = [
        "## Current Build Status (Stage Progress)",
        "Video scripts 1–5 delivered in `videos/`.",
        "tests/stage2_smoke_check.py",
    ]

    for marker in required:
        if marker not in content:
            fail(f"README missing status marker: {marker}")


def check_smoke_automation_files() -> None:
    if not SMOKE_SCRIPT_PATH.exists():
        fail("Missing scripts/run_stage_smoke.sh")
    if not SMOKE_WORKFLOW_PATH.exists():
        fail("Missing .github/workflows/smoke-check.yml")

    script_content = SMOKE_SCRIPT_PATH.read_text(encoding="utf-8")
    workflow_content = SMOKE_WORKFLOW_PATH.read_text(encoding="utf-8")

    script_markers = ["python tests/stage2_smoke_check.py", "python -m py_compile", "python -m unittest tests/test_python_utilities.py"]
    workflow_markers = ["name: smoke-check", "scripts/run_stage_smoke.sh"]
    makefile_markers = ["smoke:", "unit:", "py-compile:"]

    for marker in script_markers:
        if marker not in script_content:
            fail(f"Smoke script missing marker: {marker}")

    for marker in workflow_markers:
        if marker not in workflow_content:
            fail(f"Workflow missing marker: {marker}")

    if not PY_UNIT_TEST_PATH.exists():
        fail("Missing tests/test_python_utilities.py")

    if not MAKEFILE_PATH.exists():
        fail("Missing Makefile")
    makefile_content = MAKEFILE_PATH.read_text(encoding="utf-8")
    for marker in makefile_markers:
        if marker not in makefile_content:
            fail(f"Makefile missing marker: {marker}")

    if not CONTRIBUTING_PATH.exists():
        fail("Missing CONTRIBUTING.md")

    if not BOOTSTRAP_SCRIPT_PATH.exists():
        fail("Missing scripts/bootstrap_demo_workspace.py")


def check_bootstrap_script() -> None:
    result = subprocess.run(
        [sys.executable, str(BOOTSTRAP_SCRIPT_PATH), "--help"],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        fail(f"Bootstrap script help failed: {result.stderr.strip()}")


def check_code_inventory() -> None:
    if not CODE_INVENTORY_SCRIPT_PATH.exists():
        fail("Missing scripts/update_code_inventory.py")
    if not CODE_INVENTORY_PATH.exists():
        fail("Missing CODE_INVENTORY.md")

    before = CODE_INVENTORY_PATH.read_text(encoding="utf-8")

    result = subprocess.run(
        [sys.executable, str(CODE_INVENTORY_SCRIPT_PATH)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        fail(f"Code inventory generator failed: {result.stderr.strip()}")

    content = CODE_INVENTORY_PATH.read_text(encoding="utf-8")
    required_markers = [
        "# Code Inventory",
        "python/universal/profile_workbook.py",
        "python/demo/pnl_data_extract.py",
        "vba/universal/modUTL_Core.bas",
        "vba/demo/modDemo_ReconciliationEngine.bas",
    ]
    for marker in required_markers:
        if marker not in content:
            fail(f"Code inventory missing marker: {marker}")

    if before != content:
        fail(
            "CODE_INVENTORY.md is out of date. "
            "Run 'python scripts/update_code_inventory.py' and commit the regenerated file."
        )


def check_changelog() -> None:
    if not CHANGELOG_PATH.exists():
        fail("Missing CHANGELOG.md")

    content = CHANGELOG_PATH.read_text(encoding="utf-8")
    required_markers = [
        "# Changelog",
        "## 2026-04-20",
        "Sample files in `samples/` are treated as read-only input assets.",
    ]

    for marker in required_markers:
        if marker not in content:
            fail(f"Changelog missing marker: {marker}")


def check_project_todo() -> None:
    if not PROJECT_TODO_PATH.exists():
        fail("Missing PROJECT_TODO.md")

    content = PROJECT_TODO_PATH.read_text(encoding="utf-8")
    required_markers = [
        "# Project To-Do (Execution Backlog)",
        "## Priority 0 — Immediate Workflow Stability",
        "## Priority 4 — Comparative Analysis Follow-up",
    ]
    for marker in required_markers:
        if marker not in content:
            fail(f"PROJECT_TODO missing marker: {marker}")

def main() -> None:
    check_vba_files()
    check_demo_vba_files()
    check_sql_files()
    check_python_scripts()
    check_samples()
    check_no_sample_mutation_hint()
    check_tool_catalog()
    check_copilot_guide()
    check_additional_guides()
    check_video_scripts()
    check_readme_status_block()
    check_smoke_automation_files()
    check_bootstrap_script()
    check_code_inventory()
    check_changelog()
    check_project_todo()
    print("PASS: Stage smoke checks completed")


if __name__ == "__main__":
    main()
