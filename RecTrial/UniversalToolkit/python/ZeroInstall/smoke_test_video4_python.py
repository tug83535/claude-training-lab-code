# Finance Automation Toolkit v1.0 — iPipeline
# smoke_test_video4_python.py — runs each V4 script in sample mode and verifies outputs exist
#
# Runs:
#   1. revenue_leakage_finder.py --sample
#   2. data_contract_checker.py --sample
#   3. exception_triage_engine.py --sample
#   4. control_evidence_pack.py --sample
#   5. workbook_dependency_scanner.py --sample
#
# For each run, verifies:
#   - Script exited with return code 0
#   - Expected output files exist in the timestamped output folder
#
# Usage:
#   python smoke_test_video4_python.py
#
# A PASS/FAIL line is printed for each script. Exit code = 0 if all pass, 1 if any fail.

import subprocess
import sys
import time
from pathlib import Path

_ROOT = Path(__file__).parent

_TESTS = [
    {
        "script":           "revenue_leakage_finder.py",
        "args":             ["--sample"],
        "expected_outputs": ["leakage_report.html", "exceptions_ranked.csv",
                             "arr_waterfall.html", "run_log.json", "run_summary.txt"],
        "output_prefix":    "revenue_leakage_finder",
    },
    {
        "script":           "data_contract_checker.py",
        "args":             ["--sample"],
        "expected_outputs": ["contract_check_report.html", "issues_detail.csv",
                             "run_log.json", "run_summary.txt"],
        "output_prefix":    "data_contract_checker",
    },
    {
        "script":           "exception_triage_engine.py",
        "args":             ["--sample"],
        "expected_outputs": ["exception_triage_report.html", "ranked_exceptions.csv",
                             "top_10_action_list.csv", "run_log.json", "run_summary.txt"],
        "output_prefix":    "exception_triage_engine",
    },
    {
        "script":           "control_evidence_pack.py",
        "args":             ["--sample"],
        "expected_outputs": ["evidence_summary.html", "manifest.csv",
                             "evidence_readme.txt", "run_log.json", "run_summary.txt"],
        "output_prefix":    "control_evidence_pack",
    },
    {
        "script":           "workbook_dependency_scanner.py",
        "args":             ["--sample"],
        "expected_outputs": ["dependency_report.html", "cross_sheet_refs.csv",
                             "run_log.json", "run_summary.txt"],
        "output_prefix":    "workbook_dependency_scanner",
    },
]


def _find_latest_output(prefix: str, after_ts: float) -> Path | None:
    """Find the newest output folder with the given prefix created after `after_ts`."""
    outputs_root = _ROOT / "outputs"
    if not outputs_root.exists():
        return None
    candidates = sorted(
        [d for d in outputs_root.iterdir()
         if d.is_dir() and prefix in d.name and d.stat().st_ctime >= after_ts],
        reverse=True
    )
    return candidates[0] if candidates else None


def run_test(test: dict) -> tuple[bool, list[str]]:
    """Run one script and verify its outputs. Returns (passed, [messages])."""
    script = _ROOT / test["script"]
    messages = []

    if not script.exists():
        return False, [f"Script not found: {test['script']}"]

    ts_before = time.time() - 1  # 1s buffer for filesystem granularity

    result = subprocess.run(
        [sys.executable, str(script)] + test["args"],
        cwd=str(_ROOT),
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        messages.append(f"Exit code {result.returncode}")
        if result.stderr:
            messages.append(f"Stderr: {result.stderr.strip()[:200]}")
        return False, messages

    out_dir = _find_latest_output(test["output_prefix"], ts_before)
    if out_dir is None:
        return False, ["No output folder found after run"]

    missing = []
    for fname in test["expected_outputs"]:
        if not (out_dir / fname).exists():
            missing.append(fname)

    if missing:
        messages.append(f"Missing output files: {', '.join(missing)}")
        return False, messages

    messages.append(f"Output: {out_dir.name}")
    return True, messages


def main() -> None:
    print()
    print("=" * 60)
    print("  Finance Automation Toolkit - Smoke Test")
    print("=" * 60)
    print()

    results = []
    for test in _TESTS:
        label = test["script"]
        print(f"  Running {label} ...", end="", flush=True)
        passed, messages = run_test(test)
        status = "PASS" if passed else "FAIL"
        print(f"  {status}")
        for msg in messages:
            print(f"    {msg}")
        results.append((label, passed))

    print()
    print("-" * 60)
    total = len(results)
    passed = sum(1 for _, ok in results if ok)
    print(f"  Results: {passed}/{total} passed")
    print()

    all_passed = passed == total
    if all_passed:
        print("  ALL TESTS PASS - toolkit is ready.")
    else:
        print("  FAILURES DETECTED — check messages above.")
        failed = [label for label, ok in results if not ok]
        for label in failed:
            print(f"    FAIL: {label}")

    print()
    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
