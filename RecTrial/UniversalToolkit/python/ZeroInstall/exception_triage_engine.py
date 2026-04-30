# Finance Automation Toolkit v1.0 — iPipeline
# exception_triage_engine.py — scores and ranks exceptions so you know what to review first
#
# What it does:
#   - Reads the exceptions_ranked.csv produced by revenue_leakage_finder.py
#   - Scores each exception on four dimensions: dollar impact, confidence, recency, repeat offender
#   - Combines scores into one priority score and re-ranks
#   - Adds a plain-English "recommended action" line to each exception
#   - Outputs a full ranked report and a top-10 action list
#
# Scoring formula:
#   priority_score = impact * 0.45 + confidence * 0.30 + recency * 0.15 + repeat * 0.10
#
# Usage:
#   python exception_triage_engine.py path/to/exceptions_ranked.csv
#   python exception_triage_engine.py --sample    (uses most recent Revenue Leakage Finder run)
#
# Outputs (in outputs/YYYYMMDD_HHMMSS_exception_triage_engine/):
#   exception_triage_report.html
#   ranked_exceptions.csv
#   top_10_action_list.csv
#   run_log.json
#   run_summary.txt

import sys
from datetime import datetime, date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from common.safe_io import get_output_dir, get_toolkit_root, resolve_input_path, read_csv_safe, write_csv, write_html
from common.logging_utils import RunLogger
from common.report_utils import build_report, metric_row, data_table, note_box


TOOL_NAME = "exception_triage_engine"

# Weights (must sum to 1.0)
W_IMPACT     = 0.45
W_CONFIDENCE = 0.30
W_RECENCY    = 0.15
W_REPEAT     = 0.10

# Base confidence by exception class (0-100)
_CONFIDENCE = {
    "Class1_NoBillingContract": 95,
    "Class3_BaseQtyZero":       92,
    "Class2_StaleContract":     85,
    "Class4_InvoiceDrift":      75,
    "Class5_NameDrift":         65,
}

# Plain-English recommended action per exception class
_ACTION = {
    "Class1_NoBillingContract": (
        "No contract found for this customer. Check CRM and Salesforce for an active "
        "contract or engagement letter. If none exists, stop billing or create a contract."
    ),
    "Class2_StaleContract": (
        "Contract has expired. Confirm with the account owner whether to renew, "
        "renegotiate at current rates, or stop billing."
    ),
    "Class3_BaseQtyZero": (
        "Base quantity is 0 — every transaction bills at overage rate with no included "
        "volume. Confirm this is intentional or correct the contract entry."
    ),
    "Class4_InvoiceDrift": (
        "Invoice amount deviates from expected. Reconcile against contract terms and "
        "check for missed overage reporting, credits, or manual adjustments."
    ),
    "Class5_NameDrift": (
        "Customer name differs between billing and contract records. Standardize the "
        "name across CRM and billing systems to prevent future matching failures."
    ),
}

_CLASS_LABEL = {
    "Class1_NoBillingContract": "No Contract",
    "Class2_StaleContract":     "Stale Contract",
    "Class3_BaseQtyZero":       "Base Qty = 0",
    "Class4_InvoiceDrift":      "Invoice Drift",
    "Class5_NameDrift":         "Name Drift",
}


def _to_float(val: str) -> float:
    try:
        return float(str(val).replace(",", "").replace("$", "").replace("%", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _recency_score(period: str, all_periods: list[str]) -> float:
    """Score 0-100 based on how recent this period is vs the range in the dataset."""
    if not period or not all_periods:
        return 50.0
    try:
        dt = datetime.strptime(period.strip() + "-01", "%Y-%m-%d")
        dates = []
        for p in all_periods:
            try:
                dates.append(datetime.strptime(p.strip() + "-01", "%Y-%m-%d"))
            except ValueError:
                continue
        if not dates:
            return 50.0
        oldest = min(dates).timestamp()
        newest = max(dates).timestamp()
        if newest == oldest:
            return 100.0
        return round((dt.timestamp() - oldest) / (newest - oldest) * 100, 1)
    except (ValueError, AttributeError):
        return 50.0


def score_exceptions(rows: list[dict]) -> list[dict]:
    """Compute priority scores and add action lines. Returns new list (input not modified)."""
    if not rows:
        return []

    # Pre-compute max amount for impact scaling
    amounts = [_to_float(r.get("amount_billed", "0")) for r in rows]
    max_amount = max(amounts) if amounts else 1.0
    if max_amount == 0:
        max_amount = 1.0

    # Pre-compute all billing periods for recency scoring
    all_periods = [r.get("billing_period", "") for r in rows if r.get("billing_period", "").strip()]

    # Count appearances per customer for repeat-offender score
    from collections import Counter
    customer_counts = Counter(r.get("customer_id", "") for r in rows)

    scored = []
    for row in rows:
        exc_class = row.get("exception_class", "")
        amount    = _to_float(row.get("amount_billed", "0"))
        period    = row.get("billing_period", "").strip()

        # Impact: scale 0-100 based on dollar amount
        impact = round((amount / max_amount) * 100, 1) if max_amount else 0.0

        # Confidence: class-based baseline
        confidence = float(_CONFIDENCE.get(exc_class, 70))

        # Recency: 0-100 based on billing period
        recency = _recency_score(period, all_periods)

        # Repeat: has this customer appeared more than once?
        cid = row.get("customer_id", "")
        count = customer_counts.get(cid, 1)
        repeat = 0.0 if count <= 1 else (50.0 if count == 2 else 100.0)

        priority_score = round(
            impact * W_IMPACT + confidence * W_CONFIDENCE +
            recency * W_RECENCY + repeat * W_REPEAT, 1
        )

        action = _ACTION.get(exc_class, "Review this exception and determine appropriate next step.")
        label  = _CLASS_LABEL.get(exc_class, exc_class)

        scored.append({
            **row,
            "class_label":      label,
            "impact_score":     f"{impact:.0f}",
            "confidence_score": f"{confidence:.0f}",
            "recency_score":    f"{recency:.0f}",
            "repeat_score":     f"{repeat:.0f}",
            "priority_score":   f"{priority_score:.1f}",
            "recommended_action": action,
        })

    scored.sort(key=lambda r: float(r["priority_score"]), reverse=True)
    # Add rank
    for i, r in enumerate(scored, 1):
        r["rank"] = str(i)

    return scored


def build_html(input_path: Path, rows: list[dict], scored: list[dict]) -> str:
    n = len(scored)
    top10 = scored[:10]

    avg_score = sum(float(r["priority_score"]) for r in scored) / n if n else 0
    high_count = sum(1 for r in scored if float(r["priority_score"]) >= 70)
    medium_count = sum(1 for r in scored if 40 <= float(r["priority_score"]) < 70)

    subtitle = (
        f"Input: {input_path.name} &nbsp;|&nbsp; {n} exceptions scored &nbsp;|&nbsp; "
        f"Avg priority: {avg_score:.0f}/100"
    )

    cards = [
        {"label": "Exceptions Scored", "value": str(n), "status": "normal"},
        {"label": "High Priority (≥70)", "value": str(high_count), "status": "bad" if high_count else "ok"},
        {"label": "Medium (40-69)", "value": str(medium_count), "status": "warn" if medium_count else "ok"},
        {"label": "Avg Priority Score", "value": f"{avg_score:.0f}", "status": "normal"},
    ]

    note = note_box(
        "<strong>How scoring works:</strong> Each exception is scored on four dimensions: "
        "<strong>Dollar Impact</strong> (45%), <strong>Confidence</strong> (30%), "
        "<strong>Recency</strong> (15%), and <strong>Repeat Customer</strong> (10%). "
        "Scores are 0–100. High-priority exceptions (≥70) should be reviewed first."
    )

    top10_rows = [
        [r["rank"], r["class_label"], r.get("customer_id",""), r.get("customer_name",""),
         r.get("billing_period",""), r.get("amount_billed",""),
         r["priority_score"], r["recommended_action"]]
        for r in top10
    ]
    top10_table = data_table(
        "Top 10 Exceptions — Highest Priority First",
        ["#", "Type", "Customer ID", "Customer Name", "Period", "Amount Billed",
         "Priority Score", "Recommended Action"],
        top10_rows
    )

    all_rows = [
        [r["rank"], r["class_label"], r.get("customer_id",""), r.get("customer_name",""),
         r.get("billing_period",""), r.get("amount_billed",""),
         r["impact_score"], r["confidence_score"], r["recency_score"], r["repeat_score"],
         r["priority_score"]]
        for r in scored
    ]
    all_table = data_table(
        f"All {n} Exceptions — Ranked by Priority Score",
        ["#", "Type", "Customer ID", "Customer Name", "Period", "Amount",
         "Impact", "Confidence", "Recency", "Repeat", "Priority"],
        all_rows
    )

    footer_note = note_box(
        "Scores are relative to this dataset. Impact is scaled against the highest single invoice amount. "
        "Confidence reflects how definitively each exception class can be confirmed. "
        "See <strong>top_10_action_list.csv</strong> for the recommended action list ready to paste into a ticket or email."
    )

    sections = [metric_row(cards), note, top10_table, all_table, footer_note]
    return build_report("Exception Triage Engine", subtitle, sections)


def main(argv: list[str]) -> None:
    sample_mode = "--sample" in argv

    if sample_mode:
        # Find most recent Revenue Leakage Finder run
        outputs_root = get_toolkit_root() / "outputs"
        leakage_dirs = sorted(
            [d for d in outputs_root.iterdir()
             if d.is_dir() and "revenue_leakage_finder" in d.name],
            reverse=True
        ) if outputs_root.exists() else []

        if leakage_dirs:
            input_path = leakage_dirs[0] / "exceptions_ranked.csv"
            print(f"[Sample mode] Using most recent leakage run: {leakage_dirs[0].name}")
        else:
            print("No Revenue Leakage Finder output found. Run revenue_leakage_finder.py --sample first.")
            sys.exit(1)
    elif len(argv) < 2 or argv[1].startswith("--"):
        print("Usage: python exception_triage_engine.py path/to/exceptions_ranked.csv")
        print("       python exception_triage_engine.py --sample")
        sys.exit(0)
    else:
        input_path = resolve_input_path(argv[1])

    out_dir = get_output_dir(TOOL_NAME)
    logger  = RunLogger(TOOL_NAME, out_dir)
    logger.set_meta(input_file=str(input_path), mode="sample" if sample_mode else "real")

    print(f"Input:  {input_path}")
    print(f"Output: {out_dir}")

    try:
        rows = read_csv_safe(input_path)
        logger.rows_read = len(rows)
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        logger.error(str(e))
        logger.finish()
        sys.exit(1)

    if not rows:
        print("Input file is empty or has no data rows.")
        logger.error("Empty input file.")
        logger.finish()
        sys.exit(1)

    print(f"Scoring {len(rows)} exceptions...")
    scored = score_exceptions(rows)
    logger.rows_processed = len(rows)

    top_n = min(10, len(scored))
    for r in scored[:top_n]:
        logger.finding(
            r.get("exception_class", ""),
            f"{r.get('customer_name','')} | Period: {r.get('billing_period','')} | Score: {r['priority_score']}",
            f"Amount: {r.get('amount_billed','')} | Rank: {r['rank']}"
        )

    # All scored exceptions
    ranked_fields = [
        "rank", "exception_class", "class_label", "customer_id", "customer_name",
        "billing_period", "amount_billed", "priority_score",
        "impact_score", "confidence_score", "recency_score", "repeat_score",
        "note", "priority", "recommended_action",
    ]
    write_csv(out_dir / "ranked_exceptions.csv", scored, ranked_fields)

    # Top 10 action list — trimmed for hand-off
    action_fields = [
        "rank", "class_label", "customer_id", "customer_name",
        "billing_period", "amount_billed", "priority_score", "recommended_action",
    ]
    write_csv(out_dir / "top_10_action_list.csv", scored[:10], action_fields)

    html = build_html(input_path, rows, scored)
    write_html(out_dir / "exception_triage_report.html", html)

    logger.finish()

    high = sum(1 for r in scored if float(r["priority_score"]) >= 70)
    med  = sum(1 for r in scored if 40 <= float(r["priority_score"]) < 70)
    avg  = sum(float(r["priority_score"]) for r in scored) / len(scored) if scored else 0

    print(f"\nScored {len(scored)} exceptions  |  Avg score: {avg:.0f}  |  High: {high}  Medium: {med}")
    print(f"\nTop 3:")
    for r in scored[:3]:
        print(f"  #{r['rank']} [{r['class_label']}] {r.get('customer_name','')} | score {r['priority_score']}")
    print(f"\nReport:     {out_dir / 'exception_triage_report.html'}")
    print(f"Top 10 CSV: {out_dir / 'top_10_action_list.csv'}")
    print(f"Full CSV:   {out_dir / 'ranked_exceptions.csv'}")
    print(f"Log:        {out_dir / 'run_summary.txt'}")


if __name__ == "__main__":
    main(sys.argv)
