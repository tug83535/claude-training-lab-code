# Finance Automation Toolkit v1.0 — iPipeline
# revenue_leakage_finder.py — finds billing gaps, stale contracts, and data quality anomalies
#
# What it finds (ranked by revenue impact):
#   Class 1 — Customers billing with NO matching contract row ("polling without contract")
#   Class 2 — Stale contracts: term expired but billing still active
#   Class 3 — Base quantity = 0: every transaction bills at overage rate (data quality risk)
#   Class 4 — Invoice drift: amount billed deviates >10% from expected without overage explanation
#   Class 5 — Name drift: billing name doesn't exactly match contract name (mapping gap)
#
# Outputs (in outputs/YYYYMMDD_HHMMSS_revenue_leakage_finder/):
#   leakage_report.html      — full iPipeline-branded report with all findings
#   exceptions_ranked.csv    — one row per exception, sorted by estimated revenue impact
#   arr_waterfall.html       — visual: expected ARR vs confirmed billing vs gap by category
#   run_log.json
#   run_summary.txt
#
# Usage:
#   python revenue_leakage_finder.py contracts.csv billing.csv
#   python revenue_leakage_finder.py --sample

import sys
import difflib
from datetime import date, datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from common.safe_io import get_output_dir, get_samples_dir, read_csv_safe, write_csv, write_html
from common.logging_utils import RunLogger
from common.report_utils import build_report, metric_row, data_table, note_box


TOOL_NAME = "revenue_leakage_finder"
DRIFT_THRESHOLD = 0.10    # >10% variance triggers Class 4 flag
SIMILARITY_THRESHOLD = 0.70  # SequenceMatcher ratio for Class 5 name drift


def _to_float(val: str) -> float | None:
    try:
        return float(str(val).replace(",", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return None


def _to_date(val: str) -> date | None:
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(val.strip(), fmt).date()
        except (ValueError, AttributeError):
            continue
    return None


def _similar(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()


def run_analysis(contracts: list[dict], billing: list[dict]) -> dict:
    """Run all five leakage checks. Returns structured results dict."""
    # Anchor the stale-contract check to the billing data's time window, not the
    # runner's calendar date — otherwise contracts expire as time passes and the
    # sample counts drift year over year.
    bill_dates = []
    for r in billing:
        p = r.get("billing_period", "").strip()
        if p:
            d = _to_date(p + "-01")
            if d:
                bill_dates.append(d)
    today = (max(bill_dates) + timedelta(days=45)) if bill_dates else date.today()

    # Build contract lookup tables
    contract_by_id:   dict[str, dict] = {}
    contract_by_name: dict[str, dict] = {}
    for row in contracts:
        cid  = row.get("customer_id",   "").strip()
        name = row.get("customer_name", "").strip()
        if cid:
            contract_by_id[cid] = row
        if name:
            contract_by_name[name] = row

    all_contract_names = list(contract_by_name.keys())

    class1_rows: list[dict] = []   # billing without contract
    class2_rows: list[dict] = []   # stale contract still billing
    class3_rows: list[dict] = []   # base_quantity = 0
    class4_rows: list[dict] = []   # invoice amount drift
    class5_rows: list[dict] = []   # name drift / mapping gap

    seen_class3: set[str] = set()   # avoid flagging same contract repeatedly

    total_billed    = 0.0
    total_expected  = 0.0

    for row in billing:
        cid          = row.get("customer_id",   "").strip()
        bname        = row.get("customer_name", "").strip()
        period       = row.get("billing_period","").strip()
        amt_str      = row.get("amount_billed", "").strip()
        amount       = _to_float(amt_str) or 0.0
        total_billed += amount

        contract = contract_by_id.get(cid) or contract_by_name.get(bname)

        # ── Class 1: billing with no contract ────────────────────────────
        if contract is None:
            # Check if name drift could explain the miss
            best_match, best_score = None, 0.0
            for cname in all_contract_names:
                score = _similar(bname, cname)
                if score > best_score:
                    best_score, best_match = score, cname
            near_miss = best_match if best_score >= SIMILARITY_THRESHOLD else None

            class1_rows.append({
                "exception_class": "Class1_NoBillingContract",
                "customer_id":     cid,
                "customer_name":   bname,
                "billing_period":  period,
                "amount_billed":   f"{amount:,.2f}",
                "estimated_annual_risk": f"{amount * 12:,.0f}",
                "note": f"Possible name match: '{near_miss}' (similarity {best_score:.0%})" if near_miss else "No close contract name match found.",
                "priority": "HIGH",
            })
            continue  # skip further checks for this row

        # ── Class 2: stale contract (term expired, still billing) ─────────
        term_end = _to_date(contract.get("term_end", ""))
        if term_end and term_end < today and contract.get("status", "").strip().lower() != "dormant":
            months_expired = max(0, (today.year - term_end.year) * 12 + (today.month - term_end.month))
            class2_rows.append({
                "exception_class":   "Class2_StaleContract",
                "customer_id":       cid,
                "customer_name":     bname,
                "billing_period":    period,
                "amount_billed":     f"{amount:,.2f}",
                "term_end":          str(term_end),
                "months_expired":    str(months_expired),
                "note": "Contract expired — renegotiation may unlock higher rates.",
                "priority": "MEDIUM",
            })

        # ── Class 3: base_quantity = 0 ────────────────────────────────────
        bq_str = contract.get("base_quantity", "").strip()
        if bq_str == "0" and cid not in seen_class3:
            seen_class3.add(cid)
            class3_rows.append({
                "exception_class": "Class3_BaseQtyZero",
                "customer_id":     cid,
                "customer_name":   bname,
                "base_quantity":   bq_str,
                "base_fee":        contract.get("base_fee", ""),
                "band1_rate":      contract.get("band1_rate", ""),
                "note": "base_quantity = 0 means every transaction bills at overage rate. "
                        "Confirm this is intentional — may be a data entry error.",
                "priority": "HIGH",
            })

        # ── Class 4: invoice amount drift ─────────────────────────────────
        exp_str = contract.get("expected_annual_revenue", "").strip()
        exp_rev = _to_float(exp_str)
        billing_basis = contract.get("billing_basis", "").strip()
        if exp_rev and exp_rev > 0 and amount > 0:
            # Scale expected to match billing cadence
            if "Monthly" in billing_basis:
                scaled_expected = exp_rev / 12
            elif "Quarterly" in billing_basis:
                scaled_expected = exp_rev / 4
            else:
                scaled_expected = exp_rev

            drift = (amount - scaled_expected) / scaled_expected if scaled_expected else 0
            if abs(drift) > DRIFT_THRESHOLD:
                direction = "OVER" if drift > 0 else "UNDER"
                class4_rows.append({
                    "exception_class":   "Class4_InvoiceDrift",
                    "customer_id":       cid,
                    "customer_name":     bname,
                    "billing_period":    period,
                    "amount_billed":     f"{amount:,.2f}",
                    "expected_for_period": f"{scaled_expected:,.2f}",
                    "drift_pct":         f"{drift:+.1%}",
                    "direction":         direction,
                    "note": f"Billed {abs(drift):.1%} {direction.lower()} expected. "
                            "Check for overage, credit, or billing error.",
                    "priority": "MEDIUM" if abs(drift) < 0.25 else "HIGH",
                })

        # ── Class 5: name drift ───────────────────────────────────────────
        contract_name = contract.get("customer_name", "").strip()
        if bname and contract_name and bname != contract_name:
            score = _similar(bname, contract_name)
            if score >= SIMILARITY_THRESHOLD:
                class5_rows.append({
                    "exception_class": "Class5_NameDrift",
                    "customer_id":     cid,
                    "customer_name":   bname,
                    "billing_period":  period,
                    "note": f"Billing name '{bname}' differs from contract name '{contract_name}' "
                            f"(similarity {score:.0%}). Confirm these are the same customer.",
                    "priority": "LOW",
                })

        total_expected += _to_float(contract.get("expected_annual_revenue", "") or "0") or 0

    # Deduplicate class5 by customer_id (one flag per customer, not per billing row)
    seen5: set[str] = set()
    class5_unique = []
    for r in class5_rows:
        key = r["customer_id"]
        if key not in seen5:
            seen5.add(key)
            class5_unique.append(r)

    return {
        "class1": class1_rows,
        "class2": class2_rows,
        "class3": class3_rows,
        "class4": class4_rows,
        "class5": class5_unique,
        "total_billed":   total_billed,
        "total_expected": total_expected,
    }


def build_ranked_exceptions(results: dict) -> list[dict]:
    """Combine all classes into one ranked list, highest-impact first."""
    all_rows = []

    for row in results["class1"]:
        risk = _to_float(row.get("estimated_annual_risk", "0").replace(",", "")) or 0
        all_rows.append({"priority_score": 300 + risk, **row})

    for row in results["class2"]:
        amt = _to_float(row.get("amount_billed", "0").replace(",", "")) or 0
        all_rows.append({"priority_score": 200 + amt, **row})

    for row in results["class3"]:
        all_rows.append({"priority_score": 250, **row})

    for row in results["class4"]:
        amt = _to_float(row.get("amount_billed", "0").replace(",", "")) or 0
        drift = abs(_to_float(row.get("drift_pct", "0").rstrip("%")) or 0)
        all_rows.append({"priority_score": 100 + amt * drift, **row})

    for row in results["class5"]:
        all_rows.append({"priority_score": 50, **row})

    all_rows.sort(key=lambda r: r["priority_score"], reverse=True)

    # Clean up the sort key before export
    for r in all_rows:
        r.pop("priority_score", None)

    return all_rows


def _waterfall_bar(label: str, value: float, max_val: float,
                   color: str, pct_label: str = "") -> str:
    pct = min(100, int(value / max_val * 100)) if max_val > 0 else 0
    return (
        f'<tr>'
        f'<td style="width:200px;font-weight:bold;padding:8px 12px">{label}</td>'
        f'<td style="padding:8px 4px">'
        f'  <div style="background:#eee;border-radius:3px;height:28px;position:relative">'
        f'    <div style="background:{color};width:{pct}%;height:100%;border-radius:3px;'
        f'         display:flex;align-items:center;padding:0 8px;box-sizing:border-box;'
        f'         color:#fff;font-size:13px;font-weight:bold;white-space:nowrap">'
        f'      ${value:,.0f}'
        f'    </div>'
        f'  </div>'
        f'</td>'
        f'<td style="padding:8px 12px;color:#666;font-size:12px">{pct_label}</td>'
        f'</tr>'
    )


def build_waterfall_html(results: dict) -> str:
    """Build the ARR gap waterfall as a standalone iPipeline-branded HTML file."""
    c1 = results["class1"]
    c2 = results["class2"]
    c4 = results["class4"]

    total_exp    = results["total_expected"]
    total_billed = results["total_billed"]
    gap_total    = total_exp - total_billed

    # Estimate ARR at-risk per category
    c1_risk = sum(
        (_to_float(r.get("estimated_annual_risk", "0").replace(",", "")) or 0) for r in c1
    )
    c2_risk = sum(
        (_to_float(r.get("amount_billed", "0").replace(",", "")) or 0) * 0.15
        for r in c2
    )  # 15% uplift opportunity from renegotiation
    c4_risk = sum(
        abs(_to_float(r.get("amount_billed", "0").replace(",", "")) or 0) *
        abs(_to_float((r.get("drift_pct","0%").rstrip("%"))) or 0) / 100
        for r in c4
    )

    max_val = max(total_exp, total_billed, c1_risk, 1)

    rows_html = (
        _waterfall_bar("Expected ARR", total_exp, max_val, "#0B4779") +
        _waterfall_bar("Confirmed Billing", total_billed, max_val, "#2a7a2a",
                       f"{total_billed/total_exp:.0%} of expected" if total_exp else "") +
        _waterfall_bar("No-contract customers", c1_risk, max_val, "#a00",
                       f"{len(c1)} customers at risk") +
        _waterfall_bar("Stale-contract opportunity", c2_risk, max_val, "#c65000",
                       f"{len(set(r['customer_id'] for r in c2))} contracts") +
        _waterfall_bar("Invoice drift gap", c4_risk, max_val, "#7a5000",
                       f"{len(c4)} invoices flagged")
    )

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>ARR Gap Waterfall</title>
<style>
body{{font-family:Arial,sans-serif;background:#F9F9F9;color:#161616;margin:0}}
.hdr{{background:#0B4779;color:#fff;padding:20px 32px}}
.hdr h1{{margin:0;font-size:20px}}
.hdr .sub{{font-size:13px;color:#BFF18C;margin-top:4px}}
.body{{padding:28px 32px}}
table{{width:100%;border-collapse:collapse}}
.footer{{background:#112E51;color:#99b;font-size:11px;padding:10px 32px;margin-top:32px}}
</style>
</head>
<body>
<div class="hdr">
  <h1>ARR Gap Analysis — Finance Automation Toolkit</h1>
  <div class="sub">Revenue Leakage Finder &nbsp;|&nbsp; Generated {now}</div>
</div>
<div class="body">
<h2 style="color:#0B4779;border-bottom:2px solid #0B4779;padding-bottom:5px">
  Expected vs Confirmed Billing — Annual Revenue at Risk
</h2>
<table style="margin-top:16px">{rows_html}</table>
<p style="margin-top:20px;font-size:13px;color:#555">
  <strong>How to read this chart:</strong> The top bar shows total expected ARR from all contracts.
  Each row below shows a category of billing gap or risk. Longer bars = larger estimated impact.
  All amounts are based on contract expected_annual_revenue and invoice data in the input files.
</p>
<p style="margin-top:8px;font-size:12px;color:#888">
  <em>Note: estimated_annual_risk for no-contract customers is annualized from the most recent billing period.
  Stale-contract opportunity assumes 15% rate improvement on renegotiation — adjust as appropriate.
  All amounts are in USD unless billing file contains other currencies.</em>
</p>
</div>
<div class="footer">Finance Automation Toolkit v1.0 &nbsp;|&nbsp; iPipeline &nbsp;|&nbsp; {now}</div>
</body></html>"""


def build_main_html(contracts: list[dict], billing: list[dict], results: dict) -> str:
    c1, c2, c3, c4, c5 = (results[k] for k in ("class1","class2","class3","class4","class5"))
    n_exceptions = len(c1) + len(c2) + len(c3) + len(c4) + len(c5)
    subtitle = (
        f"{len(contracts)} contracts &nbsp;|&nbsp; {len(billing)} invoices &nbsp;|&nbsp; "
        f"{n_exceptions} exceptions found"
    )

    total_exp    = results["total_expected"]
    total_billed = results["total_billed"]

    overall_status = "bad" if (c1 or c3) else ("warn" if (c2 or c4 or c5) else "ok")
    cards = [
        {"label": "Total Exceptions", "value": str(n_exceptions),
         "status": overall_status},
        {"label": "No Contract (Class 1)", "value": str(len(c1)), "status": "bad" if c1 else "ok"},
        {"label": "Stale Contract (Class 2)", "value": str(len(c2)), "status": "warn" if c2 else "ok"},
        {"label": "Base Qty = 0 (Class 3)", "value": str(len(c3)), "status": "bad" if c3 else "ok"},
        {"label": "Invoice Drift (Class 4)", "value": str(len(c4)), "status": "warn" if c4 else "ok"},
        {"label": "Name Drift (Class 5)", "value": str(len(c5)), "status": "warn" if c5 else "ok"},
    ]

    sections = [metric_row(cards)]

    sections.append(
        note_box(
            f"<strong>Expected ARR (from contracts):</strong> ${total_exp:,.0f} &nbsp;|&nbsp; "
            f"<strong>Confirmed Billing:</strong> ${total_billed:,.0f} &nbsp;|&nbsp; "
            f"<strong>Gap:</strong> ${total_exp - total_billed:,.0f} "
            f"({(total_exp - total_billed)/total_exp:.1%} of expected)"
            if total_exp else "Expected ARR not calculable from this dataset."
        )
    )

    if c1:
        rows = [[r.get("customer_id",""), r.get("customer_name",""),
                 r.get("billing_period",""), r.get("amount_billed",""),
                 r.get("estimated_annual_risk",""), r.get("note","")]
                for r in c1]
        sections.append(data_table(
            f"Class 1 — Customers Billing with No Contract on File ({len(c1)} rows)",
            ["Customer ID","Customer Name","Period","Billed","Est. Annual Risk","Note"],
            rows, status_col=None
        ))

    if c2:
        rows = [[r.get("customer_id",""), r.get("customer_name",""),
                 r.get("term_end",""), r.get("months_expired",""),
                 r.get("amount_billed",""), r.get("note","")]
                for r in c2]
        sections.append(data_table(
            f"Class 2 — Stale Contracts (Expired Term, Active Billing) ({len(c2)} rows)",
            ["Customer ID","Customer Name","Term End","Months Expired","Last Billed","Note"],
            rows
        ))

    if c3:
        rows = [[r.get("customer_id",""), r.get("customer_name",""),
                 r.get("base_fee",""), r.get("band1_rate",""), r.get("note","")]
                for r in c3]
        sections.append(data_table(
            f"Class 3 — Base Quantity = 0 Anomaly ({len(c3)} contracts)",
            ["Customer ID","Customer Name","Base Fee","Band 1 Rate","Note"],
            rows
        ))

    if c4:
        rows = [[r.get("customer_id",""), r.get("customer_name",""),
                 r.get("billing_period",""), r.get("amount_billed",""),
                 r.get("expected_for_period",""), r.get("drift_pct",""),
                 r.get("direction",""), r.get("note","")]
                for r in c4]
        sections.append(data_table(
            f"Class 4 — Invoice Amount Drift (>{DRIFT_THRESHOLD:.0%} variance) ({len(c4)} rows)",
            ["Customer ID","Customer Name","Period","Billed","Expected","Drift %","Dir","Note"],
            rows
        ))

    if c5:
        rows = [[r.get("billing_name",""), r.get("contract_name",""),
                 r.get("customer_id",""), r.get("similarity",""), r.get("note","")]
                for r in c5]
        sections.append(data_table(
            f"Class 5 — Customer Name Drift ({len(c5)} mappings)",
            ["Billing Name","Contract Name","ID","Similarity","Note"],
            rows
        ))

    sections.append(note_box(
        "See <strong>arr_waterfall.html</strong> in this output folder for the ARR gap visual. "
        "See <strong>exceptions_ranked.csv</strong> for the full ranked exception list. "
        "Safety: input files were opened read-only. No data was changed."
    ))

    return build_report("Revenue Leakage Finder", subtitle, sections)


def main(argv: list[str]) -> None:
    sample_mode = "--sample" in argv
    samples_dir = get_samples_dir()

    if sample_mode:
        contracts_path = samples_dir / "contracts_sample.csv"
        billing_path   = samples_dir / "billing_sample.csv"
        print(f"[Sample mode] Contracts: {contracts_path.name}")
        print(f"[Sample mode] Billing:   {billing_path.name}")
    else:
        non_flag = [a for a in argv[1:] if not a.startswith("--")]
        if len(non_flag) < 2:
            print("Usage: python revenue_leakage_finder.py contracts.csv billing.csv")
            print("       python revenue_leakage_finder.py --sample")
            sys.exit(0)
        from common.safe_io import resolve_input_path
        contracts_path = resolve_input_path(non_flag[0])
        billing_path   = resolve_input_path(non_flag[1])

    out_dir = get_output_dir(TOOL_NAME)
    logger  = RunLogger(TOOL_NAME, out_dir)
    logger.set_meta(
        contracts_file=str(contracts_path),
        billing_file=str(billing_path),
        mode="sample" if sample_mode else "real",
    )

    print(f"Contracts: {contracts_path}")
    print(f"Billing:   {billing_path}")
    print(f"Output:    {out_dir}")

    try:
        contracts = read_csv_safe(contracts_path)
        billing   = read_csv_safe(billing_path)
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        logger.error(str(e))
        logger.finish()
        sys.exit(1)

    logger.rows_read = len(contracts) + len(billing)
    print(f"Loaded {len(contracts)} contracts, {len(billing)} billing rows. Analyzing...")

    results = run_analysis(contracts, billing)
    logger.rows_processed = len(contracts) + len(billing)

    for cls, label in [("class1","Class1_NoBillingContract"), ("class2","Class2_StaleContract"),
                        ("class3","Class3_BaseQtyZero"), ("class4","Class4_InvoiceDrift"),
                        ("class5","Class5_NameDrift")]:
        for r in results[cls]:
            logger.finding(
                label,
                r.get("customer_name","") + (": " + r.get("note",""))[:80],
                r.get("estimated_annual_risk", r.get("amount_billed",""))
            )

    ranked = build_ranked_exceptions(results)
    write_csv(out_dir / "exceptions_ranked.csv", ranked)
    write_html(out_dir / "leakage_report.html", build_main_html(contracts, billing, results))
    write_html(out_dir / "arr_waterfall.html",  build_waterfall_html(results))

    logger.finish()

    n = sum(len(results[k]) for k in ("class1","class2","class3","class4","class5"))
    print(f"\nFindings: {n} exceptions across 5 classes")
    print(f"  Class 1 (no contract):    {len(results['class1'])}")
    print(f"  Class 2 (stale contract): {len(results['class2'])}")
    print(f"  Class 3 (base qty = 0):   {len(results['class3'])}")
    print(f"  Class 4 (invoice drift):  {len(results['class4'])}")
    print(f"  Class 5 (name drift):     {len(results['class5'])}")
    print(f"\nReport:    {out_dir / 'leakage_report.html'}")
    print(f"Waterfall: {out_dir / 'arr_waterfall.html'}")
    print(f"CSV:       {out_dir / 'exceptions_ranked.csv'}")
    print(f"Log:       {out_dir / 'run_summary.txt'}")


if __name__ == "__main__":
    main(sys.argv)
