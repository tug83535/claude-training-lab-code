# Finance Automation Toolkit v1.0 — iPipeline
# common/sample_data.py — synthetic sample data generator for Revenue Leakage Finder demo
#
# Generates contracts_sample.csv (~150 rows) and billing_sample.csv (~330 rows)
# in the samples/ folder. All data is fictional. Customer names, IDs, and amounts
# are invented and do not represent any real iPipeline customer.
#
# Embedded exception classes (per VIDEO_4_REVISED_PLAN.md Section 6):
#   Class 1 — 12 customers billing/polling with NO contract row (polling-without-contract)
#   Class 2 — 10 stale contracts (term_end expired 3-18 months ago, still billing)
#   Class 3 — 2 contracts with base_quantity = 0 (data quality anomaly)
#   Class 4 — 9 invoices where amount_billed drifts >10% from expected (no overage reason)
#   Class 5 — 4 name-drift cases (same customer, different spelling across files)
#   Class 6 — 2 duplicate invoice rows (same invoice_id, different billing_period)
#
# Run directly to (re)generate sample files:
#   python common/sample_data.py

import csv
import os
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(42)

SAMPLES_DIR = Path(__file__).resolve().parent.parent / "samples"

# --- Customer roster (fictional insurance/financial services firms) ---
# Asterisk (*) marks firms that will appear in billing but NOT in contracts (Class 1 leakage)

_MAIN_CUSTOMERS = [
    # (name, deployment_type, environment, tier)   tier: enterprise/mid/small
    ("Northstar Insurance Group",    "SaaS",     "eSign",    "enterprise"),
    ("Northstar Insurance Group",    "On-Prem",  "On-Prem",  "enterprise"),   # 2nd deployment
    ("Harbor Life Systems",          "On-Prem",  "On-Prem",  "mid"),
    ("Pioneer Benefits Co.",         "SaaS",     "eSign",    "mid"),
    ("Summit Policy Services",       "Embedded", "eSign",    "small"),
    ("Keystone Admin Solutions",     "On-Prem",  "On-Prem",  "mid"),
    ("Atlas Brokerage Network",      "SaaS",     "Pronto2",  "mid"),
    ("Meridian Life & Casualty",     "SaaS",     "eSign",    "enterprise"),
    ("Meridian Life & Casualty",     "On-Prem",  "On-Prem",  "enterprise"),   # 2nd deployment
    ("Lakewood Financial Group",     "SaaS",     "eSign",    "small"),
    ("Crestview Insurance Partners", "SaaS",     "Pronto2",  "mid"),
    ("Ironbridge Benefits",          "On-Prem",  "On-Prem",  "mid"),
    ("Stillwater Mutual",            "SaaS",     "eSign",    "enterprise"),
    ("BluePeak Advisory",            "Embedded", "eSign",    "small"),
    ("Ridgeline Health Plans",       "SaaS",     "eSign",    "mid"),
    ("Coastal Re Solutions",         "SaaS",     "Pronto2",  "small"),
    ("Pinnacle Trust Services",      "On-Prem",  "On-Prem",  "enterprise"),
    # UK customers (GBP currency, UK_AWS environment)
    ("Cirencester Benefits Ltd",     "SaaS",     "UK_AWS",   "mid"),
    ("Aldgate Re Partners",          "SaaS",     "UK_AWS",   "small"),
]

_ADDITIONAL_CUSTOMERS = [
    ("Broadview Life Partners",      "SaaS",     "eSign",    "mid"),
    ("Cascade Insurance Group",      "On-Prem",  "On-Prem",  "mid"),
    ("Dellwood Mutual Benefits",     "SaaS",     "eSign",    "small"),
    ("Eastbrook Policy Services",    "Embedded", "eSign",    "small"),
    ("Frontier Benefits Assoc.",     "SaaS",     "Pronto2",  "mid"),
    ("Granite State Insurance",      "On-Prem",  "On-Prem",  "mid"),
    ("Harborview Life & Annuity",    "SaaS",     "eSign",    "enterprise"),
    ("Inland Empire Brokers",        "SaaS",     "eSign",    "mid"),
    ("Junction Benefits Corp.",      "On-Prem",  "On-Prem",  "mid"),
    ("Kingsbridge Casualty",         "SaaS",     "eSign",    "small"),
    ("Lakeshore Annuity Services",   "SaaS",     "Pronto2",  "mid"),
    ("Mountainview Mutual",          "On-Prem",  "On-Prem",  "mid"),
    ("Northfield Insurance Co.",     "SaaS",     "eSign",    "mid"),
    ("Oakdale Benefits Group",       "SaaS",     "eSign",    "small"),
    ("Pacific Crest Insurance",      "SaaS",     "eSign",    "enterprise"),
    ("Quorum Financial Services",    "On-Prem",  "On-Prem",  "mid"),
    ("Redwood Benefits Partners",    "SaaS",     "Pronto2",  "mid"),
    ("Silvergate Insurance",         "SaaS",     "eSign",    "mid"),
    ("Timberline Policy Group",      "On-Prem",  "On-Prem",  "small"),
    ("Union Square Benefits",        "SaaS",     "eSign",    "mid"),
    ("Valley Trust Insurance",       "Embedded", "eSign",    "small"),
    ("Westport Mutual Holdings",     "SaaS",     "eSign",    "enterprise"),
    ("Xerxes Benefits Network",      "On-Prem",  "On-Prem",  "mid"),
    ("Yosemite Life Systems",        "SaaS",     "eSign",    "mid"),
    ("Zenith Policy Services",       "SaaS",     "Pronto2",  "mid"),
    ("Arcadian Insurance Trust",     "SaaS",     "eSign",    "mid"),
    ("Bayshore Life & Casualty",     "On-Prem",  "On-Prem",  "mid"),
    ("Clearwater Benefits Inc.",     "SaaS",     "eSign",    "small"),
    ("Driftwood Annuity Group",      "SaaS",     "eSign",    "small"),
    ("Elmwood Financial Partners",   "Embedded", "eSign",    "small"),
    ("Fairview Insurance Network",   "SaaS",     "Pronto2",  "mid"),
    ("Glenwood Mutual Services",     "On-Prem",  "On-Prem",  "mid"),
    ("Highmark Benefits Co.",        "SaaS",     "eSign",    "mid"),
    ("Ironwood Policy Advisors",     "SaaS",     "eSign",    "small"),
    ("Juniper Life Holdings",        "On-Prem",  "On-Prem",  "enterprise"),
    ("Kirkland Insurance Partners",  "SaaS",     "eSign",    "mid"),
    ("Laurel Benefits Group",        "SaaS",     "Pronto2",  "small"),
    ("Maple Ridge Insurance",        "On-Prem",  "On-Prem",  "mid"),
    ("Newbury Policy Services",      "SaaS",     "eSign",    "mid"),
    ("Orchard Trust Benefits",       "SaaS",     "eSign",    "small"),
    ("Plum Creek Mutual",            "Embedded", "eSign",    "small"),
    ("Queensway Financial Group",    "SaaS",     "eSign",    "mid"),
    ("Rosewood Life Partners",       "On-Prem",  "On-Prem",  "mid"),
    ("Sagebrush Insurance Co.",      "SaaS",     "eSign",    "mid"),
    ("Thornwood Benefits Inc.",      "SaaS",     "Pronto2",  "mid"),
    ("Upland Casualty Services",     "SaaS",     "eSign",    "small"),
    ("Vineyard Life Systems",        "On-Prem",  "On-Prem",  "mid"),
    ("Whitmore Insurance Trust",     "SaaS",     "eSign",    "mid"),
    ("Xenon Policy Partners",        "Embedded", "eSign",    "small"),
    ("Yellowstone Mutual Holdings",  "SaaS",     "eSign",    "enterprise"),
    ("Zion Benefits Network",        "On-Prem",  "On-Prem",  "mid"),
    ("Ashford Life Services",        "SaaS",     "Pronto2",  "mid"),
    ("Birchwood Insurance Co.",      "SaaS",     "eSign",    "small"),
    ("Crestwood Benefits Alliance",  "On-Prem",  "On-Prem",  "mid"),
    ("Devonshire Policy Group",      "SaaS",     "eSign",    "mid"),
    ("Elsinore Mutual Partners",     "SaaS",     "eSign",    "small"),
    ("Foxhill Insurance Services",   "On-Prem",  "On-Prem",  "mid"),
    ("Greystone Life Holdings",      "SaaS",     "eSign",    "enterprise"),
    ("Hillcrest Benefits Trust",     "SaaS",     "Pronto2",  "mid"),
    ("Inverness Policy Co.",         "Embedded", "eSign",    "small"),
    ("Juneau Life Partners",         "SaaS",     "eSign",    "mid"),
    ("Kestrel Insurance Network",    "On-Prem",  "On-Prem",  "mid"),
    ("Linwood Mutual Benefits",      "SaaS",     "eSign",    "small"),
    ("Millbrook Policy Advisors",    "SaaS",     "eSign",    "mid"),
    ("Norwood Insurance Alliance",   "On-Prem",  "On-Prem",  "mid"),
    ("Overland Benefits Group",      "SaaS",     "Pronto2",  "mid"),
    ("Pinehurst Life Services",      "SaaS",     "eSign",    "small"),
    ("Quarry Hill Insurance",        "Embedded", "eSign",    "small"),
    ("Riverdale Mutual Holdings",    "SaaS",     "eSign",    "enterprise"),
    ("Stonegate Benefits Co.",       "On-Prem",  "On-Prem",  "mid"),
    ("Thornbury Policy Services",    "SaaS",     "eSign",    "mid"),
    ("Upton Life Partners",          "SaaS",     "Pronto2",  "small"),
    ("Verdant Insurance Trust",      "On-Prem",  "On-Prem",  "mid"),
    ("Weston Benefits Alliance",     "SaaS",     "eSign",    "mid"),
    ("Ximena Policy Group",          "SaaS",     "eSign",    "small"),
    ("Yellowrock Mutual Services",   "On-Prem",  "On-Prem",  "mid"),
    ("Zephyr Benefits Inc.",         "SaaS",     "Pronto2",  "small"),
    ("Andover Insurance Co.",        "SaaS",     "eSign",    "mid"),
    ("Belmont Life Holdings",        "On-Prem",  "On-Prem",  "mid"),
    ("Clearfield Benefits Trust",    "SaaS",     "eSign",    "small"),
    ("Dunmore Policy Services",      "Embedded", "eSign",    "small"),
    ("Evergreen Insurance Partners", "SaaS",     "eSign",    "enterprise"),
    ("Falmouth Mutual Benefits",     "On-Prem",  "On-Prem",  "mid"),
    ("Grantham Life Systems",        "SaaS",     "Pronto2",  "mid"),
    ("Hartwell Insurance Co.",       "SaaS",     "eSign",    "mid"),
    ("Ingram Benefits Alliance",     "On-Prem",  "On-Prem",  "mid"),
    ("Jasper Policy Services",       "SaaS",     "eSign",    "small"),
    ("Kingston Mutual Partners",     "SaaS",     "eSign",    "mid"),
    ("Lockwood Life Insurance",      "Embedded", "eSign",    "small"),
    ("Marlowe Benefits Group",       "SaaS",     "eSign",    "mid"),
    ("Norbury Policy Trust",         "On-Prem",  "On-Prem",  "mid"),
    ("Oakwood Mutual Services",      "SaaS",     "Pronto2",  "mid"),
    ("Prescott Insurance Holdings",  "SaaS",     "eSign",    "enterprise"),
    ("Quinby Benefits Co.",          "SaaS",     "eSign",    "small"),
    ("Ravenwood Life Partners",      "On-Prem",  "On-Prem",  "mid"),
    ("Sedgwick Insurance Services",  "SaaS",     "eSign",    "mid"),
    ("Talbot Benefits Alliance",     "SaaS",     "Pronto2",  "mid"),
    ("Underwood Policy Group",       "Embedded", "eSign",    "small"),
    ("Vantage Life Systems",         "SaaS",     "eSign",    "mid"),
    ("Waterford Mutual Holdings",    "On-Prem",  "On-Prem",  "enterprise"),
    ("Axford Benefits Network",      "SaaS",     "eSign",    "mid"),
    ("Buckhill Insurance Co.",       "On-Prem",  "On-Prem",  "small"),
    ("Castlewood Policy Partners",   "SaaS",     "eSign",    "mid"),
    ("Darnley Life Services",        "SaaS",     "Pronto2",  "small"),
]

# Class 1 leakage: customers who appear in billing but have NO contract row
_LEAKAGE_CUSTOMERS = [
    ("Westfield Annuity Group",      "SaaS",     "eSign"),
    ("Riverton Life Partners",       "On-Prem",  "On-Prem"),
    ("Brookfield Insurance Trust",   "SaaS",     "eSign"),
    ("Marlborough Benefits Inc.",    "SaaS",     "Pronto2"),
    ("Thornfield Policy Services",   "On-Prem",  "On-Prem"),
    ("Sunnyside Mutual Holdings",    "SaaS",     "eSign"),
    ("Cloverdale Insurance Group",   "On-Prem",  "On-Prem"),
    ("Ridgemont Life Systems",       "SaaS",     "eSign"),
    ("Greenlawn Benefits Co.",       "Embedded", "eSign"),
    ("Fairfield Policy Trust",       "SaaS",     "eSign"),
    ("Millstone Insurance Partners", "On-Prem",  "On-Prem"),
    ("Copperview Life Services",     "SaaS",     "Pronto2"),
]

# Class 5 name-drift: billing.csv uses a slightly different spelling for these contracts
# Maps exact contract name → drifted billing name
_NAME_DRIFT = {
    "Atlas Brokerage Network":      "Atlas Brokerage",
    "Meridian Life & Casualty":     "Meridian Life Casualty",   # '&' dropped — common CRM truncation
    "Crestview Insurance Partners": "Crestview Insurance",
    "Ironbridge Benefits":          "Ironbridge Benefit",
}

TODAY = date(2025, 4, 1)   # fixed reference date so data is stable
BILLING_START = date(2024, 4, 1)
BILLING_END   = date(2025, 3, 31)


def _random_date_within(start: date, end: date) -> date:
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))


def _term_dates(tier: str, stale: bool = False) -> tuple[str, str]:
    if stale:
        # Term ended 3-18 months ago
        months_ago = random.randint(3, 18)
        term_end = TODAY - timedelta(days=30 * months_ago)
        term_start = term_end - timedelta(days=random.choice([365, 730, 1095]))
    else:
        term_start = TODAY - timedelta(days=random.randint(180, 1825))
        term_end   = TODAY + timedelta(days=random.randint(30, 730))
    return str(term_start), str(term_end)


def _billing_params(tier: str) -> dict:
    """Generate base_fee, base_quantity, band1_rate, transaction_multiplier by tier."""
    if tier == "enterprise":
        base_fee      = round(random.uniform(80_000, 333_000), -2)
        base_quantity = random.randint(200_000, 2_000_000)
        band1_rate    = round(random.uniform(0.12, 0.60), 3)
        multiplier    = round(random.uniform(0.8, 1.5), 2)
    elif tier == "mid":
        base_fee      = round(random.uniform(10_000, 80_000), -2)
        base_quantity = random.randint(7_000, 200_000)
        band1_rate    = round(random.uniform(0.50, 1.80), 3)
        multiplier    = round(random.uniform(0.5, 2.0), 2)
    else:  # small
        base_fee      = round(random.uniform(0, 15_000), -2)
        base_quantity = random.randint(0, 15_000)
        band1_rate    = round(random.uniform(1.00, 2.00), 3)
        multiplier    = round(random.uniform(0.2, 3.0), 2)
    return {
        "base_fee": base_fee,
        "base_quantity": base_quantity,
        "band1_rate": band1_rate,
        "transaction_multiplier": multiplier,
    }


def _expected_revenue(base_fee: float, base_quantity: int, band1_rate: float,
                      multiplier: float) -> float:
    """Approximate annual expected revenue: base_fee + modest overage estimate."""
    avg_volume = base_quantity * multiplier * random.uniform(1.0, 1.4)
    overage_txns = max(0, avg_volume - base_quantity)
    return round(base_fee + overage_txns * band1_rate, 2)


def build_contracts() -> list[dict]:
    """Build ~150 contract rows with embedded exception classes 2 and 3."""
    rows = []
    cid = 1

    all_customers = list(_MAIN_CUSTOMERS) + list(_ADDITIONAL_CUSTOMERS)
    # Limit to produce ~150 rows total
    all_customers = all_customers[:130]

    # Pick indices for exception classes
    stale_indices = set(random.sample(range(len(_MAIN_CUSTOMERS), len(all_customers)), 10))
    base_qty_zero = set(random.sample(range(len(all_customers)), 2))

    for idx, (name, deploy, env, tier) in enumerate(all_customers):
        currency = "GBP" if env == "UK_AWS" else "USD"
        billing_basis = random.choices(
            ["Annually", "Annually_in_arrears", "Monthly_in_arrears", "Monthly", "Quarterly_in_arrears"],
            weights=[70, 7, 8, 5, 5],
        )[0]
        status = "Active"
        stale = idx in stale_indices
        if stale:
            status = "Stale"

        term_start, term_end = _term_dates(tier, stale=stale)
        params = _billing_params(tier)

        if idx in base_qty_zero:
            params["base_quantity"] = 0  # Class 3 anomaly

        exp_rev = _expected_revenue(
            params["base_fee"], params["base_quantity"],
            params["band1_rate"], params["transaction_multiplier"]
        )

        rows.append({
            "customer_id":              f"ATC-{cid:05d}",
            "customer_name":            name,
            "deployment_type":          deploy,
            "environment":              env,
            "billing_basis":            billing_basis,
            "status":                   status,
            "currency":                 currency,
            "base_fee":                 params["base_fee"],
            "base_quantity":            params["base_quantity"],
            "band1_rate":               params["band1_rate"],
            "transaction_multiplier":   params["transaction_multiplier"],
            "term_start":               term_start,
            "term_end":                 term_end,
            "expected_annual_revenue":  exp_rev,
        })
        cid += 1

    return rows


def build_billing(contracts: list[dict]) -> list[dict]:
    """Build ~330 billing rows with embedded exception classes 1, 4, 5, and 6."""
    rows = []
    inv_num = 1

    # Build a lookup of contract by name for expected-revenue reference
    contract_by_name = {r["customer_name"]: r for r in contracts}
    contract_ids     = {r["customer_name"]: r["customer_id"] for r in contracts}

    # Helper: generate one billing row
    def make_invoice(customer_id, customer_name, period: date, base_amount: float,
                     drift: float = 0.0, duplicate_id: str | None = None,
                     status: str = "Paid") -> dict:
        nonlocal inv_num
        iid = duplicate_id or f"INV-{period.year}-{inv_num:04d}"
        inv_num += 1
        amount = round(base_amount * (1 + drift) * random.uniform(0.97, 1.03), 2)
        invoice_date = period + timedelta(days=random.randint(1, 28))
        return {
            "invoice_id":     iid,
            "customer_id":    customer_id,
            "customer_name":  customer_name,
            "billing_period": period.strftime("%Y-%m"),
            "amount_billed":  amount,
            "invoice_date":   str(invoice_date),
            "status":         status,
            "notes":          "",
        }

    # Generate normal invoices for all active contracts
    drift_candidates = []
    for contract in contracts:
        name  = contract["customer_name"]
        cid   = contract["customer_id"]
        basis = contract["billing_basis"]
        rev   = float(contract["expected_annual_revenue"])
        billing_name = _NAME_DRIFT.get(name, name)  # Class 5: use drifted name if applicable

        if "Monthly" in basis:
            # Monthly billing: one invoice per month
            cur = BILLING_START
            while cur <= BILLING_END:
                rows.append(make_invoice(cid, billing_name, cur, rev / 12))
                nxt_month = cur.month + 1
                nxt_year  = cur.year + (1 if nxt_month > 12 else 0)
                cur = date(nxt_year, nxt_month % 12 or 12, 1)
        elif "Quarterly" in basis:
            for q_start in [BILLING_START, date(2024, 7, 1), date(2024, 10, 1), date(2025, 1, 1)]:
                if q_start <= BILLING_END:
                    rows.append(make_invoice(cid, billing_name, q_start, rev / 4))
        else:
            # Annual billing
            rows.append(make_invoice(cid, billing_name, BILLING_START, rev))
            drift_candidates.append(len(rows) - 1)  # track for Class 4

    # Class 4 — invoice drift: inflate or deflate 9 invoices by >10% with no overage note
    drift_picks = random.sample(drift_candidates, min(9, len(drift_candidates)))
    for idx in drift_picks:
        direction = random.choice([-1, 1])
        drift_pct = random.uniform(0.12, 0.35) * direction
        orig = rows[idx]
        orig["amount_billed"] = round(float(orig["amount_billed"]) * (1 + drift_pct), 2)

    # Class 1 — leakage customers: billing rows with NO contract row
    for name, deploy, env in _LEAKAGE_CUSTOMERS:
        # Use a fake customer_id that doesn't match any contract
        fake_id = f"ATC-ZZZZZ"   # intentionally not in contracts
        rev_est = round(random.uniform(8_000, 150_000), -2)
        rows.append(make_invoice(fake_id, name, BILLING_START, rev_est))

    # Class 6 — duplicate invoices: duplicate 2 existing invoice_ids in a different period
    dup_sources = random.sample([r for r in rows if r["customer_id"] != "ATC-ZZZZZ"], 2)
    for src in dup_sources:
        dup = dict(src)
        dup["invoice_id"]     = src["invoice_id"]   # same ID
        dup["billing_period"] = (
            BILLING_START + timedelta(days=random.randint(30, 180))
        ).strftime("%Y-%m")
        dup["notes"] = "Possible duplicate — review"
        rows.append(dup)

    random.shuffle(rows)
    return rows


def generate(output_dir: Path | None = None) -> tuple[Path, Path]:
    """Generate both sample CSV files. Returns (contracts_path, billing_path)."""
    out = output_dir or SAMPLES_DIR
    out.mkdir(parents=True, exist_ok=True)

    contracts = build_contracts()
    billing   = build_billing(contracts)

    c_path = out / "contracts_sample.csv"
    b_path = out / "billing_sample.csv"

    contract_fields = [
        "customer_id", "customer_name", "deployment_type", "environment",
        "billing_basis", "status", "currency", "base_fee", "base_quantity",
        "band1_rate", "transaction_multiplier", "term_start", "term_end",
        "expected_annual_revenue",
    ]
    billing_fields = [
        "invoice_id", "customer_id", "customer_name", "billing_period",
        "amount_billed", "invoice_date", "status", "notes",
    ]

    def _write(path: Path, rows: list[dict], fields: list[str]) -> None:
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)

    _write(c_path, contracts, contract_fields)
    _write(b_path, billing,   billing_fields)

    print(f"Generated {len(contracts)} contract rows -> {c_path.name}")
    print(f"Generated {len(billing)} billing rows   -> {b_path.name}")
    print(f"Samples written to: {out}")
    return c_path, b_path


if __name__ == "__main__":
    generate()
