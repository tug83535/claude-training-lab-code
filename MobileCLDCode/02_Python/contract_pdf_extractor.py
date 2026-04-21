"""
contract_pdf_extractor.py - Bulk PDF Contract Term Extractor

PURPOSE
-------
Scan a folder of signed customer/vendor contracts (PDF) and extract the key
business terms into a structured table:

  - Parties (customer / vendor)
  - Effective date & term length
  - Renewal notice period
  - Auto-renewal flag
  - Annual contract value
  - Payment terms (Net X)
  - SLA / uptime commitments
  - Governing law / jurisdiction
  - Termination for convenience?
  - Data processing addendum present?

WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
---------------------------------------
Neither Excel nor OneDrive can read a 40-page PDF and return structured data.
OneDrive search finds the file; it does NOT tell you what's in it. Adobe's
"Export to Excel" is per-file, manual, and often mangles tables.

USE CASE
--------
Legal team has 800 active customer contracts across 3 acquired companies in
SharePoint. Build a clean master list for the renewal forecast in 20 minutes
instead of 3 weeks.

USAGE
-----
    python contract_pdf_extractor.py /path/to/contracts/ --output contracts.xlsx

REQUIREMENTS
------------
    pip install pdfplumber python-dateutil pandas openpyxl
"""
from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, asdict
from pathlib import Path

import pandas as pd

try:
    import pdfplumber
except ImportError as e:
    raise SystemExit("pip install pdfplumber") from e


# --- Regex library (tuned, not exhaustive) -----------------------------------

RE_EFFECTIVE = re.compile(
    r"(?:effective date|commencement date|start date)[\s:]+([A-Z][a-z]+ \d{1,2},? \d{4})",
    re.IGNORECASE,
)
RE_TERM_YEARS = re.compile(
    r"(?:initial term|subscription term|contract term)[^.]*?(\d+)[\s-]*(?:year|yr)",
    re.IGNORECASE,
)
RE_TERM_MONTHS = re.compile(
    r"(?:initial term|subscription term)[^.]*?(\d+)[\s-]*(?:month|mo)",
    re.IGNORECASE,
)
RE_NOTICE = re.compile(
    r"(?:written notice|notice of (?:non-?)?renewal)[^.]*?(\d+)\s*(?:day|days)",
    re.IGNORECASE,
)
RE_AUTO_RENEW = re.compile(
    r"auto(?:matically)?[\s-]*renew|evergreen|successive renewal", re.IGNORECASE
)
RE_VALUE = re.compile(
    r"(?:total (?:contract )?value|annual (?:contract )?(?:value|fee)|ACV)[^.]*?\$\s?([\d,]+(?:\.\d{2})?)",
    re.IGNORECASE,
)
RE_NET_TERMS = re.compile(r"\bNet\s*(\d{2,3})\b")
RE_SLA = re.compile(r"(\d{2}\.\d+)\s*%?\s*(?:uptime|availability)", re.IGNORECASE)
RE_GOV_LAW = re.compile(
    r"governed by(?: the)? laws?(?: of)?(?: the (?:state|commonwealth) of)?\s+"
    r"([A-Z][a-zA-Z ]{2,25})",
)
RE_TERM_CONV = re.compile(
    r"terminat\w+ for convenience|either party may terminate[^.]*?convenience",
    re.IGNORECASE,
)
RE_DPA = re.compile(
    r"data processing (?:addendum|agreement)|DPA", re.IGNORECASE
)


@dataclass
class ContractTerms:
    filename: str
    effective_date: str | None
    term_length: str | None
    notice_days: int | None
    auto_renew: bool
    annual_value_usd: float | None
    net_terms_days: int | None
    sla_uptime_pct: float | None
    governing_law: str | None
    terminate_for_convenience: bool
    dpa_present: bool
    warnings: str


def extract_text(path: Path) -> str:
    try:
        with pdfplumber.open(path) as pdf:
            return "\n".join((page.extract_text() or "") for page in pdf.pages)
    except Exception as e:  # noqa: BLE001 - we want to continue scanning
        return f"__EXTRACT_ERROR__: {e}"


def extract_terms(text: str, filename: str) -> ContractTerms:
    warnings = []
    if text.startswith("__EXTRACT_ERROR__"):
        return ContractTerms(
            filename=filename, effective_date=None, term_length=None,
            notice_days=None, auto_renew=False, annual_value_usd=None,
            net_terms_days=None, sla_uptime_pct=None, governing_law=None,
            terminate_for_convenience=False, dpa_present=False,
            warnings=text,
        )

    # Effective date
    m = RE_EFFECTIVE.search(text)
    effective = m.group(1) if m else None

    # Term length
    term = None
    my = RE_TERM_YEARS.search(text)
    mm = RE_TERM_MONTHS.search(text)
    if my:
        term = f"{my.group(1)} year(s)"
    elif mm:
        term = f"{mm.group(1)} month(s)"
    else:
        warnings.append("term length not found")

    m = RE_NOTICE.search(text)
    notice_days = int(m.group(1)) if m else None

    auto_renew = bool(RE_AUTO_RENEW.search(text))

    # Value
    annual_value = None
    m = RE_VALUE.search(text)
    if m:
        try:
            annual_value = float(m.group(1).replace(",", ""))
        except ValueError:
            pass

    m = RE_NET_TERMS.search(text)
    net_days = int(m.group(1)) if m else None

    m = RE_SLA.search(text)
    sla = float(m.group(1)) if m else None

    m = RE_GOV_LAW.search(text)
    gov_law = m.group(1).strip() if m else None

    term_for_conv = bool(RE_TERM_CONV.search(text))
    dpa = bool(RE_DPA.search(text))

    return ContractTerms(
        filename=filename,
        effective_date=effective,
        term_length=term,
        notice_days=notice_days,
        auto_renew=auto_renew,
        annual_value_usd=annual_value,
        net_terms_days=net_days,
        sla_uptime_pct=sla,
        governing_law=gov_law,
        terminate_for_convenience=term_for_conv,
        dpa_present=dpa,
        warnings="; ".join(warnings),
    )


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--output", default="contracts.xlsx")
    ap.add_argument("--recursive", action="store_true", default=True)
    args = ap.parse_args()

    folder = Path(args.folder)
    pattern = "**/*.pdf" if args.recursive else "*.pdf"
    pdfs = sorted(folder.glob(pattern))
    if not pdfs:
        raise SystemExit(f"No PDFs found in {folder}")

    rows = []
    for i, pdf in enumerate(pdfs, 1):
        text = extract_text(pdf)
        rows.append(asdict(extract_terms(text, pdf.name)))
        if i % 25 == 0:
            print(f"  processed {i}/{len(pdfs)}")

    df = pd.DataFrame(rows)
    df.to_excel(args.output, index=False, engine="openpyxl")

    print(f"Wrote {args.output}")
    print(f"Files scanned: {len(df)}")
    print(f"With warnings: {(df['warnings'] != '').sum()}")
    print(f"Auto-renewing contracts: {df['auto_renew'].sum()}")
    if df["annual_value_usd"].notna().any():
        print(f"Total ACV detected: ${df['annual_value_usd'].sum():,.0f}")


if __name__ == "__main__":
    main()
