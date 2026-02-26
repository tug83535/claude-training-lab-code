#!/usr/bin/env python3
"""
pnl_email_report.py — Automated Email Report Distribution
===========================================================

PURPOSE: After generating the monthly report, automatically email it to a
         distribution list with key metric summary in the body.

USAGE:
    python pnl_email_report.py --to "cfo@kbt.com,fpa@kbt.com"
    python pnl_email_report.py --to "team@kbt.com" --attach report.xlsx
    python pnl_email_report.py --preview   # Show email without sending

    from pnl_email_report import EmailReporter
    reporter = EmailReporter("KeystoneBenefitTech_PL_Model.xlsx")
    reporter.generate_and_send(recipients=["cfo@kbt.com"])
"""

import os
import sys
import argparse
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import Dict, List, Optional

import pandas as pd
import numpy as np

try:
    from pnl_config import *
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pnl_config import *


# =============================================================================
# EMAIL CONFIGURATION — CHANGE THESE
# =============================================================================

EMAIL_CONFIG = {
    "smtp_server": "smtp.office365.com",   # CHANGE THIS: your SMTP server
    "smtp_port": 587,                       # CHANGE THIS: your port (587=TLS, 465=SSL)
    "sender_email": "fpa@keystonebt.com",   # CHANGE THIS: sender address
    "sender_name": "KBT FP&A",             # CHANGE THIS: display name
    "use_tls": True,
    # Authentication — set via environment variables for security:
    #   export KBT_EMAIL_USER="your_email"
    #   export KBT_EMAIL_PASS="your_password"
}

DEFAULT_RECIPIENTS = [
    # CHANGE THESE: your actual distribution list
    "cfo@keystonebt.com",
    "fpa-team@keystonebt.com",
]


class EmailReporter(PnLBase):
    """Generates and emails P&L summary reports."""

    def __init__(self, file_path: str = None, verbose: bool = True):
        super().__init__(verbose)
        self.file_path = file_path or SOURCE_FILE
        self.gl = None
        self.metrics = {}

    def _build_metrics(self) -> Dict:
        """Compute key metrics for the email body."""
        gl = self.gl

        # Latest month with data
        latest_month = int(gl["Month"].max())
        gl_latest = gl[gl["Month"] == latest_month]

        # Prior month
        prior_month = latest_month - 1
        gl_prior = gl[gl["Month"] == prior_month] if prior_month > 0 else pd.DataFrame()

        metrics = {
            "month": MONTH_FULL[latest_month - 1] if 1 <= latest_month <= 12 else f"Month {latest_month}",
            "fiscal_year": FY_LABEL,
            "total_spend": gl_latest["Amount"].sum(),
            "total_abs": gl_latest["Abs_Amount"].sum(),
            "txn_count": len(gl_latest),
            "unique_vendors": gl_latest["Vendor"].nunique(),
            "products": {},
            "departments": {},
            "mom_change": None,
        }

        # Product breakdown
        for prod in PRODUCTS:
            prod_data = gl_latest[gl_latest["Product"] == prod]
            metrics["products"][prod] = {
                "spend": prod_data["Amount"].sum(),
                "abs_spend": prod_data["Abs_Amount"].sum(),
                "txns": len(prod_data),
            }

        # Department breakdown
        for dept in DEPARTMENTS:
            dept_data = gl_latest[gl_latest["Department"] == dept]
            metrics["departments"][dept] = dept_data["Abs_Amount"].sum()

        # MoM change
        if len(gl_prior) > 0:
            prior_total = gl_prior["Abs_Amount"].sum()
            curr_total = gl_latest["Abs_Amount"].sum()
            if prior_total > 0:
                metrics["mom_change"] = (curr_total - prior_total) / prior_total

        self.metrics = metrics
        return metrics

    def build_email_body(self, html: bool = True) -> str:
        """Build the email body with key metrics."""
        m = self.metrics

        if html:
            return self._build_html_body(m)
        return self._build_text_body(m)

    def _build_text_body(self, m: Dict) -> str:
        """Plain text email body."""
        lines = [
            f"KEYSTONE BENEFITTECH — {m['fiscal_year']} FINANCIAL SUMMARY",
            f"{'='*55}",
            f"Period: {m['month']} {FISCAL_YEAR_4}",
            f"Generated: {PnLBase.timestamp()}",
            "",
            "KEY METRICS",
            f"{'—'*40}",
            f"  Total Gross Spend:  {format_currency(m['total_abs'])}",
            f"  Total Net Spend:    {format_currency(m['total_spend'])}",
            f"  Transactions:       {m['txn_count']:,}",
            f"  Unique Vendors:     {m['unique_vendors']:,}",
        ]

        if m["mom_change"] is not None:
            direction = "▲" if m["mom_change"] > 0 else "▼"
            lines.append(f"  MoM Change:         {direction} {abs(m['mom_change']):.1%}")

        lines.extend(["", "PRODUCT BREAKDOWN", f"{'—'*40}"])
        for prod, data in m["products"].items():
            lines.append(f"  {prod:15s}  {format_currency(data['abs_spend']):>12s}  ({data['txns']:,} txns)")

        lines.extend(["", "DEPARTMENT BREAKDOWN", f"{'—'*40}"])
        sorted_depts = sorted(m["departments"].items(), key=lambda x: x[1], reverse=True)
        for dept, spend in sorted_depts:
            lines.append(f"  {dept:22s}  {format_currency(spend):>12s}")

        # Variance flags
        lines.extend(["", "VARIANCE FLAGS", f"{'—'*40}"])
        gl = self.gl
        latest = int(gl["Month"].max())
        prior = latest - 1
        flag_count = 0
        if prior > 0:
            for dept in DEPARTMENTS:
                p_spend = gl[(gl["Month"] == prior) & (gl["Department"] == dept)]["Amount"].sum()
                c_spend = gl[(gl["Month"] == latest) & (gl["Department"] == dept)]["Amount"].sum()
                if p_spend != 0:
                    pct = (c_spend - p_spend) / abs(p_spend)
                    if abs(pct) > VARIANCE_PCT:
                        arrow = "▲" if pct > 0 else "▼"
                        lines.append(f"  ⚠ {dept}: {arrow} {abs(pct):.1%} MoM ({format_currency(c_spend - p_spend)})")
                        flag_count += 1

        if flag_count == 0:
            lines.append("  ✓ All departments within threshold")

        lines.extend([
            "",
            f"{'—'*55}",
            f"This report was automatically generated by {APP_NAME} v{APP_VERSION}",
            "For questions, contact the FP&A team.",
        ])

        return "\n".join(lines)

    def _build_html_body(self, m: Dict) -> str:
        """HTML email body with formatting."""
        # Product rows
        prod_rows = ""
        for prod, data in m["products"].items():
            color = PRODUCT_COLORS.get(prod, "#808080")
            prod_rows += (
                f'<tr><td style="padding:4px 8px;border-left:3px solid {color}">{prod}</td>'
                f'<td style="padding:4px 8px;text-align:right">{format_currency(data["abs_spend"])}</td>'
                f'<td style="padding:4px 8px;text-align:right">{data["txns"]:,}</td></tr>'
            )

        # Department rows
        dept_rows = ""
        sorted_depts = sorted(m["departments"].items(), key=lambda x: x[1], reverse=True)
        for dept, spend in sorted_depts:
            dept_rows += (
                f'<tr><td style="padding:4px 8px">{dept}</td>'
                f'<td style="padding:4px 8px;text-align:right">{format_currency(spend)}</td></tr>'
            )

        mom_html = ""
        if m["mom_change"] is not None:
            color = "#C00000" if m["mom_change"] > 0 else "#00B050"
            arrow = "▲" if m["mom_change"] > 0 else "▼"
            mom_html = f'<span style="color:{color};font-weight:bold">{arrow} {abs(m["mom_change"]):.1%} MoM</span>'

        return f"""
<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto">
  <div style="background:#1F4E79;color:white;padding:16px 24px;border-radius:4px 4px 0 0">
    <h2 style="margin:0">Keystone BenefitTech — {m['fiscal_year']} Financial Summary</h2>
    <p style="margin:4px 0 0;opacity:0.85">{m['month']} {FISCAL_YEAR_4} | Generated {PnLBase.timestamp()}</p>
  </div>

  <div style="padding:20px 24px;border:1px solid #ddd;border-top:none">
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px">
      <tr>
        <td style="padding:8px;text-align:center;background:#F2F2F2;border-radius:4px">
          <div style="font-size:11px;color:#808080">GROSS SPEND</div>
          <div style="font-size:20px;font-weight:bold;color:#1F4E79">{format_currency(m['total_abs'])}</div>
        </td>
        <td style="width:12px"></td>
        <td style="padding:8px;text-align:center;background:#F2F2F2;border-radius:4px">
          <div style="font-size:11px;color:#808080">TRANSACTIONS</div>
          <div style="font-size:20px;font-weight:bold;color:#1F4E79">{m['txn_count']:,}</div>
        </td>
        <td style="width:12px"></td>
        <td style="padding:8px;text-align:center;background:#F2F2F2;border-radius:4px">
          <div style="font-size:11px;color:#808080">VENDORS</div>
          <div style="font-size:20px;font-weight:bold;color:#1F4E79">{m['unique_vendors']:,}</div>
        </td>
        <td style="width:12px"></td>
        <td style="padding:8px;text-align:center;background:#F2F2F2;border-radius:4px">
          <div style="font-size:11px;color:#808080">MoM CHANGE</div>
          <div style="font-size:20px;font-weight:bold">{mom_html or '—'}</div>
        </td>
      </tr>
    </table>

    <h3 style="color:#1F4E79;border-bottom:2px solid #4472C4;padding-bottom:4px">Product Breakdown</h3>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <tr style="background:#1F4E79;color:white">
        <th style="padding:6px 8px;text-align:left">Product</th>
        <th style="padding:6px 8px;text-align:right">Spend</th>
        <th style="padding:6px 8px;text-align:right">Transactions</th>
      </tr>
      {prod_rows}
    </table>

    <h3 style="color:#1F4E79;border-bottom:2px solid #4472C4;padding-bottom:4px;margin-top:20px">Department Breakdown</h3>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <tr style="background:#1F4E79;color:white">
        <th style="padding:6px 8px;text-align:left">Department</th>
        <th style="padding:6px 8px;text-align:right">Spend</th>
      </tr>
      {dept_rows}
    </table>
  </div>

  <div style="padding:12px 24px;background:#F2F2F2;border:1px solid #ddd;border-top:none;border-radius:0 0 4px 4px;font-size:11px;color:#808080">
    {APP_NAME} v{APP_VERSION} | Auto-generated report — do not reply
  </div>
</div>
"""

    def generate_and_send(self, recipients: List[str] = None,
                          attachments: List[str] = None,
                          subject: str = None,
                          preview: bool = False):
        """Generate metrics and send the email."""
        self.gl = self._load_gl(self.file_path)
        self._build_metrics()

        to_list = recipients or DEFAULT_RECIPIENTS
        subj = subject or f"{APP_NAME} — {self.metrics['month']} {FY_LABEL} Summary"

        text_body = self.build_email_body(html=False)
        html_body = self.build_email_body(html=True)

        if preview:
            self._section("EMAIL PREVIEW")
            self._print(f"To:      {', '.join(to_list)}")
            self._print(f"Subject: {subj}")
            self._print(f"Attach:  {attachments or 'None'}")
            print()
            print(text_body)
            return

        # Build MIME message
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subj
        msg["From"] = f"{EMAIL_CONFIG['sender_name']} <{EMAIL_CONFIG['sender_email']}>"
        msg["To"] = ", ".join(to_list)

        msg.attach(MIMEText(text_body, "plain"))
        msg.attach(MIMEText(html_body, "html"))

        # Attachments
        if attachments:
            for filepath in attachments:
                if os.path.exists(filepath):
                    with open(filepath, "rb") as f:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition",
                                    f"attachment; filename={os.path.basename(filepath)}")
                    msg.attach(part)
                    self._print(f"Attached: {filepath}", "OK")

        # Send
        try:
            user = os.environ.get("KBT_EMAIL_USER", EMAIL_CONFIG["sender_email"])
            passwd = os.environ.get("KBT_EMAIL_PASS", "")

            if not passwd:
                self._print("No password set. Set KBT_EMAIL_PASS environment variable.", "ERROR")
                self._print("Email preview generated but not sent. Use --preview to see content.", "WARN")
                return

            with smtplib.SMTP(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"]) as server:
                if EMAIL_CONFIG["use_tls"]:
                    server.starttls()
                server.login(user, passwd)
                server.sendmail(EMAIL_CONFIG["sender_email"], to_list, msg.as_string())

            self._print(f"Email sent to {len(to_list)} recipients", "OK")
        except Exception as e:
            self._print(f"Send failed: {e}", "ERROR")
            self._print("Email content was generated. Use --preview to see it.", "WARN")


def main():
    parser = argparse.ArgumentParser(description="Email P&L Summary Report")
    parser.add_argument("--file", "-f", default=SOURCE_FILE)
    parser.add_argument("--to", "-t", default=None, help="Comma-separated recipient list")
    parser.add_argument("--attach", "-a", nargs="*", help="Files to attach")
    parser.add_argument("--subject", "-s", default=None)
    parser.add_argument("--preview", "-p", action="store_true", help="Preview without sending")
    args = parser.parse_args()

    recipients = args.to.split(",") if args.to else None
    reporter = EmailReporter(file_path=args.file)
    reporter.generate_and_send(
        recipients=recipients,
        attachments=args.attach,
        subject=args.subject,
        preview=args.preview
    )


if __name__ == "__main__":
    main()
