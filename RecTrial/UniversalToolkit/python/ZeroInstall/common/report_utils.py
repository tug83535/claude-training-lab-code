# Finance Automation Toolkit v1.0 — iPipeline
# common/report_utils.py — iPipeline-branded HTML report assembly
#
# Brand spec: primary #0B4779 (iPipeline Blue), navy #112E51, lime #BFF18C,
#             aqua #2BCCD3, arctic white #F9F9F9, charcoal #161616, Arial only.

from datetime import datetime

_CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; background: #F9F9F9; color: #161616; font-size: 14px; }
.hdr { background: #0B4779; color: #fff; padding: 20px 32px; }
.hdr h1 { font-size: 22px; font-weight: bold; }
.hdr .sub { font-size: 13px; color: #BFF18C; margin-top: 5px; }
.body { padding: 24px 32px; }
.metric-row { display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }
.card { background: #fff; border: 1px solid #dde; border-radius: 4px; padding: 16px 20px; min-width: 130px; }
.card .val { font-size: 28px; font-weight: bold; color: #0B4779; }
.card .lbl { font-size: 11px; color: #666; margin-top: 4px; text-transform: uppercase; letter-spacing: .4px; }
.card.ok .val  { color: #1e7e1e; }
.card.bad .val { color: #a00; }
.card.warn .val{ color: #b05000; }
h2 { color: #0B4779; border-bottom: 2px solid #0B4779; padding-bottom: 5px; margin: 28px 0 10px; font-size: 16px; }
h3 { color: #112E51; margin: 18px 0 6px; font-size: 14px; }
p  { margin: 6px 0; line-height: 1.5; }
table { width: 100%; border-collapse: collapse; background: #fff; margin-top: 8px; font-size: 13px; }
th { background: #0B4779; color: #fff; padding: 8px 12px; text-align: left; font-weight: bold; }
tr:nth-child(even) { background: #edf2f7; }
td { padding: 7px 12px; border-bottom: 1px solid #e0e0e0; vertical-align: top; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: bold; }
.badge-ok   { background: #BFF18C; color: #1a4a1a; }
.badge-bad  { background: #ffd0d0; color: #700; }
.badge-warn { background: #ffecc0; color: #7a4000; }
.badge-info { background: #d0e8ff; color: #0B4779; }
.footer { background: #112E51; color: #99b; font-size: 11px; padding: 10px 32px; margin-top: 40px; }
.section-note { background: #fff8e1; border-left: 4px solid #2BCCD3; padding: 10px 14px; margin: 12px 0; font-size: 13px; }
"""


def build_report(title: str, subtitle: str, sections: list[str]) -> str:
    """Assemble a complete iPipeline-branded HTML report.

    Args:
        title:    page title and header H1
        subtitle: smaller text below the header (e.g. file name + run timestamp)
        sections: list of HTML strings produced by the helpers below
    """
    body = "\n".join(sections)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    return (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n'
        "<head>"
        '<meta charset="UTF-8">'
        f"<title>{title}</title>"
        f"<style>{_CSS}</style>"
        "</head>\n"
        "<body>\n"
        f'<div class="hdr"><h1>{title}</h1><div class="sub">{subtitle}</div></div>\n'
        f'<div class="body">{body}</div>\n'
        f'<div class="footer">Finance Automation Toolkit v1.0 &nbsp;|&nbsp; iPipeline &nbsp;|&nbsp; {ts}</div>\n'
        "</body></html>"
    )


def metric_row(cards: list[dict]) -> str:
    """Render a horizontal row of metric cards.

    Each card dict: {label, value, status} where status is 'normal'|'ok'|'bad'|'warn'.
    """
    html = []
    for c in cards:
        cls = {"ok": "ok", "bad": "bad", "warn": "warn"}.get(c.get("status", "normal"), "")
        html.append(
            f'<div class="card {cls}">'
            f'<div class="val">{c["value"]}</div>'
            f'<div class="lbl">{c["label"]}</div>'
            f"</div>"
        )
    return '<div class="metric-row">' + "".join(html) + "</div>"


def data_table(heading: str, headers: list[str], rows: list[list],
               status_col: int | None = None) -> str:
    """Render a section heading + data table.

    status_col: if set, applies pass/fail/warn badge styling to that column index.
    """
    if not rows:
        return f"<h2>{heading}</h2><p>No records.</p>"
    th = "".join(f"<th>{h}</th>" for h in headers)
    trs = []
    for row in rows:
        cells = []
        for i, val in enumerate(row):
            s = str(val) if val is not None else ""
            if status_col is not None and i == status_col:
                up = s.upper()
                if up in ("PASS", "OK", "CLEAN", "NONE"):
                    cells.append(f'<td><span class="badge badge-ok">{s}</span></td>')
                elif up in ("FAIL", "MISSING", "ERROR", "HIGH"):
                    cells.append(f'<td><span class="badge badge-bad">{s}</span></td>')
                elif up in ("WARN", "WARNING", "REVIEW", "MEDIUM", "STALE"):
                    cells.append(f'<td><span class="badge badge-warn">{s}</span></td>')
                else:
                    cells.append(f'<td><span class="badge badge-info">{s}</span></td>')
            else:
                cells.append(f"<td>{s}</td>")
        trs.append("<tr>" + "".join(cells) + "</tr>")
    return (
        f"<h2>{heading}</h2>"
        f"<table><tr>{th}</tr>"
        + "".join(trs)
        + "</table>"
    )


def note_box(text: str) -> str:
    """Render a highlighted note box (aqua left border)."""
    return f'<div class="section-note">{text}</div>'


def badge(text: str, status: str = "info") -> str:
    """Inline badge: status is 'ok'|'bad'|'warn'|'info'."""
    cls = f"badge badge-{status}"
    return f'<span class="{cls}">{text}</span>'
