"""
output/job_store.json is the single source of truth: one entry per job ever
seen, keyed by job id, never overwritten wholesale (rank_jobs.py and
tailor_selected.py both merge into it). The ranked spreadsheet is always a
full render of the store, so there's no separate "merge with existing sheet
rows" logic to get wrong.
"""
import datetime
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

STORE_PATH = "output/job_store.json"
SHEET_PATH = "output/tailored_resumes.xlsx"
HTML_PATH = "output/dashboard.html"
HTML_REFRESH_SECONDS = 20

HEADERS = ["Rank", "Fit", "Est. TC (USD)", "Company", "Role", "Location",
           "Likely 2027?", "Start Signal", "Fit Reason", "TC Basis",
           "Tailor Depth", "Status", "Job Link", "Resume PDF", "levels.fyi",
           "Date Scored"]


def load_store():
    if os.path.exists(STORE_PATH):
        with open(STORE_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_store(store):
    with open(STORE_PATH, "w", encoding="utf-8") as f:
        json.dump(store, f, indent=2)


def blend(fit, tc_hi):
    tc_norm = min(tc_hi, 300000) / 300000 * 100
    return 0.7 * fit + 0.3 * tc_norm


def write_sheet(store):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ranked Jobs"
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(1, i, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5496")
    widths = [5, 5, 16, 20, 34, 22, 12, 26, 40, 28, 13, 14, 10, 30, 12, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    rows = sorted(
        store.values(),
        key=lambda j: blend(j.get("fit_score", 0), j.get("tc_estimate_high", 0)),
        reverse=True,
    )
    for rank, j in enumerate(rows, 1):
        loc = j.get("locations")
        loc = ", ".join(loc) if isinstance(loc, list) else str(loc or "")
        tc_lo, tc_hi = j.get("tc_estimate_low", 0), j.get("tc_estimate_high", 0)
        tc_disp = f"${tc_lo // 1000}k–${tc_hi // 1000}k" if tc_hi else "unknown"
        lv = ("https://www.levels.fyi/?compare=" +
              str(j.get("company", "")).replace(" ", "%20") + "&track=Software%20Engineer")
        row = rank + 1
        vals = [rank, j.get("fit_score", 0), tc_disp, j.get("company", ""), j.get("title", ""),
                loc, "yes" if j.get("likely_2027") else "", j.get("start_signal", ""),
                j.get("fit_reason", ""), j.get("tc_basis", ""),
                "full" if j.get("description") else "metadata-only",
                j.get("status", "scored"),
                "Apply ↗" if j.get("url") else "",
                os.path.basename(j["pdf_rel"]) if j.get("pdf_rel") else "",
                "levels.fyi ↗", j.get("date_scored", "")]
        for c, v in enumerate(vals, 1):
            ws.cell(row, c, v)
        if j.get("url"):
            _link(ws.cell(row, 13), j["url"])
        if j.get("pdf_rel"):
            _link(ws.cell(row, 14), j["pdf_rel"])
        _link(ws.cell(row, 15), lv)
        _color_fit(ws.cell(row, 2))

    wb.save(SHEET_PATH)
    return len(rows)


def _esc(s):
    return (str(s if s is not None else "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def _fit_color(v):
    try:
        v = int(v)
    except Exception:
        return "#FFC7CE"
    return "#C6EFCE" if v >= 75 else "#FFEB9C" if v >= 50 else "#FFC7CE"


def write_html(store):
    """Auto-refreshing local dashboard — meant to be left open in a browser
    tab. Unlike the xlsx, browsers don't take an exclusive lock on the file
    they're displaying, so this can stay open indefinitely without ever
    blocking the pipeline's next write."""
    rows = sorted(
        store.values(),
        key=lambda j: blend(j.get("fit_score", 0), j.get("tc_estimate_high", 0)),
        reverse=True,
    )

    trs = []
    for rank, j in enumerate(rows, 1):
        loc = j.get("locations")
        loc = ", ".join(loc) if isinstance(loc, list) else str(loc or "")
        tc_lo, tc_hi = j.get("tc_estimate_low", 0), j.get("tc_estimate_high", 0)
        tc_disp = f"${tc_lo // 1000}k–${tc_hi // 1000}k" if tc_hi else "unknown"
        fit = j.get("fit_score", 0)
        job_link = (f'<a href="{_esc(j["url"])}" target="_blank">Apply ↗</a>'
                    if j.get("url") else "")
        pdf_link = (f'<a href="{_esc(j["pdf_rel"])}" target="_blank">PDF ↗</a>'
                    if j.get("pdf_rel") else "")
        levels = ("https://www.levels.fyi/?compare=" +
                  _esc(str(j.get("company", "")).replace(" ", "%20")) +
                  "&track=Software%20Engineer")
        trs.append(f"""<tr>
  <td>{rank}</td>
  <td style="background:{_fit_color(fit)}">{fit}</td>
  <td>{_esc(tc_disp)}</td>
  <td>{_esc(j.get('company', ''))}</td>
  <td>{_esc(j.get('title', ''))}</td>
  <td>{_esc(loc)}</td>
  <td>{'yes' if j.get('likely_2027') else ''}</td>
  <td>{_esc(j.get('fit_reason', ''))}</td>
  <td>{_esc(j.get('status', ''))}</td>
  <td>{job_link}</td>
  <td>{pdf_link}</td>
  <td><a href="{levels}" target="_blank">levels.fyi ↗</a></td>
  <td>{_esc(j.get('date_scored', ''))}</td>
</tr>""")

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="refresh" content="{HTML_REFRESH_SECONDS}">
<title>Resume Tailor — Ranked Jobs</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 24px; background: #fafafa; }}
  h1 {{ font-size: 18px; color: #333; margin-bottom: 4px; }}
  .meta {{ color: #777; font-size: 13px; margin-bottom: 14px; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; background: #fff; }}
  th, td {{ border: 1px solid #ddd; padding: 6px 8px; text-align: left; vertical-align: top; }}
  th {{ background: #2F5496; color: #fff; position: sticky; top: 0; }}
  tr:hover {{ background: #f0f4fa; }}
  a {{ color: #0563C1; text-decoration: none; }}
  a:hover {{ text-decoration: underline; }}
</style>
</head>
<body>
<h1>Ranked Jobs</h1>
<div class="meta">{len(rows)} jobs · auto-refreshes every {HTML_REFRESH_SECONDS}s · generated {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
<table>
<tr><th>Rank</th><th>Fit</th><th>Est. TC</th><th>Company</th><th>Role</th><th>Location</th>
<th>Likely 2027?</th><th>Fit Reason</th><th>Status</th><th>Job</th><th>Resume</th><th>Comp</th><th>Scored</th></tr>
{''.join(trs)}
</table>
</body>
</html>
"""
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)


def write_outputs(store):
    """Writes both the xlsx (portable snapshot) and the auto-refreshing
    HTML dashboard (meant to be left open) from the same store."""
    n = write_sheet(store)
    write_html(store)
    return n


def _link(cell, target):
    cell.hyperlink = target
    cell.font = Font(color="0563C1", underline="single")


def _color_fit(cell):
    v = cell.value or 0
    try:
        v = int(v)
    except Exception:
        return
    color = "C6EFCE" if v >= 75 else "FFEB9C" if v >= 50 else "FFC7CE"
    cell.fill = PatternFill("solid", fgColor=color)
