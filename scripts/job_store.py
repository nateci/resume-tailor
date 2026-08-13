"""
output/job_store.json is the single source of truth: one entry per job ever
seen, keyed by job id, never overwritten wholesale (rank_jobs.py and
tailor_selected.py both merge into it). The ranked spreadsheet is always a
full render of the store, so there's no separate "merge with existing sheet
rows" logic to get wrong.
"""
import json
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

STORE_PATH = "output/job_store.json"
SHEET_PATH = "output/tailored_resumes.xlsx"

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
