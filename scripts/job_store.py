"""
output/job_store.json is the single source of truth: one entry per job ever
seen, keyed by job id, never overwritten wholesale (rank_jobs.py and
tailor_selected.py both merge into it). The ranked spreadsheet is always a
full render of the store, so there's no separate "merge with existing sheet
rows" logic to get wrong.
"""
import contextlib
import datetime
import json
import os
import subprocess
import sys
import time

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

STORE_PATH = "output/job_store.json"
SHEET_PATH = "output/tailored_resumes.xlsx"
HTML_PATH = "output/dashboard.html"
STORE_LOCK_PATH = "output/.store.lock"
STORE_LOCK_TIMEOUT_SECONDS = 30
STORE_LOCK_POLL_SECONDS = 0.2


class StoreLockTimeout(Exception):
    pass

HEADERS = ["Rank", "Fit", "Est. TC (USD)", "Company", "Role", "Location",
           "Likely 2027?", "Start Signal", "Fit Reason", "TC Basis",
           "Tailor Depth", "Status", "Job Link", "Resume PDF", "levels.fyi",
           "Date Scored"]


def load_store(path=STORE_PATH):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_store(store, path=STORE_PATH):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(store, f, indent=2)


@contextlib.contextmanager
def _store_lock(lock_path=STORE_LOCK_PATH):
    """Short-held mutex around the read-merge-write critical section in
    merge_and_save(), so concurrent writers (the scheduled scan, an ad-hoc
    add_job.py/tailor_selected.py run) never race on that one operation.
    os.O_CREAT | O_EXCL is an atomic create-if-absent on both POSIX and
    Windows -- no check-then-create gap like os.path.exists() + open()
    would have. Held for milliseconds, not the run's whole duration, so a
    short poll-and-retry is fine (unlike run_pipeline.ps1's separate
    whole-run lock, which is a "skip if busy" check, not a wait).
    """
    deadline = time.time() + STORE_LOCK_TIMEOUT_SECONDS
    fd = None
    while fd is None:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            if time.time() > deadline:
                raise StoreLockTimeout(
                    f"Could not acquire {lock_path} after "
                    f"{STORE_LOCK_TIMEOUT_SECONDS}s -- another process may be "
                    "stuck holding it. Delete it by hand if you're sure it's stale."
                )
            time.sleep(STORE_LOCK_POLL_SECONDS)
    try:
        yield
    finally:
        os.close(fd)
        os.remove(lock_path)


def merge_and_save(updates, path=STORE_PATH, lock_path=STORE_LOCK_PATH):
    """Merge `updates` (job_id -> entry, only the jobs THIS process actually
    touched) into whatever is currently on disk, instead of overwriting with
    a stale full snapshot from this process's own load_store() at
    start-of-run. This is what makes it safe to run rank_jobs.py /
    tailor_selected.py / add_job.py concurrently: each only needs to agree
    on the handful of entries it changed, not the other's. Returns the
    merged store for immediate use (e.g. write_outputs()).

    `path`/`lock_path` let a separate pipeline (e.g. the internship one) use
    its own store file with its own lock, so the two never contend or cross-
    contaminate each other's data.
    """
    with _store_lock(lock_path):
        current = load_store(path)
        current.update(updates)
        save_store(current, path)
        return current


def blend(fit, tc_hi):
    tc_norm = min(tc_hi, 300000) / 300000 * 100
    return 0.7 * fit + 0.3 * tc_norm


def _tc_display(tc_lo, tc_hi, tc_display):
    if not tc_hi:
        return "unknown"
    if tc_display == "hourly":
        return f"${tc_lo}–${tc_hi}/hr"
    return f"${tc_lo // 1000}k–${tc_hi // 1000}k"


def write_sheet(store, path=SHEET_PATH, tc_display="annual"):
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
        tc_disp = _tc_display(tc_lo, tc_hi, tc_display)
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

    wb.save(path)
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


def write_html(store, path=HTML_PATH, heading="Ranked Jobs", tc_display="annual"):
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
        tc_disp = _tc_display(tc_lo, tc_hi, tc_display)
        fit = j.get("fit_score", 0)
        job_link = (f'<a href="{_esc(j["url"])}" target="_blank">Apply ↗</a>'
                    if j.get("url") else "")
        pdf_link = (f'<a href="{_esc(j["pdf_rel"])}" target="_blank">PDF ↗</a>'
                    if j.get("pdf_rel") else "")
        levels = ("https://www.levels.fyi/?compare=" +
                  _esc(str(j.get("company", "")).replace(" ", "%20")) +
                  "&track=Software%20Engineer")
        trs.append(f"""<tr data-id="{_esc(j.get('id', ''))}">
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
  <td style="text-align:center"><input type="checkbox" class="applied-cb"></td>
</tr>""")

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Resume Tailor — {_esc(heading)}</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 24px; background: #fafafa; }}
  h1 {{ font-size: 18px; color: #333; margin-bottom: 4px; }}
  .meta {{ color: #777; font-size: 13px; margin-bottom: 14px; }}
  .refresh-btn {{ background: #2F5496; color: #fff; border: none; border-radius: 4px;
    padding: 4px 10px; font-size: 13px; cursor: pointer; margin-left: 8px; }}
  .refresh-btn:hover {{ background: #244275; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; background: #fff; }}
  th, td {{ border: 1px solid #ddd; padding: 6px 8px; text-align: left; vertical-align: top; }}
  th {{ background: #2F5496; color: #fff; position: sticky; top: 0; }}
  tr:hover {{ background: #f0f4fa; }}
  a {{ color: #0563C1; text-decoration: none; }}
  a:hover {{ text-decoration: underline; }}
  tr.applied-row {{ opacity: 0.5; background: #f5f5f5; }}
</style>
</head>
<body>
<h1>{_esc(heading)}</h1>
<div class="meta">{len(rows)} jobs · generated {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} · check "Applied" to sink a row to the bottom (kept, never deleted) <button class="refresh-btn" onclick="location.reload()">⟳ Refresh</button></div>
<table>
<tr><th>Rank</th><th>Fit</th><th>Est. TC</th><th>Company</th><th>Role</th><th>Location</th>
<th>Likely 2027?</th><th>Fit Reason</th><th>Status</th><th>Job</th><th>Resume</th><th>Comp</th><th>Scored</th><th>Applied</th></tr>
{''.join(trs)}
</table>
<script>
(function() {{
  var PREFIX = 'resumeTailorApplied:';
  var table = document.querySelector('table');
  function isApplied(id) {{ return localStorage.getItem(PREFIX + id) === '1'; }}
  function setApplied(id, val) {{
    if (val) localStorage.setItem(PREFIX + id, '1');
    else localStorage.removeItem(PREFIX + id);
  }}
  function reorder() {{
    var rows = Array.prototype.slice.call(table.querySelectorAll('tr[data-id]'));
    rows.sort(function(a, b) {{
      var aApplied = a.classList.contains('applied-row') ? 1 : 0;
      var bApplied = b.classList.contains('applied-row') ? 1 : 0;
      return aApplied - bApplied;
    }});
    rows.forEach(function(r) {{ table.appendChild(r); }});
  }}
  var rows = table.querySelectorAll('tr[data-id]');
  rows.forEach(function(row) {{
    var id = row.getAttribute('data-id');
    var cb = row.querySelector('.applied-cb');
    if (!id || !cb) return;
    if (isApplied(id)) {{
      cb.checked = true;
      row.classList.add('applied-row');
    }}
    cb.addEventListener('change', function() {{
      setApplied(id, cb.checked);
      row.classList.toggle('applied-row', cb.checked);
      reorder();
    }});
  }});
  reorder();
}})();
</script>
</body>
</html>
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)


def write_outputs(store, sheet_path=SHEET_PATH, html_path=HTML_PATH, heading="Ranked Jobs",
                   tc_display="annual"):
    """Writes both the xlsx (portable snapshot) and the auto-refreshing
    HTML dashboard (meant to be left open) from the same store.

    tc_display="hourly" formats Est. TC as a $/hr band instead of $Xk-$Yk --
    for the internship pipeline, whose estimates are hourly, not annual."""
    n = write_sheet(store, sheet_path, tc_display)
    write_html(store, html_path, heading, tc_display)
    return n


def commit_and_push(message):
    """Best-effort commit+push of output/ changes. Ad-hoc scripts
    (add_job.py, tailor_selected.py) run outside run_pipeline.ps1, which
    only commits its own work -- without this, any ad-hoc session leaves
    the working tree dirty, and 'git pull --rebase' in the next scheduled
    run refuses outright ("You have unstaged changes"), silently breaking
    the pipeline until someone notices and commits by hand (this actually
    happened). Never raises -- a failed commit/push here shouldn't crash an
    otherwise-successful run; it just means the next pull surfaces it.
    """
    try:
        subprocess.run(["git", "add", "output/"], check=True, capture_output=True)
        staged = subprocess.run(["git", "diff", "--cached", "--quiet"])
        if staged.returncode == 0:
            return False  # nothing to commit
        subprocess.run(["git", "commit", "-m", message], check=True, capture_output=True)
        subprocess.run(["git", "push"], check=True, capture_output=True)
        return True
    except subprocess.CalledProcessError as e:
        stderr = (e.stderr or b"").decode("utf-8", errors="replace")[:300]
        print(f"WARNING: git commit/push failed: {stderr}", file=sys.stderr)
        return False


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
