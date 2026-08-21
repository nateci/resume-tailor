#!/usr/bin/env python3
"""
One-time/rerunnable import of Nate's pre-existing manually-tracked
applications from his two Google Sheets trackers into the local
application-tracker HTMLs (see job_store.write_tracker_html), so history
predating scan_email_status.py isn't lost. The sheets must be shared as
"Anyone with the link -> Viewer" -- this fetches the CSV export directly,
no Google auth needed.

Writes output/tracker_imports_newgrad.json / _intern.json (normalized row
dicts) which write_tracker_html merges in automatically on every future
regeneration -- rerun this script any time the sheets get new manual rows.

Dedup: an imported row is skipped if the corresponding job_store already
has an Outlook-detected app_status for the same company (case-insensitive)
dated within 3 days -- treated as the same application already captured.
Sheet company abbreviations that don't match a full job_store company name
(e.g. "db" for Databricks, "spcx" for SpaceX) won't be caught by this and
may show up as a near-duplicate row -- use the tracker HTML's per-row hide
button to clean those up by hand.
"""
import csv
import datetime
import hashlib
import io
import json
import os
import sys
import urllib.request

SHEETS = [
    ("https://docs.google.com/spreadsheets/d/1FDzGjGS_zBUhw0wJm7LKGwljnNj307rr/export?format=csv&gid=241096095",
     "output/job_store.json", "output/tracker_imports_newgrad.json",
     "output/tailored_resumes.xlsx", "output/dashboard.html", "output/applied_tracker.html",
     "Ranked Jobs", "annual"),
    ("https://docs.google.com/spreadsheets/d/1-frWgpVCAdyPt-3fNC7LWVTX5EhInAXi/export?format=csv&gid=241096095",
     "output/intern_job_store.json", "output/tracker_imports_intern.json",
     "output/intern_tailored_resumes.xlsx", "output/intern_dashboard.html",
     "output/intern_applied_tracker.html",
     "Ranked Internships (targeting Spring 2027)", "hourly"),
]

STATUS_MAP = {
    "applied": "applied",
    "oa sent": "oa", "oa": "oa", "assessment": "oa", "assessment sent": "oa",
    "phone screen": "interview", "interview": "interview", "onsite": "interview",
    "technical screen": "interview", "interviewing": "interview",
    "rejected": "rejected", "declined": "rejected", "closed": "rejected",
    "offer": "offer", "offer received": "offer", "accepted": "offer",
}


def fetch_csv(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        text = r.read().decode("utf-8")
    return list(csv.reader(io.StringIO(text)))


def mdy_to_sort(mdy):
    try:
        return datetime.datetime.strptime(mdy.strip(), "%m-%d-%Y").strftime("%Y-%m-%d")
    except (ValueError, AttributeError):
        return ""


def map_status(raw):
    key = (raw or "").strip().lower()
    return STATUS_MAP.get(key, "applied"), raw


def build_rows(csv_rows):
    # Nate's sheets have a merged title row above the real header, so find
    # the header by content rather than assuming it's row 0.
    header_idx = next(
        (i for i, r in enumerate(csv_rows) if "Company" in r and "Status" in r), None)
    if header_idx is None:
        return []
    header = [h.strip() for h in csv_rows[header_idx]]

    def col(row, name, default=""):
        try:
            return row[header.index(name)].strip()
        except (ValueError, IndexError):
            return default

    rows = []
    for r in csv_rows[header_idx + 1:]:
        company = col(r, "Company")
        if not company:
            continue
        date_display = col(r, "Date Applied")
        status, orig_status = map_status(col(r, "Status"))
        notes = col(r, "Notes")
        if orig_status and orig_status.strip().lower() != status:
            notes = f"[sheet status: {orig_status}] {notes}".strip()
        url = col(r, "App Portal / URL")
        if not url.lower().startswith("http"):
            url = ""  # placeholder junk like "Go find it lazy", not a real link
        row_id = "imported:" + hashlib.md5(
            f"{company}|{col(r, 'Role / Title')}|{date_display}".encode("utf-8")).hexdigest()[:16]
        rows.append({
            "id": row_id,
            "date_sort": mdy_to_sort(date_display),
            "date_display": date_display,
            "company": company,
            "title": col(r, "Role / Title"),
            "location": col(r, "Location"),
            "status": status,
            "url": url,
            "contact": col(r, "Contact Name"),
            "resume_ver": col(r, "Resume Ver."),
            "interview_dates": col(r, "Interview Dates"),
            "notes": notes,
        })
    return rows


def dedup_against_store(rows, store):
    tagged = [(j.get("company", "").strip().lower(), str(j.get("app_status_date", ""))[:10])
              for j in store.values() if j.get("app_status")]

    def is_dup(row):
        company = row["company"].strip().lower()
        if not row["date_sort"]:
            return False
        row_date = datetime.datetime.strptime(row["date_sort"], "%Y-%m-%d")
        for tc, td in tagged:
            if tc != company or not td:
                continue
            try:
                store_date = datetime.datetime.strptime(td, "%Y-%m-%d")
            except ValueError:
                continue
            if abs((row_date - store_date).days) <= 3:
                return True
        return False

    kept, skipped = [], []
    for row in rows:
        (skipped if is_dup(row) else kept).append(row)
    return kept, skipped


def main():
    sys.path.insert(0, os.path.dirname(__file__))
    from job_store import load_store, write_outputs

    for (csv_url, store_path, imports_path, sheet_path, html_path, tracker_path,
         heading, tc_display) in SHEETS:
        print(f"Fetching {csv_url}", file=sys.stderr)
        csv_rows = fetch_csv(csv_url)
        rows = build_rows(csv_rows)
        store = load_store(store_path)
        kept, skipped = dedup_against_store(rows, store)

        with open(imports_path, "w", encoding="utf-8") as f:
            json.dump(kept, f, indent=2)
        print(f"{imports_path}: {len(kept)} imported, {len(skipped)} skipped as likely "
              f"duplicates of an Outlook-detected row", file=sys.stderr)
        for row in skipped:
            print(f"  - skipped: {row['company']} ({row['date_display']})", file=sys.stderr)

        write_outputs(store, sheet_path=sheet_path, html_path=html_path,
                      tracker_path=tracker_path, imports_path=imports_path,
                      heading=heading, tc_display=tc_display)


if __name__ == "__main__":
    main()
