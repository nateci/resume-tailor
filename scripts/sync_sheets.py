#!/usr/bin/env python3
"""
Pushes jobs with an Outlook-confirmed app_status ("interacted with" --
applied, OA, interview, rejected, or offer; set by scan_email_status.py)
into Nate's existing manual Google Sheets trackers, via a small Apps Script
Web App bound to each sheet (source: scripts/apps_script/tracker_sync.gs).
There's no OAuth on this end -- the Apps Script runs with Nate's own
Google permissions once he deploys it.

Idempotent: each row this creates is tagged with the job's id in the
sheet's hidden column K. On repeat runs the Apps Script looks up that id
to update the Status cell in place instead of duplicating the row. If a
company already has a manual row (no column K tag), it's left alone
entirely and reported back here instead of guessing it's the same
application.

Setup: see scripts/apps_script/tracker_sync.gs's header comment, then set
  SHEET_WEBHOOK_URL_NEWGRAD, SHEET_WEBHOOK_URL_INTERN, SHEET_SYNC_TOKEN
"""
import datetime
import json
import os
import sys
import urllib.request

NEWGRAD_WEBHOOK = os.environ.get("SHEET_WEBHOOK_URL_NEWGRAD")
INTERN_WEBHOOK = os.environ.get("SHEET_WEBHOOK_URL_INTERN")
SYNC_TOKEN = os.environ.get("SHEET_SYNC_TOKEN")

STATUS_TEXT = {
    "applied": "Applied",
    "oa": "OA Sent",
    "interview": "Phone Screen",
    "rejected": "Rejected",
    "offer": "Offer",
}

STORES = [
    ("output/job_store.json", NEWGRAD_WEBHOOK, "new-grad"),
    ("output/intern_job_store.json", INTERN_WEBHOOK, "intern"),
]


def fmt_date(raw):
    if not raw:
        return ""
    raw = str(raw).split("+")[0].strip()
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.datetime.strptime(raw, fmt).strftime("%m-%d-%Y")
        except ValueError:
            continue
    return ""


def build_payload(store):
    jobs = []
    for jid, j in store.items():
        status = j.get("app_status")
        if not status:
            continue
        loc = j.get("locations")
        loc = ", ".join(loc) if isinstance(loc, list) else str(loc or "")
        jobs.append({
            "id": jid,
            "company": j.get("company", ""),
            "title": j.get("title", ""),
            "location": loc,
            "url": j.get("url", ""),
            "status": STATUS_TEXT.get(status, status),
            "date_applied": fmt_date(j.get("app_status_date", "")),
        })
    return jobs


def post(url, payload):
    body = json.dumps({"token": SYNC_TOKEN, "jobs": payload}).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def main():
    if not SYNC_TOKEN:
        sys.exit("Set SHEET_SYNC_TOKEN (must match SYNC_TOKEN in both Apps Scripts).")

    sys.path.insert(0, os.path.dirname(__file__))
    from job_store import load_store

    for store_path, webhook, label in STORES:
        if not webhook:
            print(f"{label}: no webhook URL set, skipping", file=sys.stderr)
            continue
        store = load_store(store_path)
        payload = build_payload(store)
        if not payload:
            print(f"{label}: no interacted-with jobs to sync", file=sys.stderr)
            continue
        result = post(webhook, payload)
        if "error" in result:
            print(f"{label}: {result['error']}", file=sys.stderr)
            continue
        print(f"{label}: created {result.get('created', 0)}, "
              f"updated {result.get('updated', 0)}, "
              f"skipped (existing manual row) {result.get('skipped', 0)}", file=sys.stderr)
        for c in result.get("skipped_companies", []):
            print(f"  - skipped, already tracked manually: {c}", file=sys.stderr)


if __name__ == "__main__":
    main()
