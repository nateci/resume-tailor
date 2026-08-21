#!/usr/bin/env python3
"""
Scans the local, signed-in Outlook mailbox (via COM automation -- classic
Outlook only, New Outlook dropped COM/VBA support entirely) for
job-application emails -- confirmations, OA invites, interview requests,
rejections, offers -- and tags matching entries in job_store.json /
intern_job_store.json with an app_status field.

Only matches against jobs already marked status=="tailored" (i.e. postings
Nate actually tailored a resume for and presumably applied to), not the
full scored backlog -- otherwise a generic company-name match against
400+ untailored postings would produce a lot of false positives.

Requires: Outlook desktop installed, set to classic (not "New Outlook" --
toggle in the top-right of the Outlook window) and running, plus pywin32
(`pip install pywin32`). No Azure app registration / API access needed --
this drives the already-signed-in desktop client directly.
"""
import datetime
import os
import sys

import win32com.client

LOOKBACK_DAYS = 240
OL_MAIL_ITEM = 43  # olMailItem

STORES = [
    ("output/job_store.json", "output/tailored_resumes.xlsx", "output/dashboard.html",
     "output/applied_tracker.html", "output/tracker_imports_newgrad.json",
     "Ranked Jobs", "annual"),
    ("output/intern_job_store.json", "output/intern_tailored_resumes.xlsx",
     "output/intern_dashboard.html", "output/intern_applied_tracker.html",
     "output/tracker_imports_intern.json",
     "Ranked Internships (targeting Spring 2027)", "hourly"),
]

# Checked in this priority order (offer/rejected are the most unambiguous
# signals, so they're checked before the vaguer interview/OA/applied ones).
OFFER_KW = ("pleased to offer", "excited to offer", "extend an offer",
            "offer letter", "offer of employment")
REJECT_KW = ("regret to inform", "not moving forward", "will not be moving forward",
             "decided not to move forward", "other candidates", "position has been filled",
             "not selected", "unfortunately")
OA_KW = ("online assessment", "coding challenge", "coding assessment", "hackerrank",
         "codesignal", "karat", "assessment invite", "complete the assessment")
INTERVIEW_KW = ("interview", "technical screen", "phone screen", "schedule a call",
                "schedule a time", "move forward with your application")
APPLIED_KW = ("thank you for applying", "application received", "successfully applied",
              "received your application", "thanks for your interest",
              "thank you for your interest")


def classify(text):
    t = text.lower()
    if any(k in t for k in OFFER_KW):
        return "offer"
    if any(k in t for k in REJECT_KW):
        return "rejected"
    if any(k in t for k in OA_KW):
        return "oa"
    if any(k in t for k in INTERVIEW_KW):
        return "interview"
    if any(k in t for k in APPLIED_KW):
        return "applied"
    return None


def fetch_messages():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # olFolderInbox
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # newest first
    since = datetime.datetime.now() - datetime.timedelta(days=LOOKBACK_DAYS)
    restrict_str = since.strftime("%m/%d/%Y %I:%M %p")
    items = items.Restrict(f"[ReceivedTime] >= '{restrict_str}'")

    messages = []
    for m in items:
        try:
            if m.Class != OL_MAIL_ITEM:
                continue
            messages.append({
                "subject": m.Subject or "",
                "sender_name": m.SenderName or "",
                "sender_address": m.SenderEmailAddress or "",
                "body": m.Body or "",
                "received": m.ReceivedTime,
            })
        except Exception:
            continue  # skip items COM chokes on (rare malformed/legacy items)
    return messages


def main():
    print("Reading Outlook inbox via COM...", file=sys.stderr)
    messages = fetch_messages()
    print(f"{len(messages)} message(s) in the last {LOOKBACK_DAYS} days", file=sys.stderr)

    sys.path.insert(0, os.path.dirname(__file__))
    from job_store import load_store, save_store, write_outputs

    for store_path, sheet_path, html_path, tracker_path, imports_path, heading, tc_display in STORES:
        store = load_store(store_path)
        tailored = {jid: j for jid, j in store.items() if j.get("status") == "tailored"}
        matched_jobs = set()

        if not tailored:
            print(f"{store_path}: no tailored jobs yet, nothing to match against", file=sys.stderr)
        else:
            by_company = {}
            for jid, j in tailored.items():
                by_company.setdefault(j["company"].lower(), []).append(jid)

            # Messages arrive newest-first, so the first hit per job is
            # already the most recent applicable email -- once a job has an
            # app_status this run, later (older) messages about it are ignored.
            for m in messages:
                haystack = f"{m['subject']} {m['sender_name']} {m['sender_address']}".lower()
                company_hit = next((c for c in by_company if c in haystack), None)
                if not company_hit:
                    continue
                status = classify(f"{m['subject']} {m['body'][:2000]}")
                if not status:
                    continue

                for jid in by_company[company_hit]:
                    if jid in matched_jobs:
                        continue
                    store[jid]["app_status"] = status
                    store[jid]["app_status_date"] = str(m["received"])
                    matched_jobs.add(jid)

            if matched_jobs:
                save_store(store, store_path)

        # Always regenerate the tracker, even with 0 new matches this run --
        # earlier runs' app_status, or imported sheet rows, may already be
        # there and the tracker file might not exist yet.
        write_outputs(store, sheet_path=sheet_path, html_path=html_path,
                      tracker_path=tracker_path, imports_path=imports_path,
                      heading=heading, tc_display=tc_display)
        print(f"{store_path}: {len(matched_jobs)} job(s) tagged with an app_status",
              file=sys.stderr)


if __name__ == "__main__":
    main()
