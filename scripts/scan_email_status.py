#!/usr/bin/env python3
"""
Scans the local, signed-in Outlook mailbox (via COM automation -- classic
Outlook only, New Outlook dropped COM/VBA support entirely) for
job-application emails -- confirmations, OA invites, interview requests,
rejections, offers -- and tracks them automatically, in two tiers:

1. If the email's company matches a job_store entry marked status==
   "tailored" or applied=True (scripts/mark_applied.py), that entry gets
   app_status/app_status_date set directly -- richer data (real title,
   location, resume PDF link) since it's a known, already-vetted posting.
   Deliberately NOT matched against the full scored backlog: a company
   often has many scored-but-never-applied-to postings in job_store, and
   blanket-matching would flag every one of them from a single email
   about the one job actually applied to.
2. Otherwise, the company is extracted straight from the email itself
   (subject-line patterns, then sender display name, then sender domain)
   and upserted into a small sidecar file (output/tracker_auto_*.json)
   that write_tracker_html merges in alongside job_store and the Sheets
   import -- so a posting applied to without ever running
   tailor_selected.py (Nate mostly skips that step -- takes too long)
   still gets tracked, just with less detail than tier 1.

Only a message classify() recognizes as an actual application-lifecycle
email (offer/rejection/OA/interview/applied-confirmation wording) is
considered at all -- an email merely mentioning a known company's name
isn't enough by itself.

Hard floor: CUTOFF_DATE ignores anything received before it entirely, so
a wide net (or the first full LOOKBACK_DAYS bootstrap) can't drag in
stale rejections/applications from a prior semester's search.

Incremental: only fetches mail received since the last successful run
(tracked in .last_email_scan, plus a small OVERLAP_MINUTES safety margin),
not a full LOOKBACK_DAYS re-scan every time -- that full sweep only
happens once, to bootstrap a fresh checkout. Delete .last_email_scan to
force a full re-scan (e.g. after changing the classify()/extraction rules
and wanting them applied retroactively).

Requires: Outlook desktop installed, set to classic (not "New Outlook" --
toggle in the top-right of the Outlook window) and running, plus pywin32
(`pip install pywin32`). No Azure app registration / API access needed --
this drives the already-signed-in desktop client directly.
"""
import datetime
import hashlib
import json
import os
import re
import sys

import win32com.client

LOOKBACK_DAYS = 240  # only used to bootstrap the very first run
OVERLAP_MINUTES = 60  # re-check a bit of overlap past the last run as a safety margin
LAST_RUN_PATH = ".last_email_scan"
OL_MAIL_ITEM = 43  # olMailItem
# Hard floor -- last semester's search is over; ignore anything older so a
# wide net (or a first full LOOKBACK_DAYS bootstrap) can't drag in stale
# rejections/applications from a prior cycle.
CUTOFF_DATE = datetime.datetime(2026, 6, 1)

# store_path, sheet_path, html_path, tracker_path, imports_path (Sheets
# import), auto_path (emails for companies not in job_store at all)
STORES = [
    ("output/job_store.json", "output/tailored_resumes.xlsx", "output/dashboard.html",
     "output/applied_tracker.html", "output/tracker_imports_newgrad.json",
     "output/tracker_auto_newgrad.json", "Ranked Jobs", "annual"),
    ("output/intern_job_store.json", "output/intern_tailored_resumes.xlsx",
     "output/intern_dashboard.html", "output/intern_applied_tracker.html",
     "output/tracker_imports_intern.json", "output/tracker_auto_intern.json",
     "Ranked Internships (targeting Spring 2027)", "hourly"),
]

# Checked in this priority order (offer/rejected are the most unambiguous
# signals, so they're checked before the vaguer interview/OA/applied ones).
# Deliberately avoid single common words here (bare "interview",
# "unfortunately") -- those matched newsletters, university digests, and a
# security newsletter's article titles in testing. Every phrase below is
# specific enough that it's very unlikely to appear outside a real
# application-lifecycle email.
OFFER_KW = ("pleased to offer", "excited to offer", "extend an offer",
            "offer letter", "offer of employment")
REJECT_KW = ("regret to inform", "not moving forward", "will not be moving forward",
             "decided not to move forward", "other candidates", "position has been filled",
             "not selected")
OA_KW = ("online assessment", "coding challenge", "coding assessment", "hackerrank",
         "codesignal", "karat", "assessment invite", "complete the assessment",
         "invites you to take an assessment", "invitation for assessments",
         "assessment invitation")
INTERVIEW_KW = ("invite you to interview", "invited to interview", "would like to interview",
                 "schedule your interview", "advance to the interview", "interview invitation",
                 "prepare for your interview", "technical screen", "phone screen",
                 "schedule a call", "schedule a time", "move forward with your application")
APPLIED_KW = ("thank you for applying", "thank you for your application",
              "application received", "successfully applied", "successfully submitted",
              "received your application", "thanks for your interest",
              "thank you for your interest", "application confirmation",
              "employment application")

INTERN_KW = ("intern", "internship", "co-op", "coop")

# Subject-line patterns for extracting a company name, tried in order.
# Deliberately conservative (non-greedy, cut at |/,/!/. ) -- a wrong
# extraction pollutes the tracker with a garbage row, so patterns that
# only fire on clearly job-application-shaped subjects are preferred over
# looser ones that might grab a role title or generic phrase instead.
COMPANY_SUBJECT_PATTERNS = [
    r"thank you for applying (?:to|at) (?P<c>[^|,!.]+)",
    r"thank you for your application to (?P<c>[^|,!.]+)",
    r"thank you for applying for .+? at (?P<c>[^|,!.]+)",
    r"submitted your (?P<c>[^|,!.]+) job application",
    r"received your (?P<c>[^|,!.]+) application",
    r"^your (?P<c>[^|,!.]+) application\b",
    r"^(?P<c>[^|,!.]+) application confirmation\b",
    r"^(?P<c>[^|,!.]+) application update\b",
    r"^(?P<c>[^|,!.]+) application\s*[:\-]",
    r"^(?P<c>[^|,!.]+)\s*-\s*in response to your application",
    r"^(?P<c>[^|,!.]+)\s*-\s*assessment invitation",
    r"^your application to (?P<c>[^|,!.]+)\s*$",
    r"^(?P<c>[^|,!.]+) invites you to take an assessment",
    r"^(?P<c>[^|,!.]+) employment application",
]

GENERIC_SENDER_NAMES = {
    "greenhouse", "lever", "workday", "coderbyte", "ashby", "ashbyhq",
    "bamboohr", "smartrecruiters", "taleo", "jobvite", "icims", "pinpoint",
    "indeed", "linkedin",
}
GENERIC_ATS_DOMAINS = {
    "greenhouse-mail.io", "greenhouse.io", "ashbyhq.com", "hire.lever.co",
    "lever.co", "myworkday.com", "workday.com", "myworkdayjobs.com",
    "icims.com", "bamboohr.com", "app.bamboohr.com", "smartrecruiters.com",
    "taleo.net", "jobvite.com", "coderbyte.com", "pinpoint.email",
    "successfactors.com", "ultipro.com", "breezy.hr",
    "symplicity.com",  # university career-portal weekly job-digest bulk mail
    "hackerrankmail.com",  # HackerRank's own marketing mail, not an OA invite
}
DOMAIN_STRIP_LABELS = ("careers", "jobs", "talent", "email", "mail", "notifications",
                        "no-reply", "noreply", "recruiting", "hr", "postmaster")


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


def _clean_company(c):
    c = c.strip().strip("!.,")
    c = re.sub(r",?\s*Nate$", "", c, flags=re.I).strip()
    c = re.sub(r",?\s*Inc\.?$", "", c, flags=re.I).strip()
    c = re.sub(r",?\s*Incorporated$", "", c, flags=re.I).strip()
    return c


def _company_from_subject(subject):
    for pattern in COMPANY_SUBJECT_PATTERNS:
        m = re.search(pattern, subject, re.IGNORECASE)
        if m:
            c = _clean_company(m.group("c"))
            if len(c) >= 2:
                return c
    return None


def _company_from_sender_name(sender_name, sender_address):
    if not sender_name or sender_name.strip().lower() == (sender_address or "").strip().lower():
        return None
    n = sender_name.strip()
    n = re.sub(r"^careers at\s+", "", n, flags=re.I)
    n = re.sub(r"\s*(talent acquisition|recruiting team|hiring team|careers|talent|hr|assessments?|notifications?)\s*$",
               "", n, flags=re.I).strip()
    if len(n) < 2 or n.lower() in GENERIC_SENDER_NAMES:
        return None
    # A recruiter emailing directly (not through an ATS) shows their own
    # name as the sender -- that's a real person, not the company, and the
    # domain (handled by the caller's fallback) gives the real answer instead.
    if re.match(r"^[A-Z][a-zA-Z'-]+\s+[A-Z][a-zA-Z'-]+$", n):
        return None
    return n


def _company_from_domain(sender_address):
    if "@" not in (sender_address or ""):
        return None
    domain = sender_address.split("@", 1)[1].lower()
    parts = domain.split(".")
    while parts and parts[0] in DOMAIN_STRIP_LABELS:
        parts = parts[1:]
    core_domain = ".".join(parts)
    if domain in GENERIC_ATS_DOMAINS or core_domain in GENERIC_ATS_DOMAINS or not parts:
        return None
    label = re.sub(r"(hr|careers|jobs|talent|recruiting)$", "", parts[0], flags=re.I) or parts[0]
    if len(label) < 3:
        return None
    return label.capitalize()


def extract_company(subject, sender_name, sender_address):
    return (_company_from_subject(subject)
            or _company_from_sender_name(sender_name, sender_address)
            or _company_from_domain(sender_address))


def guess_intern(subject, body):
    t = f"{subject} {body[:500]}".lower()
    return any(k in t for k in INTERN_KW)


def load_last_run():
    if os.path.exists(LAST_RUN_PATH):
        try:
            return datetime.datetime.fromisoformat(open(LAST_RUN_PATH, encoding="utf-8").read().strip())
        except ValueError:
            pass
    return datetime.datetime.now() - datetime.timedelta(days=LOOKBACK_DAYS)


def save_last_run(when):
    with open(LAST_RUN_PATH, "w", encoding="utf-8") as f:
        f.write(when.isoformat())


def fetch_messages(since):
    since = max(since, CUTOFF_DATE)  # never look earlier than last semester's cutoff
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # olFolderInbox
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # newest first
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


def _word_match(company_lower, haystack):
    if len(company_lower) < 3:
        return False  # too short to substring-match safely (e.g. a 2-letter company)
    return re.search(r"(?<!\w)" + re.escape(company_lower) + r"(?!\w)", haystack) is not None


def auto_row_key(company):
    return "auto:" + hashlib.md5(company.lower().encode("utf-8")).hexdigest()[:16]


# Different emails about the SAME application often phrase the company name
# differently ("Optiver" vs "Optiver Assessments", "Anduril" vs "Anduril
# Industries") -- stripping these common corporate/ATS suffixes before
# comparing normalizes most of that away, so updates stack onto one row
# instead of fragmenting into several ("stacking via company is fine").
CORP_SUFFIX_RE = re.compile(
    r"\s*,?\s*(industries|incorporated|inc\.?|llc|corp\.?|corporation|group|"
    r"company|systems|technologies|technology|labs|laboratories|"
    r"assessments?|notifications?|careers?|talent|recruiting|hr)\.?\s*$",
    re.IGNORECASE)


def normalize_company(company):
    c = (company or "").strip()
    prev = None
    while prev != c:
        prev = c
        c = CORP_SUFFIX_RE.sub("", c).strip()
    return re.sub(r"[^a-z0-9]", "", c.lower())


def find_existing_auto_key(rows, company):
    target = normalize_company(company)
    if not target:
        return None
    for key, row in rows.items():
        if normalize_company(row.get("company", "")) == target:
            return key
    return None


def build_auto_row(existing, company, status, received, url=None):
    existing = existing or {}
    date_sort = str(received)[:10]
    return {
        "id": existing.get("id") or auto_row_key(company),
        "date_sort": date_sort,
        "date_display": datetime.datetime.strptime(date_sort, "%Y-%m-%d").strftime("%m-%d-%Y")
                         if date_sort else "",
        "company": existing.get("company") or company,  # keep the first-seen display name
        "title": existing.get("title", ""),
        "location": existing.get("location", ""),
        "status": status,
        "url": url or existing.get("url", ""),
        "contact": existing.get("contact", ""),
        "resume_ver": existing.get("resume_ver", ""),
        "interview_dates": existing.get("interview_dates", ""),
        "notes": existing.get("notes", ""),
    }


def load_auto_rows(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return {r["id"]: r for r in json.load(f)}
    return {}


def save_auto_rows(path, rows):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(list(rows.values()), f, indent=2)


def main():
    run_start = datetime.datetime.now()
    last_run = load_last_run()
    since = max(last_run - datetime.timedelta(minutes=OVERLAP_MINUTES), CUTOFF_DATE)

    print("Reading Outlook inbox via COM...", file=sys.stderr)
    messages = fetch_messages(since)
    print(f"{len(messages)} message(s) since {since.strftime('%Y-%m-%d %H:%M')} "
          f"(last scan: {last_run.strftime('%Y-%m-%d %H:%M')})", file=sys.stderr)

    sys.path.insert(0, os.path.dirname(__file__))
    from job_store import load_store, save_store, write_outputs

    stores = {}
    # store_path -> {company_lower: [jid, ...]}, restricted to status=="tailored"
    # or applied==True -- NOT the full scored backlog. A company often has many
    # scored-but-never-applied-to postings in job_store; matching against all of
    # them would blanket-flag every one as applied/rejected/etc. from a single
    # email about the one job actually applied to.
    company_index = {}
    for store_path, *_rest in STORES:
        store = load_store(store_path)
        stores[store_path] = store
        idx = {}
        for jid, j in store.items():
            if j.get("status") == "tailored" or j.get("applied"):
                idx.setdefault(j.get("company", "").lower(), []).append(jid)
        company_index[store_path] = idx

    dirty = set()
    matched_ids = set()  # id(m) of messages already claimed by a job_store match
    updated_jids = set()  # jids already set THIS run -- messages arrive newest-first,
                           # so the first hit per job is already the most recent one;
                           # without this, an older message about the same job later
                           # in the loop would clobber a newer, more-advanced status.
    tagged_counts = {sp: 0 for sp, *_ in STORES}

    for m in messages:
        text = f"{m['subject']} {m['body'][:2000]}"
        status = classify(text)
        if not status:
            continue
        haystack = f"{m['subject']} {m['sender_name']} {m['sender_address']}".lower()

        hit = None
        for store_path, idx in company_index.items():
            for company_lower, jids in idx.items():
                if _word_match(company_lower, haystack):
                    hit = (store_path, jids)
                    break
            if hit:
                break

        if hit:
            store_path, jids = hit
            matched_ids.add(id(m))
            new_jids = [jid for jid in jids if jid not in updated_jids]
            if not new_jids:
                continue  # every job for this company already got a newer status this run
            for jid in new_jids:
                job = stores[store_path][jid]
                job["applied"] = True
                job["app_status"] = status
                job["app_status_date"] = str(m["received"])
                updated_jids.add(jid)
            dirty.add(store_path)
            tagged_counts[store_path] += 1

    auto_rows = {STORES[0][5]: load_auto_rows(STORES[0][5]),
                 STORES[1][5]: load_auto_rows(STORES[1][5])}
    auto_dirty = set()
    unmatched = []
    updated_auto_keys = set()  # same newest-first "first hit wins" concern as updated_jids
    for m in messages:
        if id(m) in matched_ids:
            continue
        text = f"{m['subject']} {m['body'][:2000]}"
        status = classify(text)
        if not status:
            continue
        company = extract_company(m["subject"], m["sender_name"], m["sender_address"])
        if not company:
            unmatched.append(m["subject"])
            continue
        is_intern = guess_intern(m["subject"], m["body"])
        target_store = STORES[1] if is_intern else STORES[0]
        auto_path = target_store[5]
        rows = auto_rows[auto_path]

        key = find_existing_auto_key(rows, company)
        if key and key in updated_auto_keys:
            continue  # already got a newer status this run
        row = build_auto_row(rows.get(key), company, status, m["received"])
        rows[row["id"]] = row
        updated_auto_keys.add(row["id"])
        auto_dirty.add(auto_path)
        tagged_counts[target_store[0]] += 1

    for auto_path in auto_dirty:
        save_auto_rows(auto_path, auto_rows[auto_path])

    for store_path in dirty:
        save_store(stores[store_path], store_path)

    for store_path, sheet_path, html_path, tracker_path, imports_path, auto_path, heading, tc_display in STORES:
        # Always regenerate the tracker, even with 0 new matches this run --
        # earlier runs' data, or Sheets/auto rows, may already be there and
        # the tracker file might not exist yet.
        write_outputs(stores[store_path], sheet_path=sheet_path, html_path=html_path,
                      tracker_path=tracker_path, imports_path=imports_path,
                      auto_path=auto_path, heading=heading, tc_display=tc_display)
        print(f"{store_path}: {tagged_counts[store_path]} job(s) tagged with an app_status this run",
              file=sys.stderr)

    if unmatched:
        print(f"{len(unmatched)} application-shaped email(s) couldn't be attributed to a company "
              f"(skipped, not added anywhere):", file=sys.stderr)
        for s in unmatched:
            print(f"  - {s}", file=sys.stderr)

    save_last_run(run_start)


if __name__ == "__main__":
    main()
