#!/usr/bin/env python3
"""
Ad-hoc single-job add: given a URL (+ title/company), fetch the description,
score it, fully tailor it, and add it to the ranked sheet — all in one shot.
For a posting found outside the SimplifyJobs feed (a link someone sent you,
something you found yourself) rather than waiting for the next scan.

Usage:
  python scripts/add_job.py --url "https://..." --title "Software Engineer" --company "Acme"
  python scripts/add_job.py --url "..." --title "..." --company "..." --locations "Boston, MA"
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import sys
import urllib.request

from claude_client import resolve_claude_cmd, call_claude
from job_store import load_store, save_store, write_outputs, commit_and_push
from fetch_descriptions import fetch_one, extract_start_signal, UA
from rank_jobs import SCORE_SCHEMA, score_prompt
from tailor_selected import RESUME_TEX, PDF_DIR, tag_for, tailor_one

PROFILE = "resume/profile.json"


def probe_metadata(url):
    """Best-effort auto-detect (title, company, locations) from known ATS
    job-board APIs, so a bare link is usually enough without retyping
    title/company by hand. Returns None for any field it can't determine."""

    def get_json(u):
        req = urllib.request.Request(u, headers=UA)
        with urllib.request.urlopen(req, timeout=15) as r:
            return json.loads(r.read().decode("utf-8"))

    m = re.search(r"ashbyhq\.com/([\w-]+)/([\w-]+)", url)
    if m:
        org, jid = m.groups()
        try:
            data = get_json(f"https://api.ashbyhq.com/posting-api/job-board/{org}")
            for job in data.get("jobs", []):
                if job.get("id") == jid or jid in json.dumps(job):
                    return (job.get("title"), org.replace("-", " ").title(),
                            job.get("locationName") or job.get("location", ""))
        except Exception:
            pass
        return (None, org.replace("-", " ").title(), None)

    m = re.search(r"greenhouse\.io/([\w-]+)/jobs/(\d+)", url)
    if m:
        org, jid = m.groups()
        try:
            data = get_json(f"https://boards-api.greenhouse.io/v1/boards/{org}/jobs/{jid}")
            loc = (data.get("location") or {}).get("name", "")
            return (data.get("title"), org.replace("-", " ").title(), loc)
        except Exception:
            pass
        return (None, org.replace("-", " ").title(), None)

    m = re.search(r"lever\.co/([\w-]+)/([\w-]+)", url)
    if m:
        org, jid = m.groups()
        try:
            data = get_json(f"https://api.lever.co/v0/postings/{org}/{jid}")
            loc = (data.get("categories") or {}).get("location", "")
            return (data.get("text"), org.replace("-", " ").title(), loc)
        except Exception:
            pass
        return (None, org.replace("-", " ").title(), None)

    return (None, None, None)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True)
    ap.add_argument("--title", default=None,
                     help="Auto-detected from the ATS API when possible (Ashby/Greenhouse/Lever)")
    ap.add_argument("--company", default=None,
                     help="Auto-detected from the ATS API when possible (Ashby/Greenhouse/Lever)")
    ap.add_argument("--locations", default=None)
    ap.add_argument("--description-file", default=None,
                     help="Skip auto-fetch and use this file's contents as the job "
                          "description instead — for sites that don't fetch cleanly "
                          "(bot walls, JS-rendered SPAs) where you've pasted the real "
                          "page text by hand.")
    ap.add_argument("--resume", default=RESUME_TEX,
                     help="Base .tex file to tailor from (default: resume/resume.tex). "
                          "Use resume/resume_intern.tex for internship postings.")
    args = ap.parse_args()

    cmd_prefix = resolve_claude_cmd()
    if not cmd_prefix:
        print("claude CLI not found on PATH.", file=sys.stderr)
        sys.exit(1)

    if not args.title or not args.company:
        print("Probing ATS API for title/company...", file=sys.stderr)
        auto_title, auto_company, auto_loc = probe_metadata(args.url)
        args.title = args.title or auto_title
        args.company = args.company or auto_company
        args.locations = args.locations or auto_loc
        if args.title or args.company:
            print(f"  detected: {args.company or '?'} — {args.title or '?'}", file=sys.stderr)

    missing = [name for name, val in (("--title", args.title), ("--company", args.company)) if not val]
    if missing:
        print(f"Could not auto-detect {', '.join(missing)} for this URL — "
              f"pass {' and '.join(missing)} explicitly.", file=sys.stderr)
        sys.exit(1)

    if args.description_file:
        with open(args.description_file, encoding="utf-8") as f:
            desc = f.read().strip()
        source = "manual-paste"
        print(f"Using pasted description ({len(desc)} chars) — skipping auto-fetch.", file=sys.stderr)
    else:
        print("Fetching description...", file=sys.stderr)
        desc, source = fetch_one(args.url)
        if desc:
            print(f"  got {len(desc)} chars via {source}", file=sys.stderr)
        else:
            print("  could not fetch a description — proceeding with title+company only", file=sys.stderr)
    start_signal, likely_2027 = extract_start_signal(desc)

    job_id = "manual:" + hashlib.sha1(args.url.encode("utf-8")).hexdigest()[:12]
    job = {
        "id": job_id,
        "company": args.company,
        "title": args.title,
        "url": args.url,
        "locations": [args.locations] if args.locations else [],
        "terms": [],
        "sponsorship": "",
        "date_posted": "",
        "description": desc,
        "desc_source": source,
        "start_signal": start_signal,
        "likely_2027": likely_2027,
    }

    with open(PROFILE, encoding="utf-8") as f:
        profile = json.load(f)

    today = datetime.date.today().isoformat()

    print("Scoring...", file=sys.stderr)
    try:
        res = call_claude(cmd_prefix, score_prompt(profile, job), SCORE_SCHEMA)
        job["fit_score"] = int(res.get("fit_score", 0))
        job["fit_reason"] = res.get("fit_reason", "")
        job["tc_estimate_low"] = int(res.get("tc_estimate_low", 0))
        job["tc_estimate_high"] = int(res.get("tc_estimate_high", 0))
        job["tc_basis"] = res.get("tc_basis", "")
    except Exception as e:
        print(f"  scoring failed: {e}", file=sys.stderr)
        job["fit_score"] = 0
        job["status"] = f"score-error:{str(e)[:60]}"
    job["date_scored"] = today

    print(f"Tailoring (base: {args.resume})...", file=sys.stderr)
    os.makedirs(PDF_DIR, exist_ok=True)
    with open(args.resume, encoding="utf-8") as f:
        base_tex = f.read()
    tag = tag_for(job)
    try:
        pdf_rel, status = tailor_one(cmd_prefix, base_tex, job, tag)
        job["pdf_rel"] = pdf_rel
        job["status"] = status
        if status == "tailored" or status.startswith("too-long:"):
            job["date_tailored"] = today
    except Exception as e:
        job["status"] = f"tailor-error:{str(e)[:60]}"
        job["pdf_rel"] = ""

    store = load_store()
    store[job_id] = job
    save_store(store)
    n = write_outputs(store)

    print(f"\nFit: {job.get('fit_score')} — {job.get('fit_reason', '')}", file=sys.stderr)
    print(f"Status: {job['status']}", file=sys.stderr)
    if job.get("pdf_rel"):
        print(f"PDF: {job['pdf_rel']}", file=sys.stderr)
    print(f"Sheet now has {n} total rows.", file=sys.stderr)
    if commit_and_push(f"Ad-hoc: added {job['company']} — {job['title']}"):
        print("Committed and pushed.", file=sys.stderr)


if __name__ == "__main__":
    main()
