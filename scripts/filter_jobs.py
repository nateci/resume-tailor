#!/usr/bin/env python3
"""
Read SimplifyJobs listings.json -> keep active, full-time SWE / systems /
AI-infra roles -> diff against seen state so only NEW jobs flow downstream.

Does NOT filter on TC or grad-year. Those are handled later:
  - TC: estimated + used only for ranking (never a hard cut)
  - grad-year / start date: best-effort scraped from the posting, flagged not cut

State: output/seen_ids.json (committed back so "new since last run" survives
the machine being off between runs).
"""
import json
import os
import sys
import urllib.request

# listings.json is the machine-readable source the README rows are generated
# from. Primary path first; if it 404s we locate it via the git tree API.
PRIMARY_URL = (
    "https://raw.githubusercontent.com/SimplifyJobs/"
    "New-Grad-Positions/dev/.github/scripts/listings.json"
)
TREE_API = (
    "https://api.github.com/repos/SimplifyJobs/"
    "New-Grad-Positions/git/trees/dev?recursive=1"
)
RAW_BASE = "https://raw.githubusercontent.com/SimplifyJobs/New-Grad-Positions/dev/"

SEEN_PATH = "output/seen_ids.json"
NEW_JOBS_PATH = "output/new_jobs.json"

WANT_TITLE_KEYWORDS = (
    "software", "swe", "developer", "engineer", "backend", "back-end",
    "back end", "distributed", "systems", "platform", "infrastructure",
    "infra", "sre", "reliability", "data engineer", "machine learning",
    " ml ", "ml ", "ai ", "ai/", "mlops", "full stack", "fullstack",
)
# Titles we skip outright — clearly off-profile.
SKIP_TITLE_KEYWORDS = (
    "hardware", "fpga", "asic", "electrical", "mechanical", "product manager",
    "program manager", "designer", "ux ", "ui ", "recruiter", "sales",
    "marketing", "clearance", "quant",  # quant excluded per SWE+Data/AI scope
)


def http_get(url, headers=None, timeout=30):
    req = urllib.request.Request(url, headers=headers or {"User-Agent": "resume-tailor/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8")


def load_listings():
    try:
        return json.loads(http_get(PRIMARY_URL))
    except Exception as e:
        print(f"Primary listings path failed ({e}); locating via tree API", file=sys.stderr)
    # Fallback: find listings.json anywhere in the repo tree
    tree = json.loads(http_get(TREE_API))
    cand = [n["path"] for n in tree.get("tree", []) if n["path"].endswith("listings.json")]
    if not cand:
        raise RuntimeError("Could not locate listings.json in repo tree")
    # Prefer the shortest path (usually the canonical one)
    path = sorted(cand, key=len)[0]
    print(f"Found listings.json at {path}", file=sys.stderr)
    return json.loads(http_get(RAW_BASE + path))


def get(d, *keys, default=""):
    for k in keys:
        if k in d and d[k] not in (None, ""):
            return d[k]
    return default


def is_active(job):
    active = job.get("active", True)
    visible = job.get("is_visible", job.get("visible", True))
    return bool(active) and bool(visible)


def is_fulltime(job):
    terms = get(job, "terms", default=[])
    if isinstance(terms, list):
        terms = " ".join(str(t) for t in terms)
    terms = str(terms).lower()
    # Reject anything clearly an internship/co-op term.
    if any(w in terms for w in ("intern", "co-op", "coop", "summer")):
        return False
    return True


def wanted(job):
    title = str(get(job, "title", "role", "position")).lower()
    if any(k in title for k in SKIP_TITLE_KEYWORDS):
        return False
    return any(k in title for k in WANT_TITLE_KEYWORDS)


def job_id(job):
    jid = get(job, "id", "uuid")
    if jid:
        return str(jid)
    return f"{get(job,'company_name','company')}|{get(job,'title','role')}|{get(job,'url','application_link')}"


def normalize(job):
    return {
        "id": job_id(job),
        "company": get(job, "company_name", "company", default="Unknown"),
        "title": get(job, "title", "role", "position", default="Unknown Role"),
        "url": get(job, "url", "application_link", "apply_url", "link"),
        "locations": get(job, "locations", "location", default=[]),
        "terms": get(job, "terms", default=[]),
        "sponsorship": get(job, "sponsorship", default=""),
        "date_posted": get(job, "date_posted", "date_updated", default=""),
    }


def main():
    listings = load_listings()
    if isinstance(listings, dict):
        listings = listings.get("jobs", listings.get("listings", []))

    kept = [normalize(j) for j in listings
            if is_active(j) and is_fulltime(j) and wanted(j)]
    print(f"{len(kept)} active full-time SWE/systems/AI roles in feed", file=sys.stderr)

    seen = set()
    if os.path.exists(SEEN_PATH):
        try:
            seen = set(json.load(open(SEEN_PATH)))
        except Exception:
            seen = set()

    first_run = len(seen) == 0
    new_jobs = [] if first_run else [j for j in kept if j["id"] not in seen]
    if first_run:
        print("First run: seeding seen_ids without tailoring the backlog.", file=sys.stderr)

    all_ids = sorted({j["id"] for j in kept} | seen)
    os.makedirs("output", exist_ok=True)
    json.dump(all_ids, open(SEEN_PATH, "w"), indent=0)
    json.dump(new_jobs, open(NEW_JOBS_PATH, "w"), indent=2)

    print(f"{len(new_jobs)} NEW job(s) to process", file=sys.stderr)
    print(len(new_jobs))


if __name__ == "__main__":
    main()
