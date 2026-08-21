#!/usr/bin/env python3
"""
Internship counterpart to filter_jobs.py: reads SimplifyJobs' internship
listings.json (Summer2027-Internships repo -- despite the name, it covers
all terms: Winter/Spring/Summer/Fall) -> keeps active SWE / systems / AI-infra
internships in roughly the 2027 window -> writes output/intern_new_jobs.json.

Fully separate state from the new-grad pipeline (own seen-ids file, own
new-jobs file) so the two scans never interact.

Unlike filter_jobs.py, this does NOT skip processing on the very first run.
filter_jobs.py's first-run skip exists so a fresh deployment doesn't dump an
entire backlog into the tailoring queue at once -- but this pipeline's whole
point on its first run is to surface the current backlog of Spring-2027-ish
internships, so "new" == "everything currently kept and not yet seen", which
on a first run is just everything currently kept.
"""
import json
import os
import sys
import urllib.request

from filter_jobs import (
    WANT_TITLE_KEYWORDS, SKIP_TITLE_KEYWORDS, SKIP_COMPANY_KEYWORDS,
    http_get, get, is_active, normalize,
)

PRIMARY_URL = (
    "https://raw.githubusercontent.com/SimplifyJobs/"
    "Summer2027-Internships/dev/.github/scripts/listings.json"
)
TREE_API = (
    "https://api.github.com/repos/SimplifyJobs/"
    "Summer2027-Internships/git/trees/dev?recursive=1"
)
RAW_BASE = "https://raw.githubusercontent.com/SimplifyJobs/Summer2027-Internships/dev/"

SEEN_PATH = "output/intern_seen_ids.json"
NEW_JOBS_PATH = "output/intern_new_jobs.json"

# Internship terms roughly compatible with a Jan-2027-through-graduation
# availability window (see resume/profile_intern.json). Terms outside this
# set aren't hard-excluded if unspecified -- only excluded when the posting
# is unambiguously for a different year.
TARGET_YEAR = "2027"


def load_listings():
    try:
        return json.loads(http_get(PRIMARY_URL))
    except Exception as e:
        print(f"Primary listings path failed ({e}); locating via tree API", file=sys.stderr)
    tree = json.loads(http_get(TREE_API))
    cand = [n["path"] for n in tree.get("tree", []) if n["path"].endswith("listings.json")]
    if not cand:
        raise RuntimeError("Could not locate listings.json in repo tree")
    path = sorted(cand, key=len)[0]
    print(f"Found listings.json at {path}", file=sys.stderr)
    return json.loads(http_get(RAW_BASE + path))


def wanted(job):
    company = str(get(job, "company_name", "company")).lower()
    if any(k in company for k in SKIP_COMPANY_KEYWORDS):
        return False
    title = str(get(job, "title", "role", "position")).lower()
    if any(k in title for k in SKIP_TITLE_KEYWORDS):
        return False
    return any(k in title for k in WANT_TITLE_KEYWORDS)


def term_in_target_window(job):
    """Keep if any listed term mentions 2027, or if no usable term info is
    given (let the scoring step reason about timing from the description
    instead). Reject only when terms are given and are unambiguously NOT
    2027 (e.g. only "Summer 2026" or "Spring 2028")."""
    terms = get(job, "terms", default=[])
    if isinstance(terms, str):
        terms = [terms]
    if not terms:
        return True
    terms_l = [str(t).lower().strip() for t in terms if t]
    if not terms_l or all(t in ("", "n/a") for t in terms_l):
        return True
    return any(TARGET_YEAR in t for t in terms_l)


def main():
    listings = load_listings()
    if isinstance(listings, dict):
        listings = listings.get("jobs", listings.get("listings", []))

    kept = [normalize(j) for j in listings
            if is_active(j) and wanted(j) and term_in_target_window(j)]
    print(f"{len(kept)} active SWE/systems/AI internships in the ~2027 window", file=sys.stderr)

    seen = set()
    if os.path.exists(SEEN_PATH):
        try:
            seen = set(json.load(open(SEEN_PATH)))
        except Exception:
            seen = set()

    new_jobs = [j for j in kept if j["id"] not in seen]

    all_ids = sorted({j["id"] for j in kept} | seen)
    os.makedirs("output", exist_ok=True)
    json.dump(all_ids, open(SEEN_PATH, "w"), indent=0)
    json.dump(new_jobs, open(NEW_JOBS_PATH, "w"), indent=2)

    print(f"{len(new_jobs)} NEW internship(s) to process", file=sys.stderr)
    print(len(new_jobs))


if __name__ == "__main__":
    main()
