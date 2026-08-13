#!/usr/bin/env python3
"""
Fast nightly pass: score every NEW job against the resume/profile
(fit_score, fit_reason, TC estimate) — no tailored resume text, no LaTeX,
no PDF compile. That's the slow part; it happens on demand via
tailor_selected.py once you've picked which ranked jobs you actually want.

Merges into output/job_store.json (persistent, never overwritten wholesale)
and rewrites the ranked spreadsheet from the store.
"""
import datetime
import json
import sys

from claude_client import resolve_claude_cmd, call_claude
from job_store import load_store, save_store, write_outputs

JOBS_PATH = "output/jobs_with_desc.json"
PROFILE = "resume/profile.json"

SCORE_SCHEMA = {
    "type": "object",
    "properties": {
        "fit_score": {"type": "integer"},
        "fit_reason": {"type": "string"},
        "tc_estimate_low": {"type": "integer"},
        "tc_estimate_high": {"type": "integer"},
        "tc_basis": {"type": "string"},
    },
    "required": ["fit_score", "fit_reason", "tc_estimate_low",
                 "tc_estimate_high", "tc_basis"],
}


def score_prompt(profile, job):
    have_desc = bool(job.get("description"))
    jobtext = job["description"] if have_desc else f"{job['title']} at {job['company']}"
    return f"""You are scoring how well a job posting fits a specific new-grad candidate. Return ONLY a fit score and a new-grad TC estimate — do not write a resume.

CANDIDATE PROFILE:
{json.dumps(profile, indent=2)}

JOB:
Company: {job['company']}
Title: {job['title']}
Locations: {job.get('locations')}
Start-date signal scraped from posting (may be empty): {job.get('start_signal', '')}
Description:
{jobtext}

TASKS:
1. SCORE fit 0-100 for THIS candidate (backend/distributed-systems/AI-infra new grad, grad May 2027). Be honest and lenient — only truly off-profile roles score low. One-sentence reason.
2. ESTIMATE new-grad total comp (USD) for this company/role as a low-high band. Use general knowledge of the company's tier; if unknown, give a wide band and say so. This is an estimate, not fact.
"""


def main():
    cmd_prefix = resolve_claude_cmd()
    if not cmd_prefix:
        print("claude CLI not found on PATH. Install with:\n"
              "  npm install -g @anthropic-ai/claude-code\n"
              "then run `claude setup-token` once.", file=sys.stderr)
        sys.exit(1)

    with open(JOBS_PATH, encoding="utf-8") as f:
        jobs = json.load(f)
    if not jobs:
        print("No new jobs.", file=sys.stderr)
        return

    with open(PROFILE, encoding="utf-8") as f:
        profile = json.load(f)
    store = load_store()
    today = datetime.date.today().isoformat()

    for i, job in enumerate(jobs, 1):
        print(f"[{i}/{len(jobs)}] {job['company']} — {job['title']}", file=sys.stderr)
        entry = dict(job)
        try:
            res = call_claude(cmd_prefix, score_prompt(profile, job), SCORE_SCHEMA)
            entry["fit_score"] = int(res.get("fit_score", 0))
            entry["fit_reason"] = res.get("fit_reason", "")
            entry["tc_estimate_low"] = int(res.get("tc_estimate_low", 0))
            entry["tc_estimate_high"] = int(res.get("tc_estimate_high", 0))
            entry["tc_basis"] = res.get("tc_basis", "")
            entry["status"] = "scored"
        except Exception as e:
            entry["fit_score"] = 0
            entry["status"] = f"score-error:{str(e)[:60]}"
        entry.setdefault("pdf_rel", "")
        entry["date_scored"] = today
        store[job["id"]] = entry

    save_store(store)
    n = write_outputs(store)
    print(f"Scored {len(jobs)} job(s). Sheet now has {n} total rows.", file=sys.stderr)


if __name__ == "__main__":
    main()
