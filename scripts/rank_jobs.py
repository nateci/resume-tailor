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
import os
import sys

from claude_client import resolve_claude_cmd, call_claude
from job_store import load_store, merge_and_save, write_outputs, commit_and_push

JOBS_PATH = "output/jobs_with_desc.json"
PROFILE = "resume/profile.json"

# Survives an interrupted run: filter_jobs.py marks a job "seen" the moment
# it's queued here, regardless of whether scoring ever finishes, so a
# killed run's un-scored remainder would otherwise never be re-queued as
# "new" again. This file is the actual to-do list, independent of that.
BACKLOG_PATH = "output/.newgrad_scoring_backlog.json"
PAUSE_PATH = "output/.pause_newgrad"


def _write_backlog(jobs):
    if jobs:
        with open(BACKLOG_PATH, "w", encoding="utf-8") as f:
            json.dump(jobs, f, indent=2)
    elif os.path.exists(BACKLOG_PATH):
        os.remove(BACKLOG_PATH)

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
        fresh_jobs = json.load(f)

    # Merge in any leftover backlog from an interrupted prior run -- those
    # jobs are already marked "seen" upstream so fresh_jobs alone won't
    # contain them anymore.
    backlog = []
    if os.path.exists(BACKLOG_PATH):
        with open(BACKLOG_PATH, encoding="utf-8") as f:
            backlog = json.load(f)
    by_id = {j["id"]: j for j in backlog}
    for j in fresh_jobs:
        by_id.setdefault(j["id"], j)

    # Drop anything already scored -- a prior run may have gotten to it
    # before being interrupted before it could prune the backlog file.
    existing = load_store()
    jobs = [j for j in by_id.values() if j["id"] not in existing]

    if not jobs:
        print("No new jobs.", file=sys.stderr)
        _write_backlog([])
        return

    with open(PROFILE, encoding="utf-8") as f:
        profile = json.load(f)
    today = datetime.date.today().isoformat()

    remaining = list(jobs)
    _write_backlog(remaining)

    for i, job in enumerate(jobs, 1):
        if os.path.exists(PAUSE_PATH):
            print(f"Paused ({len(remaining)} job(s) left in {BACKLOG_PATH}) -- "
                  f"run scripts/pause_pipeline.ps1 -Resume, then re-run this "
                  f"pipeline to continue where it left off.", file=sys.stderr)
            return

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

        # Saved right after each job, not batched to the end of the loop --
        # a kill/crash mid-run now loses at most the one in-flight job
        # instead of the entire batch.
        merge_and_save({job["id"]: entry})
        remaining = [j for j in remaining if j["id"] != job["id"]]
        _write_backlog(remaining)

    merged = load_store()
    n = write_outputs(merged)
    print(f"Scored {len(jobs)} job(s). Sheet now has {n} total rows.", file=sys.stderr)
    if commit_and_push(f"Scored {len(jobs)} job(s): {today}"):
        print("Committed and pushed.", file=sys.stderr)


if __name__ == "__main__":
    main()
