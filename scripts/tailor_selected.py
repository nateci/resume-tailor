#!/usr/bin/env python3
"""
On-demand full tailoring for specific jobs picked off the ranked sheet:
generates a tailored resume.tex and compiles a PDF, only for jobs matching
--match substrings (case-insensitive, checked against company or title).

Usage:
  python scripts/tailor_selected.py --match "Fortinet" "Cohesity"
  python scripts/tailor_selected.py --match "TikTok" --force   # re-tailor even if already done
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import subprocess
import sys

from claude_client import resolve_claude_cmd, call_claude
from job_store import load_store, save_store, write_sheet

RESUME_TEX = "resume/resume.tex"
PDF_DIR = "output/pdfs"

TAILOR_SCHEMA = {
    "type": "object",
    "properties": {"tailored_tex": {"type": "string"}},
    "required": ["tailored_tex"],
}


def slugify(s):
    return re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_")[:50]


def tag_for(job):
    # Suffix with a short hash of the job id so duplicate postings (same
    # company+title, different location/id — common in this feed) never
    # collide on the same .tex/.pdf filename.
    base = slugify(f"{job['company']}_{job['title']}")
    suffix = hashlib.sha1(job["id"].encode("utf-8")).hexdigest()[:6]
    return f"{base}_{suffix}"


def tailor_prompt(base_tex, job):
    have_desc = bool(job.get("description"))
    jobtext = job["description"] if have_desc else f"{job['title']} at {job['company']}"
    depth = ("Full description available — tailor substantively: reorder and "
             "rephrase bullets to mirror the posting, adjust the skills ordering."
             if have_desc else
             "Only title + company available — make light adjustments only; do "
             "not invent anything.")
    return f"""You are tailoring a resume for a specific new-grad candidate applying to a job.

JOB:
Company: {job['company']}
Title: {job['title']}
Description:
{jobtext}

TASK: TAILOR the resume below. {depth}

HARD RULES:
- Never fabricate employers, degrees, dates, or metrics. Only reorder/rephrase existing content.
- Keep it compilable with pdflatex and preserve the preamble and all custom macros.
- Keep it one page.

RESUME (LaTeX):
{base_tex}
"""


def compile_pdf(tex_path, out_dir):
    subprocess.run(
        ["latexmk", "-pdf", "-interaction=nonstopmode", "-halt-on-error",
         f"-output-directory={out_dir}", tex_path],
        check=True, capture_output=True, timeout=120,
    )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--match", nargs="+", required=True,
                     help="Case-insensitive substrings matched against company or title")
    ap.add_argument("--force", action="store_true",
                     help="Re-tailor even if already tailored")
    args = ap.parse_args()

    cmd_prefix = resolve_claude_cmd()
    if not cmd_prefix:
        print("claude CLI not found on PATH.", file=sys.stderr)
        sys.exit(1)

    store = load_store()
    with open(RESUME_TEX, encoding="utf-8") as f:
        base_tex = f.read()
    os.makedirs(PDF_DIR, exist_ok=True)

    needles = [m.lower() for m in args.match]
    targets = [
        j for j in store.values()
        if any(n in j.get("company", "").lower() or n in j.get("title", "").lower()
               for n in needles)
        and (args.force or j.get("status") != "tailored")
    ]

    if not targets:
        print("No matching un-tailored jobs found (use --force to re-tailor "
              "already-tailored ones).", file=sys.stderr)
        return

    for i, job in enumerate(targets, 1):
        tag = tag_for(job)
        print(f"[{i}/{len(targets)}] {job['company']} — {job['title']}", file=sys.stderr)
        try:
            res = call_claude(cmd_prefix, tailor_prompt(base_tex, job), TAILOR_SCHEMA)
            tailored = res.get("tailored_tex", "")
            if not tailored:
                job["status"] = "no-tex"
            else:
                tex_out = os.path.join(PDF_DIR, f"{tag}.tex")
                with open(tex_out, "w", encoding="utf-8") as f:
                    f.write(tailored)
                try:
                    compile_pdf(tex_out, PDF_DIR)
                    pdf_path = os.path.join(PDF_DIR, f"{tag}.pdf")
                    if os.path.exists(pdf_path):
                        job["pdf_rel"] = f"pdfs/{tag}.pdf"
                        job["status"] = "tailored"
                        job["date_tailored"] = datetime.date.today().isoformat()
                    else:
                        job["status"] = "compile-failed"
                except subprocess.CalledProcessError:
                    job["status"] = "compile-failed"
        except Exception as e:
            job["status"] = f"tailor-error:{str(e)[:60]}"
        store[job["id"]] = job

    save_store(store)
    n = write_sheet(store)
    print(f"Tailored {len(targets)} job(s). Sheet now has {n} total rows.", file=sys.stderr)


if __name__ == "__main__":
    main()
