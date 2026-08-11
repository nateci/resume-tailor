#!/usr/bin/env python3
"""
Per new job, ONE Claude call returns:
  - fit_score (0-100) vs Nate's profile
  - fit_reason (one line)
  - tc_estimate_low / tc_estimate_high (USD new-grad TC, best-effort)
  - tc_basis (why)
  - tailored_tex (the full tailored resume)

Then compile -> PDF, and append to a ranked xlsx sorted by a blend of fit and TC.
Nothing is dropped; low matches just sort to the bottom.

Requires ANTHROPIC_API_KEY. LaTeX via latexmk/pdflatex.
"""
import json
import os
import re
import subprocess
import sys
import datetime

import anthropic
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

JOBS_PATH = "output/jobs_with_desc.json"
RESUME_TEX = "resume/resume.tex"
PROFILE = "resume/profile.json"
PDF_DIR = "output/pdfs"
SHEET_PATH = "output/tailored_resumes.xlsx"
MODEL = "claude-opus-4-8"

client = anthropic.Anthropic()


def slugify(s):
    return re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_")[:60]


def ask_claude(base_tex, profile, job):
    have_desc = bool(job.get("description"))
    jobtext = job["description"] if have_desc else f"{job['title']} at {job['company']}"
    depth = ("Full description available — tailor substantively: reorder and "
             "rephrase bullets to mirror the posting, adjust the skills ordering."
             if have_desc else
             "Only title + company available — make light adjustments only; do "
             "not invent anything.")

    prompt = f"""You are helping a specific new-grad candidate apply to a job. Do two things and return ONE JSON object, nothing else.

CANDIDATE PROFILE:
{json.dumps(profile, indent=2)}

JOB:
Company: {job['company']}
Title: {job['title']}
Locations: {job.get('locations')}
Start-date signal scraped from posting (may be empty): {job.get('start_signal','')}
Description:
{jobtext}

TASKS:
1. SCORE fit 0-100 for THIS candidate (backend/distributed-systems/AI-infra new grad, grad May 2027). Be honest and lenient — only truly off-profile roles score low. One-sentence reason.
2. ESTIMATE new-grad total comp (USD) for this company/role as a low-high band. Use general knowledge of the company's tier; if unknown, give a wide band and say so. This is an estimate, not fact.
3. TAILOR the resume below. {depth}

HARD RULES for the resume:
- Never fabricate employers, degrees, dates, or metrics. Only reorder/rephrase existing content.
- Keep it compilable with pdflatex and preserve the preamble and all custom macros.
- Keep it one page.

Return EXACTLY this JSON (no markdown fences):
{{
  "fit_score": <int 0-100>,
  "fit_reason": "<one sentence>",
  "tc_estimate_low": <int USD>,
  "tc_estimate_high": <int USD>,
  "tc_basis": "<short reason>",
  "tailored_tex": "<the full LaTeX document as a JSON string>"
}}

RESUME (LaTeX):
{base_tex}
"""
    msg = client.messages.create(
        model=MODEL, max_tokens=6000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = "".join(b.text for b in msg.content if b.type == "text").strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw)
    return json.loads(raw)


def compile_pdf(tex_path, out_dir):
    subprocess.run(
        ["latexmk", "-pdf", "-interaction=nonstopmode", "-halt-on-error",
         f"-output-directory={out_dir}", tex_path],
        check=True, capture_output=True, timeout=120,
    )


HEADERS = ["Rank", "Fit", "Est. TC (USD)", "Company", "Role", "Location",
           "Likely 2027?", "Start Signal", "Fit Reason", "TC Basis",
           "Tailor Depth", "Status", "Job Link", "Resume PDF", "levels.fyi",
           "Date Added"]


def open_or_create_sheet():
    if os.path.exists(SHEET_PATH):
        wb = load_workbook(SHEET_PATH)
        return wb, wb.active
    wb = Workbook()
    ws = wb.active
    ws.title = "Ranked Jobs"
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(1, i, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5496")
    widths = [5, 5, 16, 20, 34, 22, 12, 26, 40, 28, 13, 14, 10, 30, 12, 11]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    return wb, ws


def collect_existing(ws):
    """Read existing data rows (below header) so we can re-rank after adding."""
    rows = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 4).value is None:
            continue
        rows.append({
            "fit": ws.cell(r, 2).value or 0,
            "tc_hi": _num(ws.cell(r, 3).value),
            "cells": [ws.cell(r, c).value for c in range(1, len(HEADERS) + 1)],
            "links": {13: ws.cell(r, 13).hyperlink, 14: ws.cell(r, 14).hyperlink,
                      15: ws.cell(r, 15).hyperlink},
        })
    return rows


def _num(v):
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        m = re.findall(r"\d+", v.replace(",", ""))
        if m:
            return int(m[-1])
    return 0


def main():
    jobs = json.load(open(JOBS_PATH))
    if not jobs:
        print("No new jobs.", file=sys.stderr)
        return
    base_tex = open(RESUME_TEX).read()
    profile = json.load(open(PROFILE))
    os.makedirs(PDF_DIR, exist_ok=True)
    wb, ws = open_or_create_sheet()

    existing = collect_existing(ws)
    new_rows = []

    for i, job in enumerate(jobs, 1):
        tag = slugify(f"{job['company']}_{job['title']}")
        print(f"[{i}/{len(jobs)}] {tag}", file=sys.stderr)
        status, pdf_rel = "ok", ""
        fit, reason = 0, ""
        tc_lo, tc_hi, tc_basis = 0, 0, ""
        try:
            res = ask_claude(base_tex, profile, job)
            fit = int(res.get("fit_score", 0))
            reason = res.get("fit_reason", "")
            tc_lo = int(res.get("tc_estimate_low", 0))
            tc_hi = int(res.get("tc_estimate_high", 0))
            tc_basis = res.get("tc_basis", "")
            tailored = res.get("tailored_tex", "")
            if tailored:
                tex_out = os.path.join(PDF_DIR, f"{tag}.tex")
                open(tex_out, "w").write(tailored)
                try:
                    compile_pdf(tex_out, PDF_DIR)
                    if os.path.exists(os.path.join(PDF_DIR, f"{tag}.pdf")):
                        pdf_rel = f"pdfs/{tag}.pdf"
                    else:
                        status = "compile-failed"
                except subprocess.CalledProcessError:
                    status = "compile-failed"
            else:
                status = "no-tex"
        except Exception as e:
            status = f"error:{str(e)[:30]}"

        loc = job.get("locations")
        loc = ", ".join(loc) if isinstance(loc, list) else str(loc)
        tc_disp = f"${tc_lo//1000}k–${tc_hi//1000}k" if tc_hi else "unknown"
        lv = ("https://www.levels.fyi/?compare=" +
              job["company"].replace(" ", "%20") + "&track=Software%20Engineer")
        new_rows.append({
            "fit": fit, "tc_hi": tc_hi,
            "data": {
                "Fit": fit, "Est. TC (USD)": tc_disp, "Company": job["company"],
                "Role": job["title"], "Location": loc,
                "Likely 2027?": "yes" if job.get("likely_2027") else "",
                "Start Signal": job.get("start_signal", ""),
                "Fit Reason": reason, "TC Basis": tc_basis,
                "Tailor Depth": "full" if job.get("description") else "metadata-only",
                "Status": status,
            },
            "job_url": job.get("url", ""), "pdf_rel": pdf_rel, "levels": lv,
            "date": datetime.date.today().isoformat(),
        })

    # Merge existing + new, re-rank by blended score (fit primary, TC secondary)
    def blend(fit, tc_hi):
        tc_norm = min(tc_hi, 300000) / 300000 * 100
        return 0.7 * fit + 0.3 * tc_norm

    merged = existing + [
        {"fit": r["fit"], "tc_hi": r["tc_hi"], "_new": r} for r in new_rows
    ]
    merged.sort(key=lambda x: blend(x["fit"], x["tc_hi"]), reverse=True)

    # Rewrite all data rows in ranked order
    for r in range(ws.max_row, 1, -1):
        ws.delete_rows(r)

    for rank, item in enumerate(merged, 1):
        row = ws.max_row + 1
        if "_new" in item:
            n = item["_new"]
            d = n["data"]
            vals = [rank, d["Fit"], d["Est. TC (USD)"], d["Company"], d["Role"],
                    d["Location"], d["Likely 2027?"], d["Start Signal"],
                    d["Fit Reason"], d["TC Basis"], d["Tailor Depth"], d["Status"],
                    "Apply ↗" if n["job_url"] else "",
                    os.path.basename(n["pdf_rel"]) if n["pdf_rel"] else "",
                    "levels.fyi ↗", n["date"]]
            for c, v in enumerate(vals, 1):
                ws.cell(row, c, v)
            if n["job_url"]:
                _link(ws.cell(row, 13), n["job_url"])
            if n["pdf_rel"]:
                _link(ws.cell(row, 14), n["pdf_rel"])
            _link(ws.cell(row, 15), n["levels"])
        else:
            cells = item["cells"]
            cells[0] = rank  # refresh rank
            for c, v in enumerate(cells, 1):
                ws.cell(row, c, v)
            for col, link in item["links"].items():
                if link:
                    _link(ws.cell(row, col), link.target if hasattr(link, "target") else link)
        _color_fit(ws.cell(row, 2))

    wb.save(SHEET_PATH)
    print(f"Wrote {SHEET_PATH} — {len(merged)} total rows", file=sys.stderr)


def _link(cell, target):
    cell.hyperlink = target
    cell.font = Font(color="0563C1", underline="single")


def _color_fit(cell):
    v = cell.value or 0
    try:
        v = int(v)
    except Exception:
        return
    color = "C6EFCE" if v >= 75 else "FFEB9C" if v >= 50 else "FFC7CE"
    cell.fill = PatternFill("solid", fgColor=color)


if __name__ == "__main__":
    main()
