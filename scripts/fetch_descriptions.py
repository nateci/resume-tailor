#!/usr/bin/env python3
"""
Best-effort per-job fetch:
  - description text (Greenhouse/Lever/Ashby JSON APIs; Playwright for
    Workday/iCIMS/careers; plain HTTP last resort)
  - start-date / grad-year signal extracted from that text

Never fatal on a single job. Failures leave description="" and the tailor
step falls back to metadata-only for that row.
"""
import json
import re
import sys
import time
import urllib.request

IN_PATH = "output/new_jobs.json"
OUT_PATH = "output/jobs_with_desc.json"
UA = {"User-Agent": "Mozilla/5.0 (compatible; resume-tailor/1.0)"}

# Signals that a role targets a 2027 start / 2027 grads.
START_PATTERNS = [
    r"class of 20\d{2}",
    r"graduat\w+ (?:in |by |between )?\w*\s*20\d{2}",
    r"start date[:\s]+[A-Za-z0-9 ,]{3,30}",
    r"starting (?:in )?(?:january|february|march|april|may|june|july|august|"
    r"september|october|november|december|spring|summer|fall|winter)?\s*20\d{2}",
    r"(?:spring|summer|fall|winter)\s+20\d{2}",
    r"expected graduation[:\s]+[A-Za-z0-9 ,]{3,30}",
]


def http_get(url, timeout=20):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", errors="replace")


def strip_html(html):
    html = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"&nbsp;|&amp;|&#39;|&quot;|&rsquo;", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text[:6000]


def greenhouse(url):
    m = re.search(r"greenhouse\.io/([\w-]+)/jobs/(\d+)", url) or \
        re.search(r"greenhouse\.io/embed/job_app\?for=([\w-]+).*?(\d{5,})", url)
    if not m:
        return None
    org, jid = m.group(1), m.group(2)
    data = json.loads(http_get(f"https://boards-api.greenhouse.io/v1/boards/{org}/jobs/{jid}"))
    return strip_html(data.get("content", ""))


def lever(url):
    m = re.search(r"lever\.co/([\w-]+)/([\w-]+)", url)
    if not m:
        return None
    org, jid = m.group(1), m.group(2)
    data = json.loads(http_get(f"https://api.lever.co/v0/postings/{org}/{jid}"))
    parts = [data.get("descriptionPlain", "")]
    for li in data.get("lists", []):
        parts.append(li.get("text", ""))
        parts.append(strip_html(li.get("content", "")))
    return strip_html(" ".join(parts))


def ashby(url):
    m = re.search(r"ashbyhq\.com/([\w-]+)/([\w-]+)", url)
    if not m:
        return None
    org, jid = m.group(1), m.group(2)
    data = json.loads(http_get("https://api.ashbyhq.com/posting-api/job-board/" + org))
    for job in data.get("jobs", []):
        if job.get("id") == jid or jid in json.dumps(job):
            return strip_html(job.get("descriptionHtml", job.get("descriptionPlain", "")))
    return None


API_HANDLERS = [("greenhouse", greenhouse), ("lever", lever), ("ashby", ashby)]


def try_playwright(url):
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        return None
    try:
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(user_agent=UA["User-Agent"])
            pg.goto(url, timeout=25000, wait_until="domcontentloaded")
            pg.wait_for_timeout(2500)
            html = pg.content()
            b.close()
            return strip_html(html)
    except Exception as e:
        print(f"  playwright failed: {e}", file=sys.stderr)
        return None


def fetch_one(url):
    if not url:
        return "", "no-url"
    for name, handler in API_HANDLERS:
        if name in url:
            try:
                d = handler(url)
                if d and len(d) > 120:
                    return d, f"api:{name}"
            except Exception as e:
                print(f"  {name} api failed: {e}", file=sys.stderr)
    d = try_playwright(url)
    if d and len(d) > 200:
        return d, "playwright"
    try:
        d = strip_html(http_get(url))
        if d and len(d) > 200:
            return d, "http"
    except Exception as e:
        print(f"  http failed: {e}", file=sys.stderr)
    return "", "failed"


def extract_start_signal(text):
    if not text:
        return "", False
    low = text.lower()
    hits = []
    for pat in START_PATTERNS:
        for m in re.finditer(pat, low):
            hits.append(m.group(0).strip())
    hits = list(dict.fromkeys(hits))[:3]  # dedupe, cap
    signal = "; ".join(hits)
    likely_2027 = "2027" in signal
    return signal, likely_2027


def main():
    jobs = json.load(open(IN_PATH))
    out, stats = [], {}
    for i, job in enumerate(jobs, 1):
        print(f"[{i}/{len(jobs)}] {job['company']} — {job['title']}", file=sys.stderr)
        desc, source = fetch_one(job.get("url", ""))
        job["description"] = desc
        job["desc_source"] = source
        sig, likely = extract_start_signal(desc)
        job["start_signal"] = sig
        job["likely_2027"] = likely
        stats[source] = stats.get(source, 0) + 1
        out.append(job)
        time.sleep(1.0)
    json.dump(out, open(OUT_PATH, "w"), indent=2)
    print(f"Fetch sources: {stats}", file=sys.stderr)


if __name__ == "__main__":
    main()
