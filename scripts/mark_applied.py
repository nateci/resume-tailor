#!/usr/bin/env python3
"""
Flags a job_store entry as applied=True without requiring it to have gone
through tailor_selected.py -- scan_email_status.py only checks companies
from jobs marked status=="tailored" or applied=True, so a posting applied
to some other way (quick-apply, no custom resume, applied before this
pipeline existed) would otherwise never enter its matching pool even
though a real confirmation email is sitting right there in the inbox.

Usage:
  python scripts/mark_applied.py truveta
  python scripts/mark_applied.py truveta --intern
  python scripts/mark_applied.py truveta --title "live link"   (disambiguate multiple postings)
"""
import argparse
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0].rsplit("\\", 1)[0])
from job_store import load_store, save_store


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("company", help="Case-insensitive substring match against the Company field")
    ap.add_argument("--title", default=None, help="Case-insensitive substring to disambiguate multiple postings at the same company")
    ap.add_argument("--intern", action="store_true", help="Use output/intern_job_store.json instead of output/job_store.json")
    args = ap.parse_args()

    store_path = "output/intern_job_store.json" if args.intern else "output/job_store.json"
    store = load_store(store_path)

    company_q = args.company.lower()
    title_q = (args.title or "").lower()
    matches = [
        (jid, j) for jid, j in store.items()
        if company_q in j.get("company", "").lower()
        and (not title_q or title_q in j.get("title", "").lower())
    ]

    if not matches:
        sys.exit(f"No job found matching company={args.company!r} title={args.title!r} in {store_path}")
    if len(matches) > 1:
        print(f"{len(matches)} matches -- re-run with --title to narrow down:", file=sys.stderr)
        for jid, j in matches:
            print(f"  - {j['company']} | {j['title']}", file=sys.stderr)
        sys.exit(1)

    jid, j = matches[0]
    j["applied"] = True
    save_store(store, store_path)
    print(f"Marked applied: {j['company']} | {j['title']}", file=sys.stderr)
    print("Run scripts/scan_email_status.py (or click Update) to pick up its confirmation email.",
          file=sys.stderr)


if __name__ == "__main__":
    main()
