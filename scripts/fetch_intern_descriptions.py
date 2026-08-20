#!/usr/bin/env python3
"""Internship counterpart to fetch_descriptions.py -- same fetch logic
(Greenhouse/Lever/Ashby APIs, Playwright, plain HTTP), pointed at the
internship pipeline's own new-jobs file."""
from fetch_descriptions import main

IN_PATH = "output/intern_new_jobs.json"
OUT_PATH = "output/intern_jobs_with_desc.json"

if __name__ == "__main__":
    main(in_path=IN_PATH, out_path=OUT_PATH)
