# Daily Resume Tailor — Nate Cirino

Runs locally on your PC, triggered by Windows Task Scheduler when you log in
(no GitHub Actions, no API key — tailoring runs through the Claude Code CLI
against your subscription seat, not metered per-token billing):

1. Reads `listings.json` from SimplifyJobs/New-Grad-Positions (the machine-
   readable source the README rows are generated from). Falls back to the
   GitHub tree API if the path moves.
2. Keeps only **new**, active, **full-time** SWE / systems / backend / infra /
   AI-infra roles. Internships, co-ops, quant, PM, hardware, and clearance
   roles are skipped.
3. Best-effort fetches each posting's description (Greenhouse/Lever/Ashby APIs;
   Playwright for Workday/iCIMS/careers; graceful fallback) and scrapes any
   **start-date / 2027 signal**.
4. **Scores every new job** (fast, one small Claude call each — fit score
   0-100, fit reason, new-grad TC estimate). No resume rewriting, no LaTeX,
   no PDF compile at this stage, so this step stays quick even for a big
   batch of new postings.
5. Merges into `output/job_store.json` (persistent — never overwritten) and
   rewrites `output/tailored_resumes.xlsx`: one row per job ever scored,
   **ranked** by a blend of fit (70%) and estimated TC (30%). Fit cell is
   green/yellow/red. Columns link to the posting and a levels.fyi comp lookup.
6. Commits everything back to the repo.

**Full tailoring (rewriting the resume + compiling a PDF) is on demand, not
automatic** — that's the slow part, so it only runs for jobs you actually
want to apply to. Once you've skimmed the ranked sheet:

```
python scripts\tailor_selected.py --match "Fortinet" "Cohesity"
```

Matches are case-insensitive substrings against company or title; run with
several `--match` terms to do a batch at once. It writes the tailored
`.tex`, compiles the PDF into `output/pdfs/`, and updates that job's row in
place (Status → `tailored`, PDF link filled in). Already-tailored jobs are
skipped unless you pass `--force`.

## Important honesty notes

- **No job in the repo has salary data.** The "Est. TC" column is Claude's
  *estimate* from general company knowledge — not live comp data. Treat it as a
  ranking hint and verify real numbers via the levels.fyi link. That's why TC is
  never a hard filter: nothing is dropped for comp, it just sorts lower.
- **Grad-year (2027) usually isn't in the data either.** The pipeline scrapes
  the posting for start-date hints and flags likely-2027 roles in the "Likely
  2027?" column, but a blank there means "unknown," not "no." Full-time new-grad
  roles for a May 2027 grad mostly post during fall-2026 recruiting.
- **Nothing is fabricated.** The tailoring prompt forbids inventing employers,
  dates, or metrics — Claude only reorders and rephrases what's in your resume.
- **Fetch is best-effort.** Some ATS/career pages block datacenter IPs; those
  fall back to metadata-only scoring/tailoring (flagged in the "Tailor Depth"
  column).

## One-time setup

1. `resume/resume.tex` is already your real resume. Tweak `resume/profile.json`
   if you want to change how fit/TC are judged.
2. Install the Claude Code CLI: `npm install -g @anthropic-ai/claude-code`
   (requires Node.js).
3. Generate a long-lived token tied to your Claude Code subscription seat:
   `claude setup-token` (one-time, opens a browser to authorize).
4. Store it so unattended runs can read it — **never commit it to the repo**:
   `setx CLAUDE_CODE_OAUTH_TOKEN "<token from step 3>"`, then log out/in once
   so the new environment variable takes effect.
5. LaTeX (MiKTeX) must be installed for PDF compilation — installs missing
   packages automatically on first compile.
6. Python deps: `pip install openpyxl playwright` (Playwright is used for
   career pages that need a real browser to render).
7. First run only **seeds** the seen-list (no backlog tailoring). After that
   you get newly-added roles each time the pipeline runs.
8. Set up the trigger — Windows Task Scheduler:
   - Trigger: **At log on** (your user account).
   - Action: **Start a program** →
     `powershell.exe -ExecutionPolicy Bypass -File "<repo path>\scripts\run_pipeline.ps1"`
   - The script skips itself if it already completed successfully today, so
     logging in more than once a day is harmless.
   - Leave "Run whether user is logged on or not" **unchecked** — the task
     needs your logged-in session so it can read `CLAUDE_CODE_OAUTH_TOKEN`
     and push to GitHub with your credentials.
   - Add a second trigger, **On workstation unlock**, alongside "At log on"
     — locking/unlocking the screen does *not* fire a logon event on
     Windows, so without this the task only re-fires after a full sign-out
     or reboot. This trigger type isn't exposed by the simple Task Scheduler
     dropdown flow for a *new* task, but is available under the task's
     Properties → Triggers → New → "Begin the task" once the task exists.

## Get results

**Leave `output/dashboard.html` open in a browser tab** — it auto-refreshes
every 20 seconds (plain `<meta refresh>`, no server needed) and always
shows the current ranked list. Unlike the xlsx, a browser doesn't lock the
file it's displaying, so this is safe to leave open indefinitely; it'll
just repaint with fresh data after every run, no re-opening required.

`output/tailored_resumes.xlsx` is the portable snapshot (same data) for
when you actually want to open it in Excel — **don't leave it open**
while the pipeline might run, since Excel takes an exclusive lock on the
file and the next write will fail until you close it.

Rows are pre-ranked; skim top-down. Rows with Status `tailored` link to a
resume PDF in `output/pdfs/`; everything else has been scored but not yet
tailored — see below for pulling the trigger on specific ones.

## Manual run

Scan/rank only (what the scheduled task does):
`powershell -ExecutionPolicy Bypass -File scripts\run_pipeline.ps1` from the
repo root. Check `output\pipeline.log` for a run history / errors.

Tailor specific jobs on demand:
`python scripts\tailor_selected.py --match "<company or title>" [...]`

Add a single job by link (found outside the SimplifyJobs feed):
`python scripts\add_job.py --url "<job url>" [--title "..." --company "..."]`
— title/company auto-detect from the ATS API for Ashby/Greenhouse/Lever
links; pass them explicitly if auto-detection doesn't apply.

## Concurrency note

`run_pipeline.ps1` takes a lock file (`.pipeline.lock`) so two triggers
firing close together (e.g. a logon and an unlock within the same minute)
don't both start a run — this actually happened once during development and
produced duplicated, colliding output. If a run is ever killed uncleanly
(process killed rather than left to exit), delete `.pipeline.lock` by hand
before the next trigger, otherwise it'll wait for the stale PID check to
clear it on the next attempt.
