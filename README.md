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
4. One Claude call per job returns: a **fit score (0-100)** vs your profile, a
   **new-grad TC estimate**, and a **tailored `resume.tex`** — then compiles a PDF.
5. Writes `output/tailored_resumes.xlsx`: one row per job, **ranked** by a blend
   of fit (70%) and estimated TC (30%). Fit cell is green/yellow/red. Columns
   link to the posting, the tailored PDF, and a levels.fyi comp lookup.
6. Commits everything back to the repo.

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
  fall back to metadata-only tailoring (flagged in the "Tailor Depth" column).

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

## Get results

Pull the repo or open `output/tailored_resumes.xlsx` directly on your machine
— it's updated and pushed automatically after each run. Rows are pre-ranked;
skim top-down. Each row links to its tailored PDF in `output/pdfs/`.

## Manual run

`powershell -ExecutionPolicy Bypass -File scripts\run_pipeline.ps1` from the
repo root. Check `output\pipeline.log` for a run history / errors.
