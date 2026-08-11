# Daily Resume Tailor — Nate Cirino

Every morning, GitHub Actions (your PC stays off):

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

1. New GitHub repo (private is fine); drop these files in.
2. `resume/resume.tex` is already your real resume. Tweak `resume/profile.json`
   if you want to change how fit/TC are judged.
3. Repo → Settings → Secrets and variables → Actions → new secret
   `ANTHROPIC_API_KEY`.
4. First run only **seeds** the seen-list (no backlog tailoring). After that you
   get newly-added roles each morning.
5. Edit the `cron:` time in `.github/workflows/tailor.yml` (UTC) for your morning.

## Get results

Pull the repo or download `output/tailored_resumes.xlsx`. Rows are pre-ranked;
skim top-down. Each row links to its tailored PDF in `output/pdfs/`.

## Manual run

Actions tab → "Daily Resume Tailor" → "Run workflow."
