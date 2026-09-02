# Internship counterpart to run_pipeline.ps1 -- scans SimplifyJobs'
# Summer2027-Internships feed instead of New-Grad-Positions, targeting a
# Spring 2027 start (see resume/profile_intern.json). Fully separate state
# from the new-grad pipeline: its own lock file, state file, log, and
# output artifacts (intern_job_store.json / intern_tailored_resumes.xlsx /
# intern_dashboard.html), so the two scans never interact or race.
#
# Not on a Task Scheduler trigger yet -- run by hand with:
#   powershell -File scripts\run_pipeline_intern.ps1
# Add a scheduled trigger later the same way run_pipeline.ps1 is wired up,
# if you want this to run unattended too.

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

$StateFile = Join-Path $RepoRoot ".last_run_intern"
$LockFile = Join-Path $RepoRoot ".pipeline_intern.lock"
$Today = (Get-Date).ToString("yyyy-MM-dd")
$LogFile = Join-Path $RepoRoot "output\intern_pipeline.log"

function Log($msg) {
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    Write-Output $line
    Add-Content -Path $LogFile -Value $line
}

function Invoke-PyScript($scriptPath) {
    $stdout = New-Object System.Collections.Generic.List[string]
    & python $scriptPath 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            Add-Content -Path $LogFile -Value $_.Exception.Message
        } else {
            $stdout.Add($_)
            Add-Content -Path $LogFile -Value $_
        }
    }
    [PSCustomObject]@{ ExitCode = $LASTEXITCODE; StdOut = $stdout }
}

New-Item -ItemType Directory -Force -Path (Join-Path $RepoRoot "output") | Out-Null

if ((Test-Path $StateFile) -and ((Get-Content $StateFile -Raw).Trim() -eq $Today)) {
    Log "Already completed today ($Today). Skipping."
    exit 0
}

if (Test-Path $LockFile) {
    $lockPid = (Get-Content $LockFile -Raw).Trim()
    $existing = Get-Process -Id $lockPid -ErrorAction SilentlyContinue
    if ($existing) {
        Log "Another run is already in progress (PID $lockPid). Skipping."
        exit 0
    }
    Log "Found stale lock file (PID $lockPid no longer running). Continuing."
}
Set-Content -Path $LockFile -Value $PID

try {
    if (-not (Get-Command claude -ErrorAction SilentlyContinue)) {
        Log "ERROR: claude CLI not found on PATH."
        exit 1
    }
    if (-not $env:CLAUDE_CODE_OAUTH_TOKEN) {
        Log "ERROR: CLAUDE_CODE_OAUTH_TOKEN is not set for this session. Run 'claude setup-token' and 'setx CLAUDE_CODE_OAUTH_TOKEN <token>' once, then log out/in."
        exit 1
    }

    Log "=== Run started ==="

    try {
        # --autostash: an ad-hoc script (scan_email_status.py, manual testing,
        # etc.) leaving the tree dirty between scheduled runs has repeatedly
        # blocked this outright -- stash/pull/pop instead of just failing on
        # any uncommitted local change. (See run_pipeline.ps1's same fix.)
        git pull --rebase --autostash --quiet
        if ($LASTEXITCODE -ne 0) { throw "git pull failed" }

        $filterResult = Invoke-PyScript "scripts\filter_intern_jobs.py"
        if ($filterResult.ExitCode -ne 0) { throw "filter_intern_jobs.py failed" }
        $count = $filterResult.StdOut | Select-Object -Last 1
        Log "New internships: $count"

        if ([int]$count -gt 0) {
            $fetchResult = Invoke-PyScript "scripts\fetch_intern_descriptions.py"
            if ($fetchResult.ExitCode -ne 0) { throw "fetch_intern_descriptions.py failed" }

            $rankResult = Invoke-PyScript "scripts\rank_intern_jobs.py"
            if ($rankResult.ExitCode -ne 0) { throw "rank_intern_jobs.py failed" }
        }

        git add output/
        git diff --cached --quiet | Out-Null
        $hasChanges = ($LASTEXITCODE -ne 0)
        if ($hasChanges) {
            git commit -m "Ranked internships: $Today" --quiet
            git push --quiet
            Log "Committed and pushed results."
        } else {
            Log "No changes to commit."
        }

        Set-Content -Path $StateFile -Value $Today
        Log "=== Run completed successfully ==="
    } catch {
        Log "ERROR: $_"
        exit 1
    }
} finally {
    Remove-Item -Path $LockFile -Force -ErrorAction SilentlyContinue
}
