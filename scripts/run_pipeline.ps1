# Runs the full resume-tailor pipeline locally, using the Claude Code CLI
# (subscription seat, not the metered API). Meant to be triggered by Windows
# Task Scheduler on an "At log on" trigger — safe to invoke more than once a
# day, since it skips itself if it already completed successfully today.

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

$StateFile = Join-Path $RepoRoot ".last_run"
$Today = (Get-Date).ToString("yyyy-MM-dd")
$LogFile = Join-Path $RepoRoot "output\pipeline.log"

function Log($msg) {
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    Write-Output $line
    Add-Content -Path $LogFile -Value $line
}

# Runs a python script, merging stdout+stderr into the log as plain text.
# Native-command stderr normally gets wrapped in PowerShell's verbose
# ErrorRecord formatting when redirected with 2>> — this flattens it to the
# original line so informational stderr prints don't look like failures.
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
    git pull --rebase --quiet
    if ($LASTEXITCODE -ne 0) { throw "git pull failed" }

    $filterResult = Invoke-PyScript "scripts\filter_jobs.py"
    if ($filterResult.ExitCode -ne 0) { throw "filter_jobs.py failed" }
    $count = $filterResult.StdOut | Select-Object -Last 1
    Log "New jobs: $count"

    if ([int]$count -gt 0) {
        $fetchResult = Invoke-PyScript "scripts\fetch_descriptions.py"
        if ($fetchResult.ExitCode -ne 0) { throw "fetch_descriptions.py failed" }

        $tailorResult = Invoke-PyScript "scripts\tailor_and_build.py"
        if ($tailorResult.ExitCode -ne 0) { throw "tailor_and_build.py failed" }
    }

    git add output/
    $staged = git diff --cached --quiet; $hasChanges = ($LASTEXITCODE -ne 0)
    if ($hasChanges) {
        git commit -m "Ranked tailored resumes: $Today" --quiet
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
