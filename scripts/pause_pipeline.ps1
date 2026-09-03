# Pauses (or resumes) a running rank_jobs.py / rank_intern_jobs.py between
# jobs -- it's checked once per job, so it takes effect within a job or two,
# not instantly. Safe by design: each job is merge_and_save'd individually
# and the remaining to-do list lives in a backlog file, so pausing (or the
# process dying outright) never loses more than the one in-flight job.
#
# Usage:
#   scripts\pause_pipeline.ps1                 # pause the new-grad pipeline
#   scripts\pause_pipeline.ps1 -Intern         # pause the internship pipeline
#   scripts\pause_pipeline.ps1 -Resume         # clear the new-grad pause flag
#   scripts\pause_pipeline.ps1 -Intern -Resume # clear the internship pause flag
#
# Resuming doesn't restart anything by itself -- just re-run
# run_pipeline.ps1 / run_pipeline_intern.ps1 and it'll pick up the backlog.

param(
    [switch]$Intern,
    [switch]$Resume
)

$RepoRoot = Split-Path -Parent $PSScriptRoot
$FlagName = if ($Intern) { ".pause_intern" } else { ".pause_newgrad" }
$FlagPath = Join-Path $RepoRoot $FlagName
$Label = if ($Intern) { "internship" } else { "new-grad" }

if ($Resume) {
    if (Test-Path $FlagPath) {
        Remove-Item $FlagPath -Force
        Write-Output "Cleared pause flag for the $Label pipeline. Re-run its pipeline script to continue."
    } else {
        Write-Output "$Label pipeline wasn't paused."
    }
} else {
    New-Item -ItemType File -Path $FlagPath -Force | Out-Null
    Write-Output "Pause flag set for the $Label pipeline -- it'll stop after finishing the job currently in progress."
}
