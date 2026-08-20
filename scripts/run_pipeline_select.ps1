# Interactive launcher: pick which pipeline to run instead of remembering
# script names. Just runs the chosen pipeline script in this same window.

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

Write-Host ""
Write-Host "Resume Tailor -- which pipeline?" -ForegroundColor Cyan
Write-Host "  1) New Grad   (scans SimplifyJobs New-Grad-Positions)"
Write-Host "  2) Internship (scans SimplifyJobs Summer2027-Internships, targets Spring 2027)"
Write-Host ""

$choice = Read-Host "Enter 1 or 2"

switch ($choice.Trim()) {
    "1" {
        Write-Host "Running New Grad pipeline..." -ForegroundColor Green
        & (Join-Path $PSScriptRoot "run_pipeline.ps1")
    }
    "2" {
        Write-Host "Running Internship pipeline..." -ForegroundColor Green
        & (Join-Path $PSScriptRoot "run_pipeline_intern.ps1")
    }
    default {
        Write-Host "Not 1 or 2 -- nothing run." -ForegroundColor Yellow
    }
}
