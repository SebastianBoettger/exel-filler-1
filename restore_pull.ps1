# Git Restore Script (Select Commit) - ASCII safe

Set-Location (Split-Path $MyInvocation.MyCommand.Path)

if (-not (Test-Path ".git")) {
    Write-Host "ERROR: No Git repository found here." -ForegroundColor Red
    exit 1
}

Write-Host "Recent commits:" -ForegroundColor Cyan
git log --oneline --decorate -n 30

$commit = Read-Host "Enter commit hash to restore (e.g. a1b2c3d)"
if ([string]::IsNullOrWhiteSpace($commit)) {
    Write-Host "ERROR: No commit given." -ForegroundColor Red
    exit 1
}

$confirm = Read-Host "WARNING: Local changes will be lost. Continue? (yes/no)"
if ($confirm -ne "yes") {
    Write-Host "Canceled." -ForegroundColor Yellow
    exit 0
}

git fetch origin

git reset --hard $commit
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Reset failed." -ForegroundColor Red
    exit 1
}

git clean -fd

Write-Host "OK: Restored to:" -ForegroundColor Green
git log -1 --oneline
