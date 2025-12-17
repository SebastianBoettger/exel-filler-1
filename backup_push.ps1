# Git Backup Push Script (ASCII safe)

Set-Location (Split-Path $MyInvocation.MyCommand.Path)

if (-not (Test-Path ".git")) {
    Write-Host "ERROR: No Git repository found here." -ForegroundColor Red
    exit 1
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
$customMsg = Read-Host "Commit description"
if ([string]::IsNullOrWhiteSpace($customMsg)) {
    $customMsg = "Backup"
}

$commitMsg = "$timestamp | $customMsg"

git status
git add .

git commit -m "$commitMsg"
if ($LASTEXITCODE -ne 0) {
    Write-Host "INFO: Nothing to commit." -ForegroundColor Yellow
    exit 0
}

git push
if ($LASTEXITCODE -eq 0) {
    Write-Host "OK: Backup pushed." -ForegroundColor Green
    Write-Host "Commit: $commitMsg"
} else {
    Write-Host "ERROR: Push failed." -ForegroundColor Red
    exit 1
}
