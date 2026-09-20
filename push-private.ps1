# Create/push the private GitHub repo pdac-v17 using the bundled gh.exe.
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
$gh = Join-Path $PSScriptRoot ".tools\gh\Program Files\GitHub CLI\gh.exe"
if (-not (Test-Path -LiteralPath $gh)) {
    throw "Missing bundled gh.exe under .tools\gh\Program Files\GitHub CLI\gh.exe"
}
Write-Host "Checking GitHub login..."
& $gh auth status
if ($LASTEXITCODE -ne 0) {
    Write-Host "Not logged in. Starting gh auth login (complete it in this window)..."
    & $gh auth login --hostname github.com --git-protocol https --web
    if ($LASTEXITCODE -ne 0) { throw "GitHub login failed" }
}
git remote get-url origin 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) {
    $existing = git remote get-url origin
    Write-Host "origin already set: $existing"
    git push -u origin main
    exit $LASTEXITCODE
}
Write-Host "Creating private repo pdac-v17 and pushing main..."
& $gh repo create pdac-v17 --private --source=. --remote=origin --push
if ($LASTEXITCODE -ne 0) { throw "gh repo create failed" }
