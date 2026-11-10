# Build diagnostics script for SMART-SIGN-PRO
# Usage: Open PowerShell in the project root and run: .\scripts\collect_build_info.ps1

$diagDir = Join-Path -Path (Get-Location) -ChildPath 'build-diagnostics'
if (-Not (Test-Path $diagDir)) { New-Item -ItemType Directory -Path $diagDir | Out-Null }

"Date: $(Get-Date -Format o)" | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Encoding utf8
"PWD: $(Get-Location)" | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Append -Encoding utf8

# Node / npm / npx versions
node -v 2>&1 | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Append -Encoding utf8
npm -v 2>&1 | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Append -Encoding utf8
npx -v 2>&1 | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Append -Encoding utf8

# PATH helpers
"where node:" | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Append -Encoding utf8
where node 2>&1 | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Append -Encoding utf8
where npm 2>&1 | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Append -Encoding utf8
where npx 2>&1 | Out-File -FilePath (Join-Path $diagDir 'diag.txt') -Append -Encoding utf8

# Git status and remote
git status --porcelain 2>&1 | Out-File -FilePath (Join-Path $diagDir 'git-status.txt') -Encoding utf8
git remote -v 2>&1 | Out-File -FilePath (Join-Path $diagDir 'git-remote.txt') -Encoding utf8

# List tree
Get-ChildItem -Force -Recurse -ErrorAction SilentlyContinue | Select-Object FullName -First 2000 | Out-File -FilePath (Join-Path $diagDir 'file-list.txt') -Encoding utf8

# Install and build (capture full output)
Write-Host "Running npm ci... (this may take a while)"
npm ci 2>&1 | Tee-Object -FilePath (Join-Path $diagDir 'npm-install.txt')

Write-Host "Running npm run build..."
npm run build 2>&1 | Tee-Object -FilePath (Join-Path $diagDir 'build-output.txt')

Write-Host "Diagnostics saved to: $diagDir"
Write-Host "Please attach the files build-output.txt and npm-install.txt (and diag.txt) when reporting results."
