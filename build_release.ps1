# build_release.ps1
# ------------------------------------------------------------------------------
# DOCREFINE PRO BUILDER (v133 Spec-First Edition)
# ------------------------------------------------------------------------------
# VERSION SYNC: Updated from v132 to v133
$Version = "v133"
$ErrorActionPreference = "Stop"

Write-Host "🚀 STARTING PRODUCTION BUILD [$Version]..." -ForegroundColor Cyan

# 1. CLEANUP
Write-Host "🧹 Cleaning workspace..." -ForegroundColor Yellow
if (Test-Path "dist") { Remove-Item -Path "dist" -Recurse -Force }
if (Test-Path "build") { Remove-Item -Path "build" -Recurse -Force }

# 2. BUILD COMMAND
Write-Host "🔨 Compiling binary using Spec file..." -ForegroundColor Yellow

# v132 Update: Removed CLI overrides (--collect-all). 
# We now rely 100% on DocRefinePro.spec for inclusion/exclusion logic.
python -m PyInstaller DocRefinePro.spec --noconfirm --clean

# 3. VERIFICATION
$BinPath = "dist/DocRefinePro/DocRefinePro.exe"

if (Test-Path $BinPath) {
    Write-Host "✅ BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "📦 Output: $BinPath" -ForegroundColor White
}

if (-not (Test-Path $BinPath)) {
    Write-Host "❌ BUILD FAILED." -ForegroundColor Red
}
