# build_release.ps1
# ------------------------------------------------------------------------------
# DOCREFINE PRO BUILDER (v119+ PySide6 Edition)
# ------------------------------------------------------------------------------
$Version = "v119"
$ErrorActionPreference = "Stop"

Write-Host "🚀 STARTING PRODUCTION BUILD [$Version]..." -ForegroundColor Cyan

# 1. CLEANUP
Write-Host "🧹 Cleaning workspace..." -ForegroundColor Yellow
if (Test-Path "dist") { Remove-Item -Path "dist" -Recurse -Force }
if (Test-Path "build") { Remove-Item -Path "build" -Recurse -Force }

# 2. BUILD COMMAND
Write-Host "🔨 Compiling binary..." -ForegroundColor Yellow

# Note: We use 'python -m PyInstaller' to ensure we use the active environment
python -m PyInstaller --noconfirm --onedir --noconsole --clean `
    --name "DocRefinePro" `
    --collect-all "PySide6" `
    --hidden-import "PySide6.QtXml" `
    --hidden-import "PySide6.QtNetwork" `
    --paths "." `
    main.py

# 3. VERIFICATION
$BinPath = "dist/DocRefinePro/DocRefinePro.exe"

if (Test-Path $BinPath) {
    Write-Host "✅ BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "📦 Output: $BinPath" -ForegroundColor White
}

if (-not (Test-Path $BinPath)) {
    Write-Host "❌ BUILD FAILED." -ForegroundColor Red
}