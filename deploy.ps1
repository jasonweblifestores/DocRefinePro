# SAVE AS: deploy.ps1 in your Project Root
param (
    [Parameter(Mandatory=$true)]
    [string]$Version
)

# --- CONFIGURATION ---
$DriveLink = "https://drive.google.com/drive/folders/1_9IoOZK5dW6rsjUp5eq-3vT8LIOlKYkU?usp=sharing"
$GistFolder = ".\53752cda3c39550673fc5dafb96c4bed" # Folder name matches Gist ID
$JsonFile = "$GistFolder\docrefine_version.json"

# --- SAFETY CHECKS ---
if (-not (Test-Path $GistFolder)) {
    Write-Error "CRITICAL: Gist submodule folder not found at $GistFolder"
    exit
}

Write-Host "=== DOCREFINE PRO DEPLOYMENT PROTOCOL ===" -ForegroundColor Cyan
Write-Host "Target Version: $Version" -ForegroundColor Yellow
Write-Host "Channel: Permanent Drive Folder" -ForegroundColor Yellow

# 1. UPDATE GIST JSON
Write-Host "`n[1/4] Updating Update Signal (Gist)..." -ForegroundColor Green
$JsonContent = @{
    latest_version = $Version
    download_url = $DriveLink
} | ConvertTo-Json -Depth 2

$JsonContent | Set-Content -Path $JsonFile
Write-Host "JSON Updated locally."

# 2. PUSH GIST
Write-Host "`n[2/4] Publishing Signal to Cloud..." -ForegroundColor Green
Push-Location $GistFolder
git add docrefine_version.json
git commit -m "Update signal to $Version"
git push
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to push Gist."; Pop-Location; exit }
Pop-Location

# 3. TAG & PUSH MAIN REPO
Write-Host "`n[3/4] Triggering Build Pipeline..." -ForegroundColor Green
git add .
git commit -m "Release $Version"
git tag $Version
git push origin main --tags
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to push Repo."; exit }

# 4. INSTRUCTIONS
Write-Host "`n[4/4] DEPLOYMENT SEQUENCED." -ForegroundColor Cyan
Write-Host "---------------------------------------------------"
Write-Host "1. GitHub Actions is now building $Version."
Write-Host "2. Wait for the build to finish."
Write-Host "3. Download the artifacts (Win .exe / Mac .app)."
Write-Host "4. DRAG AND DROP them into this Drive Folder:"
Write-Host "   $DriveLink"
Write-Host "---------------------------------------------------"
Pause