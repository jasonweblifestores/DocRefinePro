# build_debug.ps1
Write-Host "🚧 BUILDING DEBUG VERSION (CONSOLE ENABLED)..." -ForegroundColor Yellow

# Clean previous builds to ensure no ghost files remain
Remove-Item -Path "build", "dist" -Recurse -ErrorAction SilentlyContinue

# 1. --console: Keeps the black terminal window open so we can see errors.
# 2. --collect-all PySide6: Brute-forces every Qt file into the build (fixes missing DLLs).
# 3. --debug=all: Tells the bootloader to talk to us.

pyinstaller --noconfirm --onedir --console --clean `
    --name "DocRefine_Debug" `
    --collect-all "PySide6" `
    --hidden-import "PySide6.QtXml" `
    --hidden-import "PySide6.QtNetwork" `
    --paths "." `
    main.py

Write-Host "✅ Build Complete." -ForegroundColor Green
Write-Host "👉 Please run: .\dist\DocRefine_Debug\DocRefine_Debug.exe" -ForegroundColor Cyan