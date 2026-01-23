# build_installer.ps1
# Ensure PyInstaller is present
if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    Write-Host "Installing PyInstaller..."
    pip install pyinstaller
}

$ScriptDir = $PSScriptRoot
Set-Location $ScriptDir

# 1. Clean previous builds
Remove-Item -Recurse -Force "dist", "build" -ErrorAction SilentlyContinue

# 2. Build the Main Application (BingWallpaper.exe)
# --noconsole: Run silently
Write-Host "Building Main Application..." -ForegroundColor Cyan
pyinstaller --noconfirm --onefile --clean --noconsole --name "BingWallpaper" `
    --add-data "_version.py;." `
    "bing_daily_wallpaper.py"

if (-not (Test-Path "dist\BingWallpaper.exe")) {
    Write-Error "Failed to build Main Application."
    exit 1
}

# 3. Build the Installer (InstallBingWallpaper.exe)
# We embed the newly built BingWallpaper.exe INSIDE the installer
Write-Host "Building Installer..." -ForegroundColor Cyan
pyinstaller --noconfirm --onefile --clean --noconsole --name "InstallBingWallpaper" `
    --add-data "dist\BingWallpaper.exe;." `
    --add-data "_version.py;." `
    "installer.py"

if (Test-Path "dist\InstallBingWallpaper.exe") {
    Write-Host "Build Complete!" -ForegroundColor Green
    Write-Host "Installer: dist\InstallBingWallpaper.exe"
} else {
    Write-Error "Failed to build Installer."
}