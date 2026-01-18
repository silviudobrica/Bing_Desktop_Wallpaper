# Check if PyInstaller is installed
if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    Write-Host "PyInstaller not found. Installing..."
    pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install PyInstaller. Please ensure pip is in your PATH."
        exit 1
    }
}

$ScriptDir = $PSScriptRoot
Set-Location $ScriptDir

# Build the installer
# --onefile: Create a single EXE
# --uac-admin: Request Admin privileges on launch
# --noconsole: Hide console window (for GUI)
# --clean: Clean cache
# --noconfirm: Don't ask to overwrite
# --add-data: Bundle dependent files

Write-Host "Building Installer EXE..."

# Removed --uac-admin and --noconsole to fix bootloader issues.
# The app will handle window hiding or just show a console for log output (which is fine for an installer).
pyinstaller --noconfirm --onefile --clean --name "InstallBingWallpaper" `
    --add-data "bing_daily_wallpaper.py;." `
    --add-data "requirements.txt;." `
    "installer.py"

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build Successful!" -ForegroundColor Green
    Write-Host "Installer is located at: $(Join-Path $ScriptDir 'dist\InstallBingWallpaper.exe')"
} else {
    Write-Error "Build Failed."
}
