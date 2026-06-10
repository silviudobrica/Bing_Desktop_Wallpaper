# build_installer.ps1
# Ensure PyInstaller is present
if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    Write-Host "Installing PyInstaller..."
    pip install pyinstaller
}

$SignTool = Get-Command signtool -ErrorAction SilentlyContinue

$ScriptDir = $PSScriptRoot
Set-Location $ScriptDir

# Generate EXE version resource metadata from centralized version file.
Write-Host "Generating version resource metadata..." -ForegroundColor Cyan
python .\generate_version_info.py
if (-not (Test-Path "file_version_info.txt")) {
    Write-Error "Failed to generate file_version_info.txt"
    exit 1
}

# 1. Clean previous builds
Remove-Item -Recurse -Force "dist", "build" -ErrorAction SilentlyContinue

# 2. Build the Main Application (BingWallpaper.exe)
Write-Host "Building Main Application..." -ForegroundColor Cyan
pyinstaller --noconfirm --clean "BingWallpaper.spec"

if (-not (Test-Path "dist\BingWallpaper.exe")) {
    Write-Error "Failed to build Main Application."
    exit 1
}

# 3. Build the Installer (InstallBingWallpaper.exe)
Write-Host "Building Installer..." -ForegroundColor Cyan
pyinstaller --noconfirm --clean "InstallBingWallpaper.spec"

function Sign-Executable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath
    )

    if (-not (Test-Path $ExePath)) {
        Write-Host "Skipping sign, file missing: $ExePath" -ForegroundColor Yellow
        return
    }

    if (-not $SignTool) {
        Write-Host "signtool not found, skipping signing." -ForegroundColor Yellow
        return
    }

    if (-not $env:SIGN_CERT_PATH -or -not $env:SIGN_CERT_PASSWORD) {
        Write-Host "SIGN_CERT_PATH or SIGN_CERT_PASSWORD not set, skipping signing." -ForegroundColor Yellow
        return
    }

    Write-Host "Signing $ExePath ..." -ForegroundColor Cyan
    & $SignTool.Source sign `
        /f $env:SIGN_CERT_PATH `
        /p $env:SIGN_CERT_PASSWORD `
        /fd SHA256 `
        /tr http://timestamp.digicert.com `
        /td SHA256 `
        "$ExePath"

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Signing failed for $ExePath"
        exit 1
    }
}

Sign-Executable -ExePath "dist\BingWallpaper.exe"
Sign-Executable -ExePath "dist\InstallBingWallpaper.exe"

if (Test-Path "dist\InstallBingWallpaper.exe") {
    Write-Host "Build Complete!" -ForegroundColor Green
    Write-Host "Installer: dist\InstallBingWallpaper.exe"
} else {
    Write-Error "Failed to build Installer."
}