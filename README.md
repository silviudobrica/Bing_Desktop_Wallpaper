# Bing Daily Wallpaper

A lightweight Windows system tray application that automatically sets your desktop background to the daily Bing image.

## Features
- **Daily Updates**: Automatically checks for and downloads the latest Bing wallpaper (default: every 12 hours).
- **In-App Updater**: Checks GitHub releases, downloads the latest installer, launches it, and exits the app cleanly for upgrade.
- **Standalone App**: Runs entirely offline after installation. No Python or external dependencies required for the end user.
- **User-Level Install**: Installs to your local profile (`%LOCALAPPDATA%`) without requiring Administrator privileges.
- **Windows Installed Apps Integration**: Appears in Windows Apps/Installed Apps with proper uninstall registration.
- **System Tray Icon**: Sits quietly in your taskbar. Right-click for Preview/Gallery, Settings, Check Now, Check for Updates, interval control, and Exit.
- **Settings + Status UI**: Configure interval/proxy/startup in-app and view live health status.
- **Preview Gallery**: View and re-apply past wallpapers easily.
- **Lock Screen (optional)**: Also apply the daily image to the Windows lock screen. Not available on PCs where an organization policy enforces the lock screen image.
- **Backfill and Fallback**: Downloads wallpapers you missed while the PC was off (up to 7 days) and falls back to 1920x1080 when Bing has no UHD version.
- **Retention**: Optionally delete wallpapers older than N days to keep the Pictures folder small.
- **Region**: Choose the Bing market (en-US, de-DE, fr-FR, ...) for the daily image.
- **Resilient Checks**: After a failed check the app retries with increasing backoff (2 min to 1 h) and stays quiet for the first few failures, so a slow network at login no longer produces error pop-ups.
- **Windows Polish**: High-DPI awareness, keyboard shortcuts, and executable version metadata.
- **Smart Logging**: Auto-rotating logs keep your system clean.
- **Verified Updates**: Downloaded installers are checked against the SHA-256 digest GitHub publishes for the release asset before they run.

## Keyboard Shortcuts
- **Esc**: Hide/close active app window.
- **Ctrl+R** or **F5**: Trigger immediate wallpaper check.
- **Ctrl+,**: Open Settings.

## Update Flow
1. Run **Check for Updates** from tray menu or Settings.
2. If a newer version exists, choose **Download and install now**.
3. The app downloads the latest installer to `%TEMP%\BingWallpaper\updates`, launches it, and closes itself.
4. Complete upgrade in the installer UI.

## Requirements
* **End Users:** Windows 10 or 11 (No other software required).
* **Developers:** Python 3.10+ (Only if running from source or building the installer). Install tooling with `pip install -r requirements-dev.txt`; unit tests run with `python -m pytest`. CI on GitHub Actions runs the checks and builds the installer.

## Quick Start
1. Download the latest installer (`InstallBingWallpaper.exe`).
2. Run it and click **Install**.
3. The app will launch automatically and appear in your system tray.

See [WALKTHROUGH.md](WALKTHROUGH.md) for detailed installation, building, and troubleshooting instructions.