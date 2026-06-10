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
- **Windows Polish**: High-DPI awareness, keyboard shortcuts, and executable version metadata.
- **Smart Logging**: Auto-rotating logs keep your system clean.

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
* **Developers:** Python 3.10+ (Only if running from source or building the installer).

## Quick Start
1. Download the latest installer (`InstallBingWallpaper.exe`).
2. Run it and click **Install**.
3. The app will launch automatically and appear in your system tray.

See [WALKTHROUGH.md](WALKTHROUGH.md) for detailed installation, building, and troubleshooting instructions.