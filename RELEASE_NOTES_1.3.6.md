# Bing Daily Wallpaper v1.3.6

## Release Title
Bing Daily Wallpaper v1.3.6 - Updater Reliability and Repair Install Improvements

## Release Notes

This is a focused reliability release that improves the in-app updater behavior and makes troubleshooting update flow easier.

### Highlights

- Improved updater asset detection to support common installer naming variants.
- Added clearer updater logging for selected asset, download URL, and download destination.
- Added manual repair/reinstall option when already on the latest version.
- Rebuilt binaries with version metadata `1.3.6.0`.

### Improvements

#### Updater Asset Resolution
- The updater now accepts common `.exe` installer name patterns (not only one exact filename shape).
- This reduces cases where release assets exist but are not selected for download.

#### Better Updater Diagnostics
- Added explicit logs for:
  - selected release installer asset
  - direct installer URL usage
  - local target path in `%TEMP%\BingWallpaper\updates`
  - successful installer download

#### Repair/Reinstall Path (Up-to-date Case)
- When manually checking for updates and no newer version is available, users are now prompted to optionally download and run the installer anyway.
- This supports repair/reinstall scenarios without changing version numbers.

### Notes

- The updater still uses configured proxy settings for GitHub metadata and installer downloads.
- Automatic update behavior remains unchanged for newer-version detection; the repair/reinstall prompt applies to manual checks.

### Upgrade Guidance

- Existing users can use **Check for Updates** from tray or Settings.
- If already up to date, choose the repair/reinstall prompt to force installer download and launch.
- Downloaded installer path: `%TEMP%\BingWallpaper\updates`.
