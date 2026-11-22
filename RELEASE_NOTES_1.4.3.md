# Bing Daily Wallpaper v1.4.3

## Release Title
Bing Daily Wallpaper v1.4.3 - Flatpak in-app update fix and a Check for Updates button

## Release Notes

### Fixes

#### In-app update failed on system-wide Flatpak installs
- "Check for Updates" downloaded the bundle and then failed with "Flatpak system operation InstallBundle not allowed for user". A system-wide install needs the administrator password, and the updater ran `flatpak install` with `--noninteractive`, which forbids the password prompt. For system-wide installs the updater no longer passes that flag, so the desktop can ask for the password. Per-user installs are unchanged and still need no password.
- Versions 1.4.1 and 1.4.2 cannot update themselves on a system-wide install. To get this release, download `BingWallpaper.flatpak` from the release page and run `flatpak install --system --reinstall BingWallpaper.flatpak`. Settings and downloaded wallpapers are kept.

### New: Check for Updates in Settings (Linux)
- The Settings dialog of the Linux app has a "Check for Updates" button, like the Windows app. Until now the check was only in the tray icon's menu.
- The update dialogs open on top of the Settings dialog when the check is started from there.

### Maintenance
- 2 new tests (system-wide install command, the Settings button). 115 in total.

### Versioning
- Centralized app version updated to `1.4.3`.
- EXE version resource regenerated via `generate_version_info.py`.
