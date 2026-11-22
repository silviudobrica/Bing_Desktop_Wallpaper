# Bing Daily Wallpaper v1.4.5

## Release Title
Bing Daily Wallpaper v1.4.5 - In-app update for system-wide Flatpak installs

## Release Notes

### Fixes

#### In-app update still failed on system-wide Flatpak installs
- The fix in 1.4.3 did not work. Started without a terminal, `flatpak install` never lets the desktop ask for the administrator password, with or without `--noninteractive`, so the update kept failing with "Flatpak system operation InstallBundle not allowed for user".
- For a system-wide install the updater now runs the install through `pkexec`, which shows the desktop's password prompt and then installs the bundle. Per-user installs are unchanged and still update without a password.
- Cancelling the password prompt now reports "The password prompt was cancelled." instead of raw command output.
- Versions 1.4.1 to 1.4.4 cannot update themselves on a system-wide install. To get this release, download `BingWallpaper.flatpak` from the release page and run `flatpak install --system --reinstall BingWallpaper.flatpak`, or switch to a per-user install (see the walkthrough). Settings and downloaded wallpapers are kept.

### Documentation
- The walkthrough describes the `pkexec` step, names 1.4.5 as the first version that can update a system-wide install, and explains the "Directory not empty" error (the same version reinstalled more than once while the app was running).

### Maintenance
- The system-wide install test now covers the `pkexec` command and the cancelled prompt. 115 tests in total.

### Versioning
- Centralized app version updated to `1.4.5`.
- EXE version resource regenerated via `generate_version_info.py`.
