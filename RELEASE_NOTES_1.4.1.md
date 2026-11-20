# Bing Daily Wallpaper v1.4.1

## Release Title
Bing Daily Wallpaper v1.4.1 - Flatpak lock screen fix and in-app updates

## Release Notes

### Fixes

#### Lock screen on KDE Plasma (Flatpak)
- The Flatpak wrote the lock screen image into its private sandbox config directory, which Plasma never reads, and reported success. It now rewrites the host's `~/.config/kscreenlockerrc` in place, so the image shows on the next lock. If that file does not exist yet, the status panel says to open System Settings > Screen Locking once so Plasma creates it.

#### Windows Settings dialog
- The window is wider, so the "Custom minutes (optional)" label is no longer clipped.

### New: in-app updates for the Flatpak
- "Check for Updates" (daily in the background, or from the tray and Settings) now offers to download the release's `BingWallpaper.flatpak`, verifies it against the SHA-256 GitHub publishes for the asset, installs it through the host's `flatpak install` (user or system, matching where the app is installed) and restarts the app. This mirrors the Windows updater.
- The manifest grants `org.freedesktop.Flatpak` for this. A future Flathub submission would drop it and rely on store updates.
- The download and checksum step is shared with the Windows updater.

### Maintenance
- 10 new tests (release asset selection, verified downloads, host install command, the Qt update flow). 109 in total.
- Linux platform tests run in an isolated home directory.

### Versioning
- Centralized app version updated to `1.4.1`.
- EXE version resource regenerated via `generate_version_info.py`.
