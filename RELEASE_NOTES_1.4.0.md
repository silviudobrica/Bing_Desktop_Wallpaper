# Bing Daily Wallpaper v1.4.0

## Release Title
Bing Daily Wallpaper v1.4.0 - Linux (KDE Plasma) Flatpak and shared core

## Release Notes

This release brings the app to Linux as a Flatpak, KDE Plasma first, with the same window, settings and tray menu as on Windows. To get there the Windows app was split into a shared, tested service and a thin Windows layer.

### New: Linux app (Flatpak, PySide6)

- `BingWallpaper.flatpak` is attached to this release. Install with `flatpak install --user BingWallpaper.flatpak` (needs the Flathub remote for the KDE 6.11 runtime).
- Wallpaper on KDE Plasma through the shell scripting interface, on GNOME through gsettings or the XDG Wallpaper portal, elsewhere through the portal.
- Lock screen on KDE (`kscreenlockerrc`) and GNOME (gsettings / portal).
- System proxy support: the proxy is resolved before every request through the ProxyResolver portal, KDE's `kioslaverc` (manual, PAC and environment modes), GNOME's settings or the environment. A changed system proxy is used by the next download without a restart. Settings can still pin an explicit proxy and CA bundle.
- Autostart through the Background portal, notifications through the Notification portal, single-instance handling, tray icon with the current wallpaper.
- In-app updates like on Windows: "Check for Updates" downloads the release's Flatpak bundle, verifies its SHA-256, installs it through the host's `flatpak install` and restarts the app. From a plain checkout the dialog opens the release page.
- Runs from a plain checkout too: `pip install -r requirements-linux.txt && python3 bing_qt_app.py`.

### Changed: shared core

- Download, backfill, retention, configuration, status and the check loop moved from the Tk app into `bing_service.py`, behind a small platform interface (`platform_windows.py`, `platform_linux.py`). The Windows app's behaviour is unchanged, with two small fixes: "Wallpaper updated" is no longer announced when setting the wallpaper failed, and the deferred update-error dialog no longer references an unbound variable.
- Proxy lookup goes through the service on both platforms; on Windows the PAC-detected proxy is still saved to `config.json` on first run.

### Maintenance

- 49 new tests: the service against a fake platform and HTTP session, the Linux platform's proxy parsers, lock screen config and autostart, and the Qt UI offscreen. 99 tests in total.
- CI: a Linux job (lint, offscreen Qt tests) and a Flatpak build job on the Flathub KDE 6.11 image; the release job attaches the bundle and its SHA-256 next to the Windows installer.
- Flatpak manifest on `org.kde.Platform` 6.11 with `io.qt.PySide.BaseApp` 6.11; pinned Python dependencies in `flatpak/python3-requirements.json`.

### Versioning

- Centralized app version updated to `1.4.0`.
- EXE version resource regenerated via `generate_version_info.py`.
