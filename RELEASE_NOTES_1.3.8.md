# Bing Daily Wallpaper v1.3.8

## Release Title
Bing Daily Wallpaper v1.3.8 - Lock Screen, Backfill, Retention and Startup Race Fix

## Release Notes

This release adds opt-in lock screen support, backfill of missed days, image retention and a Bing market setting. It also fixes a startup race that produced a spurious "Image download failed" error on most app launches, and hardens several internal paths around threading and file writes.

### Fixes

#### Startup download race (WinError 32)
- On launch, the app previously started two wallpaper checks at the same time. Both downloaded the day's image to the same temporary file, and the loser failed with `[WinError 32] The process cannot access the file because it is being used by another process`, triggering a false "Operation failed" notification even though the wallpaper was set correctly.
- Wallpaper checks are now serialized: a check that starts while another is running is skipped and logged.
- The redundant startup thread was removed; the background loop performs the initial check itself.
- Each download now writes to a unique temporary file, and leftover partial downloads are cleaned up at startup.
- Image verification failures are now logged with the underlying error instead of being silently discarded.

#### Thread safety for the UI
- Status panel updates triggered from worker threads (wallpaper checks, update checks, HTTPS test) are now marshalled onto the Tk main thread. Tkinter is not thread-safe, and direct updates from workers could hang or crash the UI.

#### Single-instance detection
- The mutex check now reads the Win32 last-error value reliably via `ctypes.get_last_error()`, so a second instance is detected consistently.

#### Atomic configuration saves
- `config.json` is written to a temporary file and atomically replaced, so an interruption mid-write can no longer leave a truncated or empty configuration file.

### New Features

#### Lock screen support (opt-in)
- New Settings option **Also set the Windows lock screen image**. Uses the Windows personalization API as the current user; no admin rights or extra dependencies.
- On PCs where an organization policy enforces the lock screen image, the app detects the policy and reports it instead of failing silently.

#### Backfill of missed days and resolution fallback
- Each check now fetches the last 8 days from Bing and downloads any images missing locally (configurable 0-7 days, default 7), so days when the PC was off are no longer lost.
- When Bing has no UHD version of an image, the app falls back to 1920x1080 instead of failing.

#### Image retention
- New **Keep images for (days)** setting deletes older wallpapers after each check. Default `0` keeps everything, as before. The current wallpaper is never deleted.

#### Bing market
- New **Bing market** setting (e.g. `en-US`, `de-DE`, `fr-FR`) replaces the hard-coded US region.

#### Quieter, faster recovery from failed checks
- After a failed check the app retries after 2, 5, 15, 30 and then 60 minutes instead of waiting a full interval.
- No error notification for the first two consecutive failures, which covers the network not being ready right after login. The status panel still shows the error.

#### Verified update downloads
- The in-app updater now verifies the downloaded installer against the SHA-256 digest GitHub publishes for the release asset (or a hash in the release notes). A mismatch discards the file and reports an error.
- `build_installer.ps1` writes `InstallBingWallpaper.exe.sha256` next to the installer for inclusion in release notes.

### Maintenance

- Pure logic moved into `bing_core.py` (version comparison, Bing API parsing, proxy/PAC parsing, retry backoff, retention, installer asset and checksum selection), shared by the app and the installer.
- New pytest suite in `tests/` (45 tests), wired into the pre-commit hooks.
- New GitHub Actions workflow: checks and tests on every push/PR, then builds and uploads the installer as an artifact.
- Dependency pins updated to match current releases: `requests` 2.31.0 -> 2.34.2, `Pillow` 10.2.0 -> 12.1.0 (the old Pillow pin predates the fix for CVE-2024-28219). Build tooling pinned in new `requirements-dev.txt`.
- Pre-release suffixes such as `1.3.8-beta2` are now compared as `1.3.8` by the update check (previously parsed as `1.3.82`).

### Other

- The tray "Check Now" action now runs on a daemon thread so it cannot delay application exit.
- The fallback version string in `installer.py` was brought in line with `_version.py`.

### Versioning

- Centralized app version updated to `1.3.8`.
- EXE version resource regenerated via `generate_version_info.py`.
