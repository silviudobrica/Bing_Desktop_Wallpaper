# Bing Daily Wallpaper v1.3.8

## Release Title
Bing Daily Wallpaper v1.3.8 - Startup Download Race Fix and Stability Improvements

## Release Notes

This release fixes a startup race that produced a spurious "Image download failed" error on most app launches, and hardens several internal paths around threading and file writes.

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

### Other

- The tray "Check Now" action now runs on a daemon thread so it cannot delay application exit.
- The fallback version string in `installer.py` was brought in line with `_version.py`.

### Versioning

- Centralized app version updated to `1.3.8`.
- EXE version resource regenerated via `generate_version_info.py`.
