# Bing Daily Wallpaper v1.3.9

## Release Title
Bing Daily Wallpaper v1.3.9 - Lock Screen Fix

## Release Notes

This release fixes the lock screen option introduced in 1.3.8, which reported success but left the lock screen unchanged on most PCs.

### Fixes

#### Lock screen image was never shown
- Windows kept showing Windows Spotlight: setting a lock screen picture does not switch the lock screen mode, and while Spotlight (the Windows 11 default) or a slideshow is active the stored picture is ignored. Enabling the option now switches the current user's lock screen to Picture mode, exactly as the Settings app does when you choose "Picture". The Settings checkbox says so.
- The app now uses the `LockScreen.SetImageFileAsync` API, the call known to work from a plain (unpackaged) process, and keeps `TrySetLockScreenImageAsync` only as a fallback. The log records which one was used.
- Windows refuses a lock screen image whose file name equals the previously set one, which broke re-applying today's image via "Check Now". The app now hands Windows a uniquely named copy (under `%LOCALAPPDATA%\Programs\BingWallpaper\lockscreen`) and removes older copies on the next call.
- Enabling the option in Settings before the first wallpaper of the session was applied did nothing until the next day; it now applies the most recent downloaded image immediately.
- The PowerShell bridge is passed as an encoded command with no inherited stdin, removing two ways the call could fail inside the windowed EXE.

### Versioning

- Centralized app version updated to `1.3.9`.
- EXE version resource regenerated via `generate_version_info.py`.
