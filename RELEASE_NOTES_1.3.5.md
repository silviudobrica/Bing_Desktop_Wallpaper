# Bing Daily Wallpaper v1.3.5

## Release Title
Bing Daily Wallpaper v1.3.5 - Updater, Windows Integration, and Release Hardening

## Release Notes

This release completes the major product hardening pass and improves the app's Windows-native behavior, reliability, and release pipeline.

### Highlights

- Added a full in-app settings and status experience.
- Added in-app update checks with direct download-and-run installer handoff.
- Improved Windows install/uninstall integration and cleanup behavior.
- Improved upgrade reliability when replacing app binaries.
- Added high-DPI and keyboard usability improvements.
- Added executable version metadata and optional code-signing support in build pipeline.
- Added repository pre-commit safeguards for version metadata, syntax, and packaging wiring.

### New and Improved

#### Settings and Status UI
- Added a dedicated Settings window for interval, proxy, startup toggle, and quick actions.
- Added a live status panel showing last check, last success/error, current image, interval, and startup state.
- Added shortcut support (`Esc`, `Ctrl+R`, `F5`, `Ctrl+,`) for faster navigation.

#### Updater Experience
- Added manual and automatic GitHub release checks.
- Added update notifications and prompts.
- Added direct update flow: download latest `InstallBingWallpaper.exe`, launch installer, exit app cleanly.

#### Windows Integration
- Added Windows Installed Apps registration (user scope) with uninstall metadata.
- Added Start Menu shortcut support.
- Improved uninstall cleanup (shortcuts, registry, install directory) with retry/fallback handling.

#### Installer and Upgrade Reliability
- Added process-stop handling and atomic replace + retry when updating `BingWallpaper.exe`.
- Reduced upgrade failures caused by transient file locks.

#### Build and Release Pipeline
- Added EXE version resource embedding through spec files.
- Added automatic `file_version_info.txt` generation from `_version.py`.
- Added optional signing path in build script (uses `signtool` + signing environment variables when available).

#### Contributor Safeguards
- Added pre-commit checks:
  - Version metadata sync check
  - Core Python syntax compile check
  - Packaging smoke check

### Notes
- Signing is optional and skipped automatically when `signtool` or signing environment variables are not available.
- Update checks and updater download require internet access.

### Upgrade Guidance
- Existing users can update through **Check for Updates** in the app or by running the latest installer manually.
- If the updater cannot fetch release metadata or installer assets, the app falls back to release page guidance.
