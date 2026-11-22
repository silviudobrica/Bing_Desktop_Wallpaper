# Bing Daily Wallpaper v1.4.4

## Release Title
Bing Daily Wallpaper v1.4.4 - Documentation for the Linux Flatpak

## Release Notes

This release has no functional changes. It brings the README and the walkthrough up to date with the Linux app.

### Documentation

#### README
- The introduction and requirements cover Linux as well as Windows.
- The Linux section describes notifications, says that "Check for Updates" is in the tray menu and in Settings, and explains that a per-user install updates without a password while a system-wide install asks for the administrator password.
- Install note: use `--user`, add `--reinstall` to install over an existing copy, and settings and wallpapers are kept.
- "Verified Updates" is listed under Features again; "Update Flow" and "Quick Start" are marked as Windows.

#### Walkthrough
- New "Installing and updating the Flatpak" section: per-user versus system-wide installs, `--reinstall`, switching from system-wide to per-user, what is kept across updates, and the one-time manual update that 1.4.1 and 1.4.2 need on a system-wide install.
- New "Linux troubleshooting" section: missing notifications, the "Directory not empty" uninstall error, and where the logs are.
- Notifications are described in the Linux section.
- "Publishing a release" lists every file that carries the version and mentions the Flatpak bundle attached to the release.

### Versioning
- Centralized app version updated to `1.4.4`.
- EXE version resource regenerated via `generate_version_info.py`.
