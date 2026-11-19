# Bing Daily Wallpaper v1.3.10

## Release Title
Bing Daily Wallpaper v1.3.10 - Gallery Order Fix

## Release Notes

This release puts the wallpaper gallery back in date order.

### Fixes

#### Gallery order
- The "Latest 5" strip and its scrolling now list wallpapers by the date in their file name, newest first (today, yesterday, ...). Previously they were ordered by file modification time, and because backfill downloads older days after today's image, those older days appeared first.
- The lock screen option, when enabled before the session's first wallpaper was applied, now picks the newest downloaded image by date for the same reason.

### Versioning

- Centralized app version updated to `1.3.10`.
- EXE version resource regenerated via `generate_version_info.py`.
