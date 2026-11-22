# Bing Daily Wallpaper v1.4.2

## Release Title
Bing Daily Wallpaper v1.4.2 - Notifications in the Flatpak

## Release Notes

### Fixes

#### Notifications never appeared in the Flatpak
- The app sent notifications through the tray icon, which Qt delivers to `org.freedesktop.Notifications`. The sandbox cannot reach that service, so every notification ("Wallpaper updated", "Update available", failures) was dropped while the app believed it had been shown. In the Flatpak, notifications now always go through the XDG Notification portal, which needs no extra permission. Outside the Flatpak the tray icon is still used first, with the portal and `notify-send` as fallbacks.
- Notifications are sent with normal priority instead of low, so desktops that hide low-priority popups still show them.
- A notification that cannot be delivered is now written to the log instead of disappearing silently.

### Maintenance
- 4 new tests (portal-first delivery in the Flatpak, tray-first outside it, notification priority, failure reporting). 113 in total.

### Versioning
- Centralized app version updated to `1.4.2`.
- EXE version resource regenerated via `generate_version_info.py`.
