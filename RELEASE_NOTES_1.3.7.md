# Bing Daily Wallpaper v1.3.7

## Release Title
Bing Daily Wallpaper v1.3.7 - Responsive Window Scaling and Layout Refinements

## Release Notes

This release improves the desktop UI responsiveness so the main window and settings window scale more naturally while keeping key content visible.

### Highlights

- Converted the main preview window to a stronger responsive grid layout with adaptive row sizing.
- Updated the bottom gallery strip to show the latest 5 wallpapers in fixed slots that scale with window size.
- Made the Settings window fully resizable with grid-based field expansion.
- Bumped app version to `1.3.7` and refreshed version resource metadata.

### Improvements

#### Main Window Responsive Behavior
- The window layout now consistently keeps these sections visible while resizing:
  - Top action buttons
  - Status section
  - Enlarged current image preview
  - Latest 5 image strip at the bottom
- Main preview and bottom thumbnail strip now scale proportionally with window size.
- Added adaptive layout tuning for smaller and larger window heights.

#### Thumbnail and Preview Scaling
- Bottom gallery now prioritizes the latest 5 wallpapers only (as displayed), each in a responsive fixed slot.
- Thumbnails and date labels dynamically resize to preserve readability and avoid clipping.
- Preview image redraw logic now follows available frame dimensions more closely.

#### Settings Window Responsiveness
- Settings changed from fixed-size modal to resizable modal.
- Inputs and action controls now use grid column weights so fields/buttons stretch correctly.
- Better use of horizontal space for proxy and CA bundle fields.

### Versioning

- Centralized app version updated to `1.3.7`.
- EXE version resource generation remains automated via `generate_version_info.py`.

### Notes

- Existing behavior for update checks, proxy handling, and startup registration is unchanged.
- This release focuses on UI scaling and usability during manual window resizing.
- The main preview window now opens at a slightly smaller default size while preserving responsive resize behavior.
