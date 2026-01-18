# Bing Daily Wallpaper - User Guide & Walkthrough

## Installation Guide

### Option 1: Using the Installer (Recommended)
1. **Build the Installer**:
   Open PowerShell in the project directory and run:
   ```powershell
   .\build_installer.ps1
   ```
   This will create a standalone executable in the `dist` folder.

2. **Run the Installer**:
   Navigate to the `dist` folder and run `InstallBingWallpaper.exe`.
   - Click **Install**.
   - Ensure "Create Startup Shortcut" is checked to have the app run automatically when you log in.
   - The installer will handle setting up Python if it's missing (requires internet).

### Option 2: Manual Run (For Developers)
1. Install Python 3.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the script:
   ```bash
   python bing_daily_wallpaper.py
   ```

## Usage Guide

### System Tray
Once running, you will see a small icon in your system tray (bottom right, near the clock). You may need to click the `^` arrow to see hidden icons.

- **Double-Click (or Right-Click -> Preview)**: Opens the Image Gallery window.
- **Right-Click -> Check for Updates**: Forces an immediate check for a new wallpaper.
- **Right-Click -> Exit**: Closes the application completely.

### Image Gallery
The Preview window shows the currently set wallpaper.
- **Bottom List**: Shows previously downloaded images.
- **Click an Image**: Instantly sets that image as your desktop background.

## Troubleshooting
- **Logs**: If the app isn't working, check the log file located at:
  `%USERPROFILE%\Pictures\Bing\bing_wallpaper.log`
- **Startup**: If the app doesn't start on login, check your "Startup" folder or re-run the installer.
