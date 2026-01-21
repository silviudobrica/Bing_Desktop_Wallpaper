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
   - **Proxy Settings**: If you are behind a corporate firewall, enter your Proxy URL (e.g., `proxy.example.com`) and Port (e.g., `80`) before installing. You can use the **Detect System Proxy** button to automatically fill these fields.
   - Click **Install**.
   - Ensure "Create Startup Shortcut" is checked to have the app run automatically when you log in.
   - The installer will handle setting up Python if it's missing (requires internet).

### Option 2: Manual Run (For Developers)
1. Install Python 3 (Python 3.11+ recommended).
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
- **Right-Click -> Check Interval**: Change how often the app checks for new wallpapers (Default: 12 Hours).
  - **Presets**: 15 min, 30 min, 1h, 4h, 6h, 12h, 24h.
  - **Custom**: Enter a custom duration in minutes.
  - **Disabled**: Disable automatic background checks (manual updates only).
- **Right-Click -> Exit**: Closes the application completely.

### Image Gallery
The Preview window shows the currently set wallpaper.
- **Bottom List**: Shows previously downloaded images.
- **Click an Image**: Instantly sets that image as your desktop background.

### Configuration
The application stores its configuration in:
`%LOCALAPPDATA%\Programs\BingWallpaper\config.json`

Example configuration:
```json
{
  "check_interval_minutes": 720,
  "proxy_url": "proxy.example.com",
  "proxy_port": "80"
}
```

## Troubleshooting
- **Logs**: If the app isn't working, check the log file located at:
  `%USERPROFILE%\Pictures\Bing\bing_wallpaper.log`
- **Startup**: If the app doesn't start on login, check your "Startup" folder or re-run the installer.
- **Proxy**: If images aren't downloading, verify your proxy settings in `config.json` or reinstall with the correct proxy details.
