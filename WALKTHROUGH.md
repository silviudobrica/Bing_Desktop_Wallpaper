# Bing Daily Wallpaper - User Guide & Walkthrough

## Installation Guide

### Option 1: Using the Installer (Recommended for Users)
1. **Download/Build**: Obtain `InstallBingWallpaper.exe` from the `dist` folder.
2. **Run**: Double-click the installer.
   - **No Admin Rights Needed**: The app installs to your local user profile, so you do not need Administrator privileges.
   - **Offline Ready**: The installer includes all necessary files. You do not need an internet connection to install (though you need one to fetch wallpapers!).
3. **Finish**: Click **Install**. The app will start automatically, and a shortcut will be added to your Startup folder (if selected).

### Option 2: Building from Source (For Developers)

If you want to modify the code or rebuild the installer:

1. **Install Python**: Ensure Python 3.10+ is installed.
2. **Install Dependencies**:
   ```powershell
   pip install -r requirements.txt
   pip install pyinstaller
   ```
### Run the Build Script

PowerShell:
   ```powershell
   .\build_installer.ps1
   ```

This script will:

*   **Compile** bing\_daily\_wallpaper.py into a standalone EXE.
    
*   **Bundle** that EXE inside installer.py.
    
*   **Output** the final installer to the dist folder.
    

Usage Guide
-----------

### System Tray

The app runs in the background. Look for the Bing icon in your system tray (near the clock).

**Right-Click Menu:**

*   **Preview / Gallery:** Opens the visual gallery of downloaded wallpapers.
    
*   **Check Now:** Forces an immediate update check.
    
*   **Interval:** Set how often to check (15 mins to 24 hours).
    
*   **Exit:** Quits the app completely.
    

Configuration & Data Locations
------------------------------

The application follows standard Windows practices for file storage:

*   **Wallpapers:** %USERPROFILE%\\Pictures\\Bing_(Your downloaded images are kept here so you can easily find them.)_
    
*   **Configuration:** %LOCALAPPDATA%\\Programs\\BingWallpaper\\config.json_(Stores your proxy settings and update interval preference.)_
    
*   **Logs:** %LOCALAPPDATA%\\Programs\\BingWallpaper\\logs\\_(Technical logs for troubleshooting.)_
    

Troubleshooting
---------------

### "I don't see the icon!"

*   Click the ^ (Show hidden icons) arrow in the taskbar.
    
*   Check if the process BingWallpaper.exe is running in Task Manager.
    

### Images aren't downloading

*   **Check Internet:** Ensure you can visit bing.com in your browser.
    
*   **Check Logs:** Open the log folder (see path above) and view bing\_wallpaper.log. Look for "ConnectionError" or "Proxy" errors.
    
*   **Proxy Settings:** If you are on a corporate network:
    
    *   The app attempts to auto-detect proxies.
        
    *   You can manually edit config.json to add your proxy URL and Port if auto-detection fails.
        

### App doesn't start on login

*   Run the installer again and ensure "Run at Startup" is checked.
    
*   Alternatively, press Win+R, type shell:startup, and ensure a shortcut to **Bing Wallpaper** exists there.