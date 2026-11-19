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
   pip install -r requirements-dev.txt
   ```
### Run the Build Script

PowerShell:
   ```powershell
   .\build_installer.ps1
   ```

Optional signing (if you have a PFX certificate):
   ```powershell
   $env:SIGN_CERT_PATH = "C:\path\to\codesign.pfx"
   $env:SIGN_CERT_PASSWORD = "your-password"
   .\build_installer.ps1
   ```

If signing variables are not set, the build still succeeds and skips signing.

Optional pre-commit check (recommended for contributors):
   ```powershell
   pip install pre-commit
   pre-commit install
   ```

This enables a local hook that verifies `_version.py` and `file_version_info.txt` stay in sync before each commit.
It also compiles the core Python files to catch syntax errors before commit.
Additionally, a lightweight smoke check validates packaging-critical wiring (spec metadata and build script version generation), and the unit tests run.

### Tests and CI

Pure logic (version comparison, Bing API parsing, proxy parsing, retry backoff, retention, installer asset and checksum selection) lives in `bing_core.py` with no GUI or network dependencies. Run the unit tests with:

   ```powershell
   python -m pytest -q
   ```

GitHub Actions (`.github/workflows/ci.yml`) runs the compile step, version sync check, smoke check and unit tests on every push and pull request, then builds the installer on `windows-latest` and uploads `InstallBingWallpaper.exe` plus its `.sha256` file as a workflow artifact.

### Publishing a release

Releases are published by CI when a version tag is pushed. The tag must match `_version.py`:

1. Bump `__version__` in `_version.py`, run `python generate_version_info.py`, and write `RELEASE_NOTES_<version>.md`.
2. Commit, then tag and push:

   ```powershell
   git tag v1.3.9
   git push origin main v1.3.9
   ```

3. The `release` job waits for the tests and the build, verifies that the tag matches `_version.py`, creates the GitHub release named `Bing Daily Wallpaper v<version>` with `RELEASE_NOTES_<version>.md` as its body (GitHub-generated notes if the file is missing), and attaches the CI-built `InstallBingWallpaper.exe` and its `.sha256`. Tags containing a `-` (for example `v1.4.0-beta1`) are marked as pre-releases.

The in-app updater picks up the new release from the GitHub API and verifies the download against the digest GitHub computes for the attached asset.

### Release checksums

`build_installer.ps1` writes `dist\InstallBingWallpaper.exe.sha256`. The in-app updater verifies the downloaded installer against the SHA-256 digest GitHub publishes for release assets; if a release has no digest, it looks for a 64-character hash next to the asset name in the release notes. A mismatch discards the download and reports an error.

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
*   **Settings:** Opens full app settings (interval, proxy, startup, update check, quick folder actions).
    
*   **Check Now:** Forces an immediate update check.

*   **Check for Updates:** Checks GitHub for a newer app version and offers direct install.
    
*   **Interval:** Set how often to check (15 mins to 24 hours).
    
*   **Exit:** Quits the app completely.

### Keyboard Shortcuts

*   **Esc:** Hide/close active app window.
*   **Ctrl+R** or **F5:** Trigger immediate wallpaper check.
*   **Ctrl+,**: Open Settings.
    

Configuration & Data Locations
------------------------------

The application follows standard Windows practices for file storage:

*   **Wallpapers:** %USERPROFILE%\\Pictures\\Bing_(Your downloaded images are kept here so you can easily find them.)_
    
*   **Configuration:** %LOCALAPPDATA%\\Programs\\BingWallpaper\\config.json_(Stores your proxy settings and update interval preference.)_

Wallpaper options (Settings > Wallpaper)
---------------------------------------

*   **Bing market:** Region of the daily image, e.g. `en-US`, `de-DE`, `fr-FR`. Default `en-US`.
*   **Keep images for (days):** Wallpapers older than this are deleted after each check. `0` keeps everything (default). The image currently in use is never deleted.
*   **Backfill missed days:** Downloads up to 7 recent images you missed while the PC was off, within the retention window. Default `7`, `0` disables.
*   **Also set the Windows lock screen image:** Off by default. Uses the Windows lock screen API as the current user (no admin rights). Enabling it switches the lock screen from Windows Spotlight (or a slideshow) to Picture mode, because Windows ignores the picture while Spotlight is active; to go back to Spotlight, untick the option and pick Windows Spotlight again in Settings > Personalization > Lock screen. The change is visible the next time you lock the PC. If your organization enforces a lock screen image by policy, the app detects this and reports it in the status panel instead of silently failing.

The corresponding `config.json` keys are `market`, `keep_images_days`, `backfill_days` and `set_lock_screen`.

If Bing does not offer a UHD version of an image, the app falls back to 1920x1080 automatically. After a failed check the app retries after 2, 5, 15, 30 and then 60 minutes instead of waiting a full interval, and only shows an error notification from the third consecutive failure on.
    
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

### Upgrade / Update behavior

*   From tray or Settings, choose **Check for Updates**.
*   If a new version is available and you choose install, the app downloads installer to `%TEMP%\BingWallpaper\updates`.
*   The app launches the installer and exits itself to avoid file lock issues during upgrade.

### Uninstall behavior

*   The app registers in Windows Installed Apps for normal uninstall flow.
*   Uninstall removes app files, startup/start menu/desktop shortcuts, and installed-apps registry entry.
*   If a file is briefly locked, delayed cleanup is scheduled automatically.

Phase 4 Verification Checklist
-----------------------------

After building and installing, verify the following:

*   **DPI / UI:** Open Settings and Preview on high DPI scaling (125%/150%+). Windows should be crisp and usable.
*   **Keyboard:** In Preview window, test: `Esc` (hide), `Ctrl+R` or `F5` (Check Now), `Ctrl+,` (open Settings).
*   **Updater:** Use **Check for Updates** from tray or Settings. If a newer release exists, choose automatic download/install.
*   **Metadata:** In Explorer, open EXE Properties -> Details and confirm Product/File version fields are populated.
*   **Signing (if configured):** In EXE Properties -> Digital Signatures, verify signature and timestamp are present.