# platform_windows.py
"""Windows implementation of the service's Platform interface: desktop
wallpaper via Win32, lock screen via the WinRT API (PowerShell bridge),
PAC-based proxy detection from the registry, and the Startup shortcut."""
import base64
import ctypes
import os
import shutil
import subprocess
import sys
import time
import winreg
from pathlib import Path

import bing_core as core
from bing_service import Platform, log_msg

# Windows constants
SPI_SETDESKWALLPAPER = 0x0014  # 20
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDWININICHANGE = 0x02

# PowerShell bridge to the WinRT lock screen API. This runs as the current
# user, needs no admin rights and no extra Python packages.
#
# LockScreen.SetImageFileAsync is the call that works from an unpackaged
# process (every known working script uses it). The newer
# UserProfilePersonalizationSettings.TrySetLockScreenImageAsync is kept as
# a fallback only. The script contains no double quotes and is passed with
# -EncodedCommand, so no command-line quoting can break it.
LOCK_SCREEN_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$imagePath = '__IMAGE_PATH__'
try {
    [Windows.System.UserProfile.LockScreen,Windows.System.UserProfile,ContentType=WindowsRuntime] | Out-Null
    [Windows.System.UserProfile.UserProfilePersonalizationSettings,Windows.System.UserProfile,ContentType=WindowsRuntime] | Out-Null
    [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $methods = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 }
    $asTaskGeneric = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
    $asTaskAction = $methods | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
    function Await($op, $type) {
        $task = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($op))
        $task.Wait(-1) | Out-Null
        return $task.Result
    }
    function AwaitAction($op) {
        $task = $asTaskAction.Invoke($null, @($op))
        $task.Wait(-1) | Out-Null
    }
    function InnerMessage($err) {
        $ex = $err.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        return $ex.Message
    }
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($imagePath)) ([Windows.Storage.StorageFile])
    $method = ''
    $firstError = ''
    try {
        AwaitAction ([Windows.System.UserProfile.LockScreen]::SetImageFileAsync($file))
        $method = 'LockScreen.SetImageFileAsync'
    } catch {
        $firstError = InnerMessage $_
    }
    if (-not $method) {
        if (-not [Windows.System.UserProfile.UserProfilePersonalizationSettings]::IsSupported()) {
            [Console]::Error.WriteLine('LockScreen API failed (' + $firstError + ') and lock screen personalization is not supported on this Windows edition.')
            exit 2
        }
        $ok = Await ([Windows.System.UserProfile.UserProfilePersonalizationSettings]::Current.TrySetLockScreenImageAsync($file)) ([bool])
        if (-not $ok) {
            [Console]::Error.WriteLine('Windows refused the lock screen image (' + $firstError + '; blocked by policy, or unsupported format).')
            exit 3
        }
        $method = 'UserProfilePersonalizationSettings.TrySetLockScreenImageAsync'
    }
    Write-Output $method
    exit 0
} catch {
    [Console]::Error.WriteLine((InnerMessage $_))
    exit 1
}
"""


class WindowsPlatform(Platform):
    name = "windows"
    autosave_detected_proxy = True
    supports_lock_screen = True

    def __init__(self, data_dir, startup_shortcut, script_path=None):
        super().__init__()
        self.data_dir = Path(data_dir)
        self.startup_shortcut = Path(startup_shortcut)
        self.script_path = Path(script_path) if script_path else None

    # ------------------------------------------------------------ wallpaper
    def set_wallpaper(self, image_path):
        ctypes.windll.user32.SystemParametersInfoW(
            SPI_SETDESKWALLPAPER,
            0,
            str(image_path),
            SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE
        )

    # ------------------------------------------------------------ lock screen
    def lock_screen_policy(self):
        """Return a description of the policy forcing the lock screen image, or
        None when the user is free to change it. Managed (corporate) machines
        commonly pin the lock screen; WinRT then refuses our image."""
        checks = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Policies\Microsoft\Windows\Personalization", "LockScreenImage"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Policies\Microsoft\Windows\Personalization", "NoChangingLockScreen"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Policies\Microsoft\Windows\Personalization", "NoChangingLockScreen"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP", "LockScreenImagePath"),
        ]
        for hive, subkey, value_name in checks:
            try:
                with winreg.OpenKey(hive, subkey) as key:
                    value, _ = winreg.QueryValueEx(key, value_name)
                    if value:
                        return f"{value_name} policy ({value})"
            except OSError:
                continue
        return None

    def switch_lock_screen_to_picture_mode(self):
        """Turn off Windows Spotlight and the slideshow for the current user so
        the picture we set is what the lock screen actually shows. Writes only
        HKCU values the Settings app itself writes when "Picture" is chosen.
        Returns the names of the values that were changed."""
        changed = []
        for subkey, value_name, wanted in core.LOCK_SCREEN_PICTURE_MODE_VALUES:
            try:
                with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, subkey, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
                    try:
                        current, _ = winreg.QueryValueEx(key, value_name)
                    except OSError:
                        current = None
                    if current == wanted:
                        continue
                    if current is None and value_name == "SlideshowEnabled":
                        continue  # absent means off; do not create it
                    winreg.SetValueEx(key, value_name, 0, winreg.REG_DWORD, wanted)
                    changed.append(value_name)
            except OSError as e:
                log_msg(f"Could not update lock screen mode value {value_name}: {e}", "error")
        if changed:
            log_msg(f"Lock screen switched to Picture mode ({', '.join(changed)} set to 0)")
        return changed

    def prepare_lock_screen_copy(self, image_path):
        """Copy the wallpaper under a unique name. Windows refuses to set a lock
        screen image whose file name matches the previously set one, and
        "Check Now" re-applies the same bing_<date>.jpg; a fresh copy per call
        avoids that. Older copies are removed, the new one is kept until the
        next call so Windows can finish reading it."""
        copy_dir = self.data_dir / "lockscreen"
        copy_dir.mkdir(parents=True, exist_ok=True)
        copy_path = copy_dir / core.lock_screen_copy_name(image_path, time.time())
        shutil.copyfile(image_path, copy_path)
        for stale in core.select_stale_lock_screen_copies(copy_dir.glob("lockscreen_*"), copy_path):
            try:
                stale.unlink()
            except OSError:
                pass
        return copy_path

    def set_lock_screen(self, image_path):
        self.switch_lock_screen_to_picture_mode()
        copy_path = self.prepare_lock_screen_copy(Path(image_path))
        script = LOCK_SCREEN_SCRIPT.replace("__IMAGE_PATH__", str(copy_path).replace("'", "''"))
        encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
            capture_output=True, text=True, timeout=90,
            stdin=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if result.returncode == 0:
            method = (result.stdout or "").strip().splitlines()
            return True, (method[-1] if method else "WinRT")
        detail = (result.stderr or result.stdout or "").strip().splitlines()
        return False, (detail[0] if detail else f"exit code {result.returncode}")

    # ------------------------------------------------------------ proxy
    def system_proxies(self, url=None):
        # requests already honours HTTP(S)_PROXY from the environment; the
        # PAC-derived proxy is persisted to config.json on first run instead.
        return {}

    def get_pac_url_from_registry(self):
        locations = [
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"),
        ]
        for hive, subkey in locations:
            try:
                with winreg.OpenKey(hive, subkey) as key:
                    pac_url, _ = winreg.QueryValueEx(key, "AutoConfigURL")
                    if pac_url:
                        return pac_url
            except Exception:
                continue
        return None

    def get_proxy_from_pac(self, pac_url):
        try:
            text = self.service.fetch_text(pac_url, timeout=5) if self.service else None
            if text:
                return core.extract_pac_proxy(text)
        except Exception:
            pass
        return None

    def detect_proxy_values(self):
        pac_url = self.get_pac_url_from_registry()
        if not pac_url:
            return "", ""
        proxy = self.get_proxy_from_pac(pac_url)
        if not proxy:
            return "", ""
        host, port = core.parse_proxy_string(proxy)
        return host or "", port or ""

    # ------------------------------------------------------------ startup
    def is_startup_enabled(self):
        return self.startup_shortcut.exists()

    def set_startup_enabled(self, enabled):
        if enabled:
            target = Path(sys.executable)
            self.startup_shortcut.parent.mkdir(parents=True, exist_ok=True)

            args = ""
            if not getattr(sys, 'frozen', False) and self.script_path:
                args = f'"{str(self.script_path.resolve())}"'

            ps_script = f"""
            $ws = New-Object -ComObject WScript.Shell
            $s = $ws.CreateShortcut('{str(self.startup_shortcut)}')
            $s.TargetPath = '{str(target)}'
            $s.WorkingDirectory = '{str(self.data_dir)}'
            $s.Arguments = '{args}'
            $s.Save()
            """
            subprocess.run(["powershell", "-Command", ps_script], check=True, capture_output=True,
                           creationflags=subprocess.CREATE_NO_WINDOW)
        elif self.startup_shortcut.exists():
            self.startup_shortcut.unlink()

    # ------------------------------------------------------------ misc
    def open_path(self, path):
        os.startfile(path)
