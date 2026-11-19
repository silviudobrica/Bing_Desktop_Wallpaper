# bing_daily_wallpaper.py
import os
import sys
import ctypes
import requests
import datetime
import time
import threading
import logging
import json
import subprocess
import webbrowser
import tempfile
import shutil
import base64
from pathlib import Path
from logging.handlers import RotatingFileHandler
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from PIL import Image, ImageTk, UnidentifiedImageError
import pystray
from pystray import MenuItem as item
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog
import winreg
import re
from requests.exceptions import SSLError
import bing_core as core
from bing_core import (
    DEFAULT_MARKET, MARKET_PRESETS, MAX_BACKFILL_DAYS,
)

def check_single_instance():
    mutex_name = "Local\\BingWallpaperTrayAppMutex"
    # use_last_error=True makes ctypes capture GetLastError right after the call,
    # so it cannot be clobbered by other Win32 calls made by the runtime.
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
    mutex = kernel32.CreateMutexW(None, False, mutex_name)
    if ctypes.get_last_error() == 183: # ERROR_ALREADY_EXISTS
        print("Application is already running.")
        sys.exit(0)
    return mutex # Keep a reference so it isn't garbage collected

# Windows constants
SPI_SETDESKWALLPAPER = 0x0014 # 20
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDWININICHANGE = 0x02

# Import centralized version
try:
    from _version import __version__ as VERSION
except ImportError:
    VERSION = "1.3.10"

# Configuration
# Bing API, retry and retention constants live in bing_core.py.
RELEASES_API = "https://api.github.com/repos/SilviuDobrica/Bing_Desktop_Wallpaper/releases/latest"
APP_NAME = "BingWallpaper"

# Paths
DATA_DIR = Path(os.environ["LOCALAPPDATA"]) / "Programs" / APP_NAME
LOG_DIR = DATA_DIR / "logs"
CONFIG_FILE = DATA_DIR / "config.json"
IMAGE_DIR = Path(os.environ["USERPROFILE"]) / "Pictures" / "Bing"
STARTUP_SHORTCUT = Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup" / "Bing Wallpaper.lnk"

# Ensure directories exist
LOG_DIR.mkdir(parents=True, exist_ok=True)
IMAGE_DIR.mkdir(parents=True, exist_ok=True)

# Setup Rotating Logging
LOG_FILE = LOG_DIR / "bing_wallpaper.log"
handler = RotatingFileHandler(LOG_FILE, maxBytes=5*1024*1024, backupCount=2, encoding='utf-8')
logging.basicConfig(
    handlers=[handler],
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(threadName)s] - %(message)s'
)

# Interval presets
INTERVAL_PRESETS = {
    "15 Minutes": 15,
    "30 Minutes": 30,
    "1 Hour": 60,
    "4 Hours": 240,
    "6 Hours": 360,
    "12 Hours": 720,
    "24 Hours": 1440,
    "Disabled": 0
}

def log_msg(msg, level="info"):
    print(msg)
    if level == "error":
        logging.error(msg)
    else:
        logging.info(msg)

class BingTrayApp:
    def __init__(self):
        self.icon = None
        self.root = None
        self.last_check = 0
        self._next_check_at = 0
        self.consecutive_failures = 0
        self.current_image_path = None
        self.running = True
        self.session = self._create_retry_session()
        self.last_success = 0
        self.last_error = "None"
        self.last_error_time = 0
        self.last_status_message = "Waiting for first check"
        self.status_vars = {}
        self.last_update_check = 0
        self.latest_release_url = None
        self.latest_release_version = None
        self.update_in_progress = False
        self._refresh_preview_ui = None
        self._render_gallery_page = None
        # Serializes wallpaper checks so startup, Check Now, F5 and the interval
        # loop never download the same image concurrently.
        self._check_lock = threading.Lock()
        self.cleanup_partial_downloads()

        self.config = self.load_config()
        interval_minutes = self.config.get("check_interval_minutes", 720)
        self.check_interval = interval_minutes * 60 if interval_minutes > 0 else 0
        
        log_msg(f"Initializing Bing Wallpaper App v{VERSION}")
        
        if not self.config.get("proxy_url"):
            self.detect_and_save_proxy()

    def _create_retry_session(self):
        session = requests.Session()
        retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
        session.mount('http://', HTTPAdapter(max_retries=retries))
        session.mount('https://', HTTPAdapter(max_retries=retries))
        return session

    def enable_high_dpi(self):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            log_msg("High-DPI mode enabled (Per-Monitor aware).")
            return
        except Exception:
            pass
        try:
            ctypes.windll.user32.SetProcessDPIAware()
            log_msg("High-DPI mode enabled (System aware).")
        except Exception:
            log_msg("High-DPI mode unavailable on this system.")

    def bind_window_shortcuts(self, window):
        window.bind('<Escape>', lambda e: window.withdraw() if window == self.root else window.destroy())
        window.bind('<Control-r>', lambda e: threading.Thread(target=self.check_and_update, args=(True,), daemon=True).start())
        window.bind('<Control-R>', lambda e: threading.Thread(target=self.check_and_update, args=(True,), daemon=True).start())
        window.bind('<Control-comma>', lambda e: self.show_settings_window())
        window.bind('<F5>', lambda e: threading.Thread(target=self.check_and_update, args=(True,), daemon=True).start())

    def get_verify_option(self):
        ca_bundle_path = str(self.config.get("ca_bundle_path", "")).strip()
        if not ca_bundle_path:
            return True
        if Path(ca_bundle_path).exists():
            return ca_bundle_path
        log_msg(f"Configured CA bundle path not found: {ca_bundle_path}", "error")
        return True

    def session_get(self, url, **kwargs):
        kwargs.setdefault("verify", self.get_verify_option())
        return self.session.get(url, **kwargs)

    def load_config(self):
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
        except Exception as e:
            log_msg(f"Error loading config: {e}", "error")
        return {}

    def save_config(self):
        # Write to a sibling temp file and atomically replace, so a crash or
        # power loss mid-write never leaves a truncated config.json behind.
        temp_path = CONFIG_FILE.with_name(CONFIG_FILE.name + ".tmp")
        try:
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(temp_path, CONFIG_FILE)
        except Exception as e:
            log_msg(f"Error saving config: {e}", "error")
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except Exception:
                pass

    def format_ts(self, timestamp):
        if not timestamp:
            return "Never"
        return datetime.datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")

    def is_newer_version(self, candidate):
        return core.is_newer_version(candidate, VERSION)

    def notify_user(self, title, message):
        try:
            if self.icon:
                self.icon.notify(message, title)
        except Exception:
            pass
        log_msg(f"{title}: {message}")

    def run_on_ui_thread(self, func):
        """Run func on the Tk main thread. Tk is not thread-safe, so worker
        threads must never touch widgets or StringVars directly."""
        if threading.current_thread() is threading.main_thread():
            func()
            return
        if self.root:
            try:
                self.root.after(0, func)
            except Exception:
                pass

    def update_status_panel(self):
        if not self.status_vars:
            return
        if threading.current_thread() is not threading.main_thread():
            self.run_on_ui_thread(self.update_status_panel)
            return
        interval_minutes = self.config.get("check_interval_minutes", 720)
        self.status_vars["last_check"].set(self.format_ts(self.last_check))
        self.status_vars["last_success"].set(self.format_ts(self.last_success))
        self.status_vars["last_error"].set(f"{self.format_ts(self.last_error_time)} - {self.last_error}" if self.last_error_time else self.last_error)
        self.status_vars["current_image"].set(self.current_image_path.name if self.current_image_path else "None")
        self.status_vars["interval"].set(f"{interval_minutes} minutes" if interval_minutes > 0 else "Disabled")
        self.status_vars["startup"].set("Enabled" if self.is_startup_enabled() else "Disabled")
        self.status_vars["status"].set(self.last_status_message)

    def record_error(self, message, notify=None):
        """Record a failure. notify=None toasts only when the message changed;
        True/False force or suppress the toast."""
        previous = self.last_error
        self.last_error = message
        self.last_error_time = time.time()
        self.last_status_message = "Last operation failed"
        self.update_status_panel()
        if notify is None:
            notify = message != previous
        if notify:
            self.notify_user("Bing Wallpaper", f"Operation failed: {message}")
        else:
            log_msg(f"Operation failed (notification suppressed): {message}")

    def get_int_setting(self, key, default, minimum, maximum):
        return core.clamp_int(self.config.get(key, default), default, minimum, maximum)

    def get_market(self):
        return core.normalize_market(self.config.get("market", DEFAULT_MARKET))

    def bing_api_url(self, count=1):
        return core.bing_api_url(self.get_market(), count)

    def get_latest_release_info(self):
        try:
            headers = {
                "Accept": "application/vnd.github+json",
                "User-Agent": f"{APP_NAME}/{VERSION}"
            }
            resp = self.session_get(RELEASES_API, timeout=10, proxies=self.get_proxy_dict(), headers=headers)
            resp.raise_for_status()
            data = resp.json()
            tag_name = data.get("tag_name", "")
            html_url = data.get("html_url", "")
            # Accept common naming variants, e.g. InstallBingWallpaper.exe or
            # InstallBingWallpaper-1.3.5.exe. The SHA-256 comes from GitHub's
            # asset digest, or from the release notes as a fallback.
            asset = core.select_installer_asset(data.get("assets", []), data.get("body", ""))
            installer_url = html_url
            installer_name = ""
            installer_sha256 = None
            if asset:
                installer_url = asset["url"] or html_url
                installer_name = asset["name"]
                installer_sha256 = asset["sha256"]
                log_msg(f"Selected installer asset: {installer_name}"
                        + (f" (sha256 {installer_sha256[:12]}...)" if installer_sha256 else " (no checksum published)"))
            if not tag_name:
                return None
            return {
                "version": str(tag_name).lstrip("v"),
                "url": installer_url or html_url,
                "html_url": html_url,
                "installer_name": installer_name or "InstallBingWallpaper.exe",
                "sha256": installer_sha256,
            }
        except Exception as e:
            log_msg(f"Update check failed: {e}", "error")
            return None

    def download_update_installer(self, info):
        url = info.get("url", "")
        if not url or not url.lower().endswith(".exe"):
            log_msg(f"No direct installer URL in release metadata: {url}")
            return None

        installer_name = info.get("installer_name", "InstallBingWallpaper.exe")
        temp_dir = Path(tempfile.gettempdir()) / APP_NAME / "updates"
        temp_dir.mkdir(parents=True, exist_ok=True)
        target = temp_dir / installer_name
        temp_target = target.with_suffix(".download")
        log_msg(f"Downloading installer from {url} to {target}")

        headers = {
            "User-Agent": f"{APP_NAME}/{VERSION}"
        }
        resp = self.session_get(url, timeout=90, proxies=self.get_proxy_dict(), headers=headers, stream=True)
        resp.raise_for_status()

        try:
            with open(temp_target, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            expected_sha256 = info.get("sha256")
            if expected_sha256:
                actual_sha256 = core.sha256_of_file(temp_target)
                if actual_sha256 != expected_sha256:
                    raise RuntimeError(
                        "Installer checksum mismatch: the download does not match the hash published "
                        f"on GitHub (expected {expected_sha256[:12]}..., got {actual_sha256[:12]}...). "
                        "The file was discarded."
                    )
                log_msg("Installer SHA-256 verified against the release digest.")
            else:
                log_msg("No SHA-256 published for this installer; skipping checksum verification.")

            os.replace(temp_target, target)
        finally:
            try:
                if temp_target.exists():
                    temp_target.unlink()
            except Exception:
                pass
        log_msg(f"Installer downloaded successfully: {target}")
        return target

    def launch_installer_and_exit(self, installer_path):
        subprocess.Popen([str(installer_path)], cwd=str(installer_path.parent))
        self.notify_user("Bing Wallpaper", "Launching updater and closing app.")
        self.running = False
        if self.icon:
            self.icon.stop()
        if self.root:
            self.root.after(200, self.root.quit)

    def perform_update_install(self, info):
        if self.update_in_progress:
            return
        self.update_in_progress = True
        try:
            self.last_status_message = f"Downloading update v{info['version']}..."
            self.update_status_panel()

            installer_path = self.download_update_installer(info)
            if not installer_path:
                if self.root:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Update",
                        "Installer asset was not found in the release. Opening release page instead."
                    ))
                webbrowser.open(info.get("html_url") or info.get("url"))
                return

            self.last_status_message = f"Downloaded update installer: {installer_path.name}"
            self.update_status_panel()

            if self.root:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Update",
                    "The updater will now start. The app will close to complete the upgrade."
                ))
                self.root.after(300, lambda: self.launch_installer_and_exit(installer_path))
            else:
                self.launch_installer_and_exit(installer_path)
        except SSLError as e:
            self.record_error(f"Update download SSL verification failed: {e}")
            fallback_url = info.get("url") or info.get("html_url")
            if fallback_url:
                log_msg(f"Falling back to browser download due to SSL verification failure: {fallback_url}", "error")
                if self.root:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Update",
                        "The app could not verify your network SSL certificate while downloading the installer.\n\n"
                        "This commonly happens on corporate proxies with SSL inspection.\n"
                        "The release page/download URL will be opened in your browser as a fallback."
                    ))
                webbrowser.open(fallback_url)
                self.last_status_message = "Updater fallback: download opened in browser"
                self.update_status_panel()
        except Exception as e:
            self.record_error(f"Update install failed: {e}")
            if self.root:
                self.root.after(0, lambda: messagebox.showerror("Update", f"Failed to install update: {e}"))
        finally:
            self.update_in_progress = False

    def check_for_updates(self, manual=False):
        info = self.get_latest_release_info()
        self.last_update_check = time.time()

        if not info:
            if manual:
                if self.root:
                    self.root.after(0, lambda: messagebox.showwarning("Update Check", "Unable to fetch release information."))
                self.notify_user("Bing Wallpaper", "Unable to check for updates right now.")
            return

        self.latest_release_version = info["version"]
        self.latest_release_url = info["url"]

        if self.is_newer_version(info["version"]):
            self.last_status_message = f"Update available: v{info['version']}"
            self.update_status_panel()
            self.notify_user("Bing Wallpaper", f"Update available: v{info['version']}")
            if manual and self.root:
                def ask_install():
                    install_now = messagebox.askyesno(
                        "Update Available",
                        f"A new version (v{info['version']}) is available.\nDownload and install now?"
                    )
                    if install_now:
                        threading.Thread(target=self.perform_update_install, args=(info,), daemon=True).start()
                    else:
                        open_page = messagebox.askyesno(
                            "Update",
                            "Open the release page in your browser instead?"
                        )
                        if open_page:
                            webbrowser.open(info["html_url"])
                self.root.after(0, ask_install)
        elif manual:
            self.last_status_message = f"Up to date: v{VERSION}"
            self.update_status_panel()
            log_msg(f"No newer version found. Current={VERSION}, Latest={info['version']}")
            if self.root:
                def ask_repair_download():
                    download_anyway = messagebox.askyesno(
                        "Update Check",
                        f"You are up to date (v{VERSION}).\n\nDo you want to download and run the installer anyway for repair/reinstall?"
                    )
                    if download_anyway:
                        threading.Thread(target=self.perform_update_install, args=(info,), daemon=True).start()
                self.root.after(0, ask_repair_download)
        else:
            self.last_status_message = f"Up to date: v{VERSION}"
            self.update_status_panel()

    def is_startup_enabled(self):
        return STARTUP_SHORTCUT.exists()

    def set_startup_enabled(self, enabled):
        try:
            if enabled:
                target = Path(sys.executable) if getattr(sys, 'frozen', False) else Path(sys.executable)
                startup_dir = STARTUP_SHORTCUT.parent
                startup_dir.mkdir(parents=True, exist_ok=True)

                args = ""
                if not getattr(sys, 'frozen', False):
                    args = f'"{str(Path(__file__).resolve())}"'

                ps_script = f"""
                $ws = New-Object -ComObject WScript.Shell
                $s = $ws.CreateShortcut('{str(STARTUP_SHORTCUT)}')
                $s.TargetPath = '{str(target)}'
                $s.WorkingDirectory = '{str(DATA_DIR)}'
                $s.Arguments = '{args}'
                $s.Save()
                """
                subprocess.run(["powershell", "-Command", ps_script], check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
            else:
                if STARTUP_SHORTCUT.exists():
                    STARTUP_SHORTCUT.unlink()
        except Exception as e:
            self.record_error(f"Startup toggle failed: {e}")

    def open_path(self, path):
        try:
            os.startfile(path)
        except Exception as e:
            self.record_error(f"Unable to open path: {e}")

    def test_https_connectivity(self):
        endpoints = [
            ("GitHub Releases API", RELEASES_API),
            ("Bing Image API", self.bing_api_url(1)),
        ]
        results = []
        for label, url in endpoints:
            try:
                resp = self.session_get(url, timeout=10, proxies=self.get_proxy_dict())
                resp.raise_for_status()
                results.append(f"[OK] {label}: HTTP {resp.status_code}")
            except Exception as e:
                results.append(f"[FAIL] {label}: {e}")

        ok_count = sum(1 for row in results if row.startswith("[OK]"))
        full_msg = "\n".join(results)
        log_msg(f"HTTPS test results:\n{full_msg}")

        if ok_count == len(endpoints):
            self.last_status_message = "HTTPS test passed"
            self.update_status_panel()
            if self.root:
                self.root.after(0, lambda: messagebox.showinfo("HTTPS Test", full_msg))
        else:
            self.record_error("HTTPS test failed")
            if self.root:
                self.root.after(0, lambda: messagebox.showwarning("HTTPS Test", full_msg))

    def detect_proxy_values(self):
        pac_url = self.get_pac_url_from_registry()
        if not pac_url:
            return "", ""
        proxy = self.get_proxy_from_pac(pac_url)
        if not proxy:
            return "", ""
        host, port = self.parse_proxy_string(proxy)
        return host or "", port or ""

    def set_interval(self, minutes, label):
        try:
            log_msg(f"Setting interval to: {label} ({minutes} minutes)")
            self.config["check_interval_minutes"] = minutes
            self.save_config()
            
            self.check_interval = minutes * 60 if minutes > 0 else 0
            self._next_check_at = 0  # run a check right away
            
            if self.icon:
                self.icon.menu = self.create_menu()
        except Exception as e:
            log_msg(f"Error setting interval: {e}", "error")

    def show_custom_interval_dialog(self):
        """Show dialog to input custom interval in minutes"""
        if not self.root:
            self.create_root()
        
        was_hidden = not self.root.winfo_viewable()
        if was_hidden:
            self.root.deiconify()
            self.root.update()
        
        result = simpledialog.askinteger(
            "Custom Interval",
            "Enter check interval in minutes:",
            parent=self.root,
            minvalue=1,
            maxvalue=10080
        )
        
        if was_hidden:
            self.root.withdraw()
        
        if result:
            self.set_interval(result, f"Custom ({result} min)")

    def show_settings_window(self):
        if not self.root:
            self.create_root()

        # Ensure the root window is visible so modal dialogs are not hidden.
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

        settings_win = tk.Toplevel(self.root)
        settings_win.title("Settings")
        settings_win.geometry("540x600")
        settings_win.minsize(540, 600)
        settings_win.resizable(True, True)
        settings_win.transient(self.root)
        settings_win.grab_set()
        settings_win.bind('<Escape>', lambda e: settings_win.destroy())

        frame = ttk.Frame(settings_win, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        settings_win.grid_rowconfigure(0, weight=1)
        settings_win.grid_columnconfigure(0, weight=1)

        frame.grid_columnconfigure(0, weight=0)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=0)
        frame.grid_rowconfigure(16, weight=1)

        ttk.Label(frame, text="Check Interval", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
        labels = list(INTERVAL_PRESETS.keys())
        interval_var = tk.StringVar(value="12 Hours")
        curr_interval = self.config.get("check_interval_minutes", 720)
        for label, minutes in INTERVAL_PRESETS.items():
            if minutes == curr_interval:
                interval_var.set(label)
                break

        interval_combo = ttk.Combobox(frame, textvariable=interval_var, values=labels, state="readonly", width=22)
        interval_combo.grid(row=1, column=0, sticky="ew")

        ttk.Label(frame, text="Custom minutes (optional)").grid(row=1, column=1, sticky="w", padx=(12, 0))
        custom_interval_var = tk.StringVar(value="" if curr_interval in INTERVAL_PRESETS.values() else str(curr_interval))
        ttk.Entry(frame, textvariable=custom_interval_var, width=10).grid(row=1, column=2, sticky="w")

        ttk.Separator(frame).grid(row=2, column=0, columnspan=3, sticky="ew", pady=12)

        ttk.Label(frame, text="Proxy", font=("Segoe UI", 10, "bold")).grid(row=3, column=0, sticky="w", pady=(0, 6))
        ttk.Label(frame, text="Host").grid(row=4, column=0, sticky="w")
        proxy_host_var = tk.StringVar(value=self.config.get("proxy_url", ""))
        ttk.Entry(frame, textvariable=proxy_host_var, width=28).grid(row=4, column=1, columnspan=2, sticky="ew")
        ttk.Label(frame, text="Port").grid(row=5, column=0, sticky="w", pady=(6, 0))
        proxy_port_var = tk.StringVar(value=self.config.get("proxy_port", ""))
        ttk.Entry(frame, textvariable=proxy_port_var, width=10).grid(row=5, column=1, sticky="w", pady=(6, 0))

        def auto_detect_proxy():
            host, port = self.detect_proxy_values()
            if host:
                proxy_host_var.set(host)
                proxy_port_var.set(port)
            else:
                messagebox.showinfo("Proxy Detection", "No proxy detected from system settings.")

        ttk.Button(frame, text="Auto-Detect Proxy", command=auto_detect_proxy).grid(row=5, column=2, sticky="w", pady=(6, 0))

        ttk.Label(frame, text="CA Bundle (PEM, optional)").grid(row=6, column=0, sticky="w", pady=(10, 0))
        ca_bundle_var = tk.StringVar(value=self.config.get("ca_bundle_path", ""))
        ttk.Entry(frame, textvariable=ca_bundle_var, width=50).grid(row=6, column=1, columnspan=2, sticky="ew", pady=(10, 0))

        def browse_ca_bundle():
            selected = filedialog.askopenfilename(
                parent=settings_win,
                title="Select CA Bundle (PEM)",
                filetypes=[("PEM files", "*.pem"), ("All files", "*.*")]
            )
            if selected:
                ca_bundle_var.set(selected)

        ttk.Button(frame, text="Browse...", command=browse_ca_bundle).grid(row=7, column=2, sticky="w", pady=(6, 0))

        ttk.Separator(frame).grid(row=8, column=0, columnspan=3, sticky="ew", pady=12)

        ttk.Label(frame, text="Wallpaper", font=("Segoe UI", 10, "bold")).grid(row=9, column=0, sticky="w", pady=(0, 6))

        ttk.Label(frame, text="Bing market").grid(row=10, column=0, sticky="w")
        market_var = tk.StringVar(value=self.get_market())
        market_combo = ttk.Combobox(frame, textvariable=market_var, values=MARKET_PRESETS, width=12)
        market_combo.grid(row=10, column=1, sticky="w")
        ttk.Label(frame, text="(region of the daily image, e.g. en-US)").grid(row=10, column=2, sticky="w", padx=(6, 0))

        ttk.Label(frame, text="Keep images for (days)").grid(row=11, column=0, sticky="w", pady=(6, 0))
        keep_days_var = tk.StringVar(value=str(self.get_int_setting("keep_images_days", 0, 0, 3650)))
        ttk.Entry(frame, textvariable=keep_days_var, width=8).grid(row=11, column=1, sticky="w", pady=(6, 0))
        ttk.Label(frame, text="(0 = keep forever)").grid(row=11, column=2, sticky="w", padx=(6, 0), pady=(6, 0))

        ttk.Label(frame, text="Backfill missed days").grid(row=12, column=0, sticky="w", pady=(6, 0))
        backfill_var = tk.StringVar(value=str(self.get_int_setting("backfill_days", MAX_BACKFILL_DAYS, 0, MAX_BACKFILL_DAYS)))
        ttk.Entry(frame, textvariable=backfill_var, width=8).grid(row=12, column=1, sticky="w", pady=(6, 0))
        ttk.Label(frame, text=f"(0-{MAX_BACKFILL_DAYS}, downloads images you missed)").grid(row=12, column=2, sticky="w", padx=(6, 0), pady=(6, 0))

        lock_screen_var = tk.BooleanVar(value=bool(self.config.get("set_lock_screen", False)))
        ttk.Checkbutton(frame, text="Also set the Windows lock screen image (turns off Windows Spotlight)", variable=lock_screen_var).grid(row=13, column=0, columnspan=3, sticky="w", pady=(8, 0))

        ttk.Separator(frame).grid(row=14, column=0, columnspan=3, sticky="ew", pady=12)

        startup_var = tk.BooleanVar(value=self.is_startup_enabled())
        ttk.Checkbutton(frame, text="Run at Startup", variable=startup_var).grid(row=15, column=0, columnspan=3, sticky="w")

        actions = ttk.Frame(frame)
        actions.grid(row=16, column=0, columnspan=3, sticky="ew", pady=(14, 8))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)
        actions.grid_columnconfigure(2, weight=1)
        actions.grid_columnconfigure(3, weight=1)
        ttk.Button(actions, text="Open Wallpaper Folder", command=lambda: self.open_path(IMAGE_DIR)).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(actions, text="Open Logs", command=lambda: self.open_path(LOG_DIR)).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(
            actions,
            text="Test HTTPS",
            command=lambda: threading.Thread(target=self.test_https_connectivity, daemon=True).start()
        ).grid(row=0, column=2, sticky="ew", padx=(0, 6))
        ttk.Button(
            actions,
            text="Check for Updates",
            command=lambda: threading.Thread(target=self.check_for_updates, args=(True,), daemon=True).start()
        ).grid(row=0, column=3, sticky="ew")

        def save_settings():
            try:
                custom_raw = custom_interval_var.get().strip()
                if custom_raw:
                    custom_minutes = int(custom_raw)
                    if custom_minutes <= 0:
                        raise ValueError("Custom interval must be positive")
                    self.set_interval(custom_minutes, f"Custom ({custom_minutes} min)")
                else:
                    selected = interval_var.get()
                    minutes = INTERVAL_PRESETS[selected]
                    self.set_interval(minutes, selected)

                self.config["proxy_url"] = proxy_host_var.get().strip()
                self.config["proxy_port"] = proxy_port_var.get().strip()
                self.config["ca_bundle_path"] = ca_bundle_var.get().strip()

                market = market_var.get().strip()
                if market and core.normalize_market(market) != market:
                    raise ValueError("Market must look like en-US")
                self.config["market"] = market or DEFAULT_MARKET

                keep_days = int(keep_days_var.get().strip() or "0")
                if keep_days < 0:
                    raise ValueError("Keep images days cannot be negative")
                self.config["keep_images_days"] = keep_days

                backfill_days = int(backfill_var.get().strip() or "0")
                if not 0 <= backfill_days <= MAX_BACKFILL_DAYS:
                    raise ValueError(f"Backfill days must be between 0 and {MAX_BACKFILL_DAYS}")
                self.config["backfill_days"] = backfill_days

                lock_screen_enabled = bool(lock_screen_var.get())
                lock_screen_was_enabled = bool(self.config.get("set_lock_screen", False))
                self.config["set_lock_screen"] = lock_screen_enabled
                self.save_config()

                if lock_screen_enabled and not lock_screen_was_enabled:
                    policy = self.lock_screen_policy_source()
                    if policy:
                        messagebox.showwarning(
                            "Lock Screen",
                            "Your organization enforces the lock screen image on this PC\n"
                            f"({policy}).\n\nThe option is saved, but Windows will keep the enforced image."
                        )
                    else:
                        image = self.current_image_path or self.latest_local_image()
                        if image:
                            self.apply_lock_screen(image)
                self.set_startup_enabled(startup_var.get())

                self.last_status_message = "Settings saved"
                self.update_status_panel()
                settings_win.destroy()
            except Exception as e:
                messagebox.showerror("Settings", f"Unable to save settings: {e}")

        btn_row = ttk.Frame(frame)
        btn_row.grid(row=17, column=0, columnspan=3, sticky="e")
        ttk.Button(btn_row, text="Cancel", command=settings_win.destroy).pack(side=tk.RIGHT)
        ttk.Button(btn_row, text="Save", command=save_settings).pack(side=tk.RIGHT, padx=(0, 8))
        settings_win.bind('<Return>', lambda e: save_settings())

    def get_proxy_dict(self):
        return core.build_proxy_dict(self.config.get("proxy_url", ""), self.config.get("proxy_port", ""))

    def detect_and_save_proxy(self):
        try:
            pac_url = self.get_pac_url_from_registry()
            if pac_url:
                proxy = self.get_proxy_from_pac(pac_url)
                if proxy:
                    host, port = self.parse_proxy_string(proxy)
                    if host:
                        self.config['proxy_url'] = host
                        self.config['proxy_port'] = port
                        self.save_config()
        except Exception: pass

    def get_pac_url_from_registry(self):
        locations = [
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"),
        ]
        for hive, subkey in locations:
            try:
                with winreg.OpenKey(hive, subkey) as key:
                    pac_url, _ = winreg.QueryValueEx(key, "AutoConfigURL")
                    if pac_url: return pac_url
            except Exception: continue
        return None

    def get_proxy_from_pac(self, pac_url):
        try:
            resp = self.session_get(pac_url, timeout=5)
            if resp.status_code == 200:
                return core.extract_pac_proxy(resp.text)
        except Exception: pass
        return None

    def parse_proxy_string(self, proxy_str):
        return core.parse_proxy_string(proxy_str)

    def get_bing_images(self, count=1):
        """Return the latest `count` Bing images, newest first, as dicts with
        'date' (YYYYMMDD), 'urls' (candidate download URLs, best first) and
        'title'. Returns None on failure."""
        try:
            resp = self.session_get(self.bing_api_url(count), timeout=10, proxies=self.get_proxy_dict())
            resp.raise_for_status()
            return core.parse_bing_images(resp.json()) or None
        except Exception as e:
            log_msg(f"API Fetch Error: {e}", "error")
            return None

    def cleanup_partial_downloads(self):
        """Remove leftover temp files from interrupted or failed downloads."""
        for pattern in ("bing_*.tmp", "bing_*.part"):
            for leftover in IMAGE_DIR.glob(pattern):
                try:
                    leftover.unlink()
                    log_msg(f"Removed stale partial download: {leftover.name}")
                except Exception as e:
                    log_msg(f"Could not remove stale partial download {leftover.name}: {e}", "error")

    def download_image(self, urls, date_str):
        """Download the image for date_str, trying each URL in `urls` in order
        (a str is accepted too). A 404, wrong content type or corrupt file moves
        on to the next candidate; network errors abort."""
        if isinstance(urls, str):
            urls = [urls]
        filename = core.image_filename(date_str)
        file_path = IMAGE_DIR / filename

        if file_path.exists() and file_path.stat().st_size > 0:
            return file_path

        # Unique per process/thread so two downloads can never collide on the
        # same temp file (this was the source of WinError 32 sharing violations).
        temp_path = IMAGE_DIR / f"{filename}.{os.getpid()}.{threading.get_ident()}.part"
        try:
            for index, url in enumerate(urls):
                resp = self.session_get(url, timeout=30, proxies=self.get_proxy_dict(), stream=True)
                if resp.status_code == 404 and index < len(urls) - 1:
                    log_msg(f"Not available at this resolution, trying next: {url}")
                    continue
                resp.raise_for_status()

                content_type = resp.headers.get('Content-Type', '')
                if 'image' not in content_type:
                    log_msg(f"Download Error: unexpected Content-Type '{content_type}' for {url}", "error")
                    continue

                with open(temp_path, 'wb') as f:
                    for chunk in resp.iter_content(chunk_size=8192):
                        f.write(chunk)

                try:
                    with Image.open(temp_path) as img:
                        img.verify()
                except Exception as e:
                    log_msg(f"Downloaded image failed verification ({url}): {e}", "error")
                    continue

                os.replace(temp_path, file_path)
                return file_path

            log_msg(f"Download Error: no usable image found for {date_str}", "error")
            return None

        except Exception as e:
            # If another check finished the same image meanwhile, use it.
            if file_path.exists() and file_path.stat().st_size > 0:
                return file_path
            log_msg(f"Download Error: {e}", "error")
            return None
        finally:
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except Exception:
                pass

    def set_wallpaper(self, image_path):
        if not image_path or not image_path.exists():
            return
        try:
            log_msg(f"Setting wallpaper: {image_path.name}")
            ctypes.windll.user32.SystemParametersInfoW(
                SPI_SETDESKWALLPAPER,
                0,
                str(image_path),
                SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE
            )
            self.current_image_path = image_path
            self.update_tray_icon(image_path)
            if self.config.get("set_lock_screen", False):
                self.apply_lock_screen(image_path)
        except Exception as e:
            log_msg(f"Wallpaper Set Error: {e}", "error")

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

    def lock_screen_policy_source(self):
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
        copy_dir = DATA_DIR / "lockscreen"
        copy_dir.mkdir(parents=True, exist_ok=True)
        copy_path = copy_dir / core.lock_screen_copy_name(image_path, time.time())
        shutil.copyfile(image_path, copy_path)
        for stale in core.select_stale_lock_screen_copies(copy_dir.glob("lockscreen_*"), copy_path):
            try:
                stale.unlink()
            except OSError:
                pass
        return copy_path

    def apply_lock_screen(self, image_path):
        """Set the lock screen image. Never blocks the Tk thread."""
        if threading.current_thread() is threading.main_thread():
            threading.Thread(target=self._set_lock_screen_blocking, args=(image_path,), daemon=True).start()
        else:
            self._set_lock_screen_blocking(image_path)

    def _set_lock_screen_blocking(self, image_path):
        policy = self.lock_screen_policy_source()
        if policy:
            self.record_error(f"Lock screen is managed by your organization ({policy}); the app cannot change it")
            return False
        try:
            self.switch_lock_screen_to_picture_mode()
            copy_path = self.prepare_lock_screen_copy(Path(image_path))
            script = self.LOCK_SCREEN_SCRIPT.replace("__IMAGE_PATH__", str(copy_path).replace("'", "''"))
            encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
            result = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
                capture_output=True, text=True, timeout=90,
                stdin=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            if result.returncode == 0:
                method = (result.stdout or "").strip().splitlines()
                method = method[-1] if method else "WinRT"
                log_msg(f"Lock screen updated: {Path(image_path).name} via {method}")
                return True
            detail = (result.stderr or result.stdout or "").strip().splitlines()
            detail = detail[0] if detail else f"exit code {result.returncode}"
            self.record_error(f"Lock screen update failed: {detail}")
        except Exception as e:
            self.record_error(f"Lock screen update failed: {e}")
        return False

    def latest_local_image(self):
        """Most recent downloaded wallpaper on disk, or None."""
        images = core.sort_images_newest_first(IMAGE_DIR.glob("bing_*.jpg"), os.path.getmtime)
        return images[0] if images else None

    def cleanup_old_images(self):
        """Delete wallpapers older than keep_images_days (0 = keep forever).
        The image currently in use is never deleted."""
        keep_days = self.get_int_setting("keep_images_days", 0, 0, 3650)
        if keep_days <= 0:
            return 0
        removed = 0
        for img in core.select_old_images(IMAGE_DIR.glob("bing_*.jpg"), keep_days, self.current_image_path):
            try:
                img.unlink()
                removed += 1
            except Exception as e:
                log_msg(f"Could not delete old wallpaper {img.name}: {e}", "error")
        if removed:
            log_msg(f"Retention: removed {removed} wallpaper(s) older than {keep_days} days")
        return removed

    def backfill_images(self, entries):
        """Download any of the given (older) Bing entries that are missing
        locally. Returns the number of images downloaded."""
        keep_days = self.get_int_setting("keep_images_days", 0, 0, 3650)
        downloaded = 0
        for entry in entries:
            if not self.running:
                break
            if not core.within_retention(entry["date"], keep_days):
                continue
            file_path = IMAGE_DIR / core.image_filename(entry["date"])
            if file_path.exists() and file_path.stat().st_size > 0:
                continue
            if self.download_image(entry["urls"], entry["date"]):
                downloaded += 1
                log_msg(f"Backfilled missed wallpaper: {file_path.name}")
        return downloaded

    def update_tray_icon(self, image_path):
        if self.icon:
            try:
                with Image.open(image_path) as img:
                    thumb = img.copy()
                    thumb.thumbnail((64, 64))
                    self.icon.icon = thumb
            except Exception: pass

    def check_and_update(self, force=False):
        if not self._check_lock.acquire(blocking=False):
            log_msg("Wallpaper check already in progress; skipping duplicate request.")
            return
        try:
            self._check_and_update_locked(force)
        finally:
            self._check_lock.release()

    def _refresh_preview_if_visible(self):
        if self.root and self.root.winfo_viewable():
            self.setup_ui(self.root)

    def _schedule_next_check(self, seconds):
        self._next_check_at = time.time() + max(0, seconds)

    def _on_check_failed(self, message):
        self.consecutive_failures += 1
        if self.check_interval > 0:
            delay = core.retry_delay_seconds(self.consecutive_failures, self.check_interval)
            self._schedule_next_check(delay)
            log_msg(f"Wallpaper check failed ({self.consecutive_failures} in a row); retrying in {delay // 60} min")
        # Stay quiet for the first few failures (network often is not up yet
        # right after login); toast once when the threshold is reached, and
        # again only if the error changes.
        notify = core.should_notify_failure(self.consecutive_failures, message, self.last_error)
        self.record_error(message, notify=notify)

    def _check_and_update_locked(self, force=False):
        self.last_check = time.time()
        try:
            backfill_days = self.get_int_setting("backfill_days", MAX_BACKFILL_DAYS, 0, MAX_BACKFILL_DAYS)
            images = self.get_bing_images(1 + backfill_days)
            if not images:
                self._on_check_failed("No image info returned from Bing API")
                return

            today = images[0]
            path = self.download_image(today["urls"], today["date"])
            if not path:
                self._on_check_failed("Image download failed")
                return

            if force or self.current_image_path != path:
                self.set_wallpaper(path)
                self.last_status_message = f"Wallpaper updated: {path.name}"
                self.notify_user("Bing Wallpaper", f"Wallpaper updated: {path.name}")
            elif self.icon and self.icon.icon is None:
                self.update_tray_icon(path)
                self.last_status_message = f"Wallpaper available: {path.name}"

            self.last_success = time.time()
            self.last_error = "None"
            self.last_error_time = 0
            self.consecutive_failures = 0
            self._schedule_next_check(self.check_interval)

            self.run_on_ui_thread(self._refresh_preview_if_visible)
            self.update_status_panel()

            # Housekeeping after today's wallpaper is in place.
            changed = self.backfill_images(images[1:])
            changed += self.cleanup_old_images()
            if changed:
                self.run_on_ui_thread(self._refresh_preview_if_visible)

        except Exception as e:
            log_msg(f"Update Loop Error: {e}", "error")
            self._on_check_failed(str(e))

    def background_loop(self):
        # Initial forced check at startup. Running it here instead of in a
        # separate thread avoids racing the interval check below.
        self.check_and_update(force=True)
        while self.running:
            try:
                if self.check_interval > 0 and time.time() >= self._next_check_at:
                    self.check_and_update(force=False)

                # Auto-check updates once every 24 hours.
                if (time.time() - self.last_update_check) > 86400:
                    self.check_for_updates(manual=False)
                time.sleep(5)
            except Exception:
                time.sleep(60)

    # --- MENU WITH CUSTOM OPTION RESTORED ---
    def create_menu(self):
        curr = self.config.get("check_interval_minutes", 720)
        
        def make_setter(m, l):
            return lambda i, it: self.set_interval(m, l)
            
        sub_items = []
        for label, mins in INTERVAL_PRESETS.items():
            state = True if curr == mins else False
            sub_items.append(item(label, make_setter(mins, label), checked=lambda i, s=state: s))
        
        # Add Custom Option
        sub_items.append(item('─────────', None, enabled=False))
        is_custom = curr not in INTERVAL_PRESETS.values()
        custom_label = f"Custom ({curr} min)" if is_custom else "Custom..."
        
        def custom_setter(icon, item):
            if self.root:
                self.root.after(0, self.show_custom_interval_dialog)

        sub_items.append(item(custom_label, custom_setter, checked=lambda i, s=is_custom: s))

        return pystray.Menu(
            item('Preview / Gallery', self.on_open_preview, default=True),
            item('Settings', self.on_open_settings),
            item('Check Now', lambda i, it: threading.Thread(target=self.check_and_update, args=(True,), daemon=True).start()),
            item('Check for Updates', lambda i, it: threading.Thread(target=self.check_for_updates, args=(True,), daemon=True).start()),
            pystray.Menu.SEPARATOR,
            item('Interval', pystray.Menu(*sub_items)),
            pystray.Menu.SEPARATOR,
            item('Exit', self.on_exit)
        )

    def on_open_preview(self, icon, item):
        if not self.root:
            self.create_root()
        self.root.after(0, self.show_preview_window)

    def on_open_settings(self, icon, item):
        if not self.root:
            self.create_root()
        self.root.after(0, self.show_settings_window)

    def on_exit(self, icon, item):
        self.running = False
        icon.stop()
        if self.root: self.root.quit()

    def show_preview_window(self):
        if not self.root: self.create_root()
        self.root.deiconify()
        self.root.lift()
        self.setup_ui(self.root)

    def create_root(self):
        self.enable_high_dpi()
        self.root = tk.Tk()
        self.root.title(f"Bing Wallpaper v{VERSION}")
        self.root.geometry("840x760")
        self.root.minsize(620, 560)
        self.root.protocol("WM_DELETE_WINDOW", self.root.withdraw)
        self.bind_window_shortcuts(self.root)
        self.setup_ui(self.root)

    def on_thumbnail_click(self, img_path):
        self.set_wallpaper(img_path)
        self.last_status_message = f"Wallpaper updated: {img_path.name}"
        if callable(self._refresh_preview_ui):
            self._refresh_preview_ui(force_reload=True)
        self.update_status_panel()

    # --- PREVIEW UI WITH RESPONSIVE LAYOUT ---
    def setup_ui(self, win):
        self._refresh_preview_ui = None
        self._render_gallery_page = None
        for w in win.winfo_children():
            w.destroy()

        main_frame = ttk.Frame(win, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")
        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)

        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(0, weight=0)
        main_frame.grid_rowconfigure(1, weight=0)
        main_frame.grid_rowconfigure(2, weight=5)
        main_frame.grid_rowconfigure(3, weight=2)

        actions_frame = ttk.Frame(main_frame)
        actions_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        actions_frame.grid_columnconfigure(0, weight=1)
        actions_frame.grid_columnconfigure(1, weight=1)
        actions_frame.grid_columnconfigure(2, weight=1)
        ttk.Button(actions_frame, text="Check Now", command=lambda: threading.Thread(target=self.check_and_update, args=(True,), daemon=True).start()).grid(row=0, column=0, sticky="ew")
        ttk.Button(actions_frame, text="Settings", command=self.show_settings_window).grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Button(actions_frame, text="Open Logs", command=lambda: self.open_path(LOG_DIR)).grid(row=0, column=2, sticky="ew", padx=(6, 0))

        status_frame = ttk.LabelFrame(main_frame, text="Status", padding=8)
        status_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.status_vars = {
            "status": tk.StringVar(),
            "last_check": tk.StringVar(),
            "last_success": tk.StringVar(),
            "last_error": tk.StringVar(),
            "current_image": tk.StringVar(),
            "interval": tk.StringVar(),
            "startup": tk.StringVar(),
        }

        status_rows = [
            ("State", "status"),
            ("Last Check", "last_check"),
            ("Last Success", "last_success"),
            ("Last Error", "last_error"),
            ("Current Image", "current_image"),
            ("Interval", "interval"),
            ("Startup", "startup"),
        ]
        value_labels = []
        for idx, (label, key) in enumerate(status_rows):
            ttk.Label(status_frame, text=f"{label}:", width=14).grid(row=idx, column=0, sticky="w", pady=1)
            value_label = ttk.Label(status_frame, textvariable=self.status_vars[key], anchor="w")
            value_label.grid(row=idx, column=1, sticky="ew", pady=1)
            value_labels.append(value_label)
        status_frame.grid_columnconfigure(1, weight=1)
        self.update_status_panel()

        preview_frame = ttk.Frame(main_frame)
        preview_frame.grid(row=2, column=0, sticky="nsew")
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        preview_label = ttk.Label(preview_frame, anchor="center")
        preview_label.grid(row=0, column=0, sticky="nsew")
        preview_name_label = ttk.Label(preview_frame, font=("Segoe UI", 10, "bold"), anchor="center")
        preview_name_label.grid(row=1, column=0, sticky="ew", pady=(8, 0))

        latest_frame = ttk.LabelFrame(main_frame, text="Latest 5", padding=6)
        latest_frame.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        for col in range(5):
            latest_frame.grid_columnconfigure(col, weight=1, uniform="latest")
        latest_frame.grid_rowconfigure(0, weight=1)
        latest_frame.grid_columnconfigure(5, weight=0)
        latest_frame.grid_rowconfigure(1, weight=0)

        images = core.sort_images_newest_first(IMAGE_DIR.glob("bing_*.jpg"), os.path.getmtime)
        gallery_offset = {"value": 0}
        thumb_slots = []

        for idx in range(5):
            slot = ttk.Frame(latest_frame, padding=2)
            slot.grid(row=0, column=idx, sticky="nsew", padx=4, pady=2)
            slot.grid_rowconfigure(0, weight=1)
            slot.grid_columnconfigure(0, weight=1)

            img_label = ttk.Label(slot, anchor="center", cursor="hand2")
            img_label.grid(row=0, column=0, sticky="nsew")
            date_label = ttk.Label(slot, anchor="center", font=("Consolas", 8))
            date_label.grid(row=1, column=0, sticky="ew", pady=(4, 0))
            thumb_slots.append((img_label, date_label))

        gallery_scroll = ttk.Scrollbar(latest_frame, orient="horizontal")
        gallery_scroll.grid(row=1, column=0, columnspan=5, sticky="ew", padx=4, pady=(4, 0))
        gallery_status_label = ttk.Label(latest_frame, text="", anchor="e")
        gallery_status_label.grid(row=1, column=5, sticky="e", padx=(8, 2), pady=(4, 0))

        preview_cache = {
            "path": None,
            "image": None,
        }
        thumb_source_cache = {}

        layout_state = {
            "thumb_min_height": 60,
            "thumb_container_min": 120,
        }

        render_state = {
            "preview_size": None,
            "thumb_size": None,
            "preview_after": None,
            "thumb_after": None,
        }

        def update_gallery_scrollbar():
            total = len(images)
            max_offset = max(total - 5, 0)

            if total <= 5:
                gallery_status_label.configure(text=f"{total}/{total}" if total else "0/0")
                gallery_scroll.set(0.0, 1.0)
                return

            start = gallery_offset["value"] / total
            end = min((gallery_offset["value"] + 5) / total, 1.0)
            gallery_scroll.set(start, end)
            gallery_status_label.configure(
                text=f"{gallery_offset['value'] + 1}-{min(gallery_offset['value'] + 5, total)}/{total}"
            )

        def set_gallery_offset(offset):
            max_offset = max(len(images) - 5, 0)
            clamped = min(max(offset, 0), max_offset)
            if clamped == gallery_offset["value"] and callable(self._render_gallery_page):
                update_gallery_scrollbar()
                return
            gallery_offset["value"] = clamped
            if callable(self._render_gallery_page):
                self._render_gallery_page()

        def on_gallery_scroll(*args):
            if not args:
                return
            total = len(images)
            max_offset = max(total - 5, 0)
            if total <= 5:
                update_gallery_scrollbar()
                return

            command = args[0]
            if command == "moveto" and len(args) > 1:
                fraction = float(args[1])
                set_gallery_offset(int(round(fraction * max_offset)))
            elif command == "scroll" and len(args) > 2:
                delta = int(args[1])
                amount = 1
                if args[2] == "pages":
                    amount = 5
                set_gallery_offset(gallery_offset["value"] + delta * amount)

        gallery_scroll.configure(command=on_gallery_scroll)

        def get_preview_source(path):
            if not path or not path.exists():
                return None
            if preview_cache["path"] == path and preview_cache["image"] is not None:
                return preview_cache["image"]
            try:
                with Image.open(path) as img:
                    source = img.copy()
                preview_cache["path"] = path
                preview_cache["image"] = source
                return source
            except (UnidentifiedImageError, OSError):
                preview_cache["path"] = None
                preview_cache["image"] = None
                return None

        def get_thumb_source(path):
            if path in thumb_source_cache:
                return thumb_source_cache[path]
            try:
                with Image.open(path) as img:
                    source = img.copy()
                thumb_source_cache[path] = source
                return source
            except (UnidentifiedImageError, OSError):
                return None

        def refresh_current_preview(force_reload=False):
            path = self.current_image_path
            if force_reload:
                preview_cache["path"] = None
                preview_cache["image"] = None
                render_state["preview_size"] = None
            if not path or not path.exists():
                preview_label.configure(image="", text="No wallpaper set.")
                preview_label.image = None
                preview_name_label.configure(text="")
                return

            preview_name_label.configure(text=path.name)
            source = get_preview_source(path)
            if source is None:
                preview_label.configure(image="", text="Error displaying image")
                preview_label.image = None
                preview_name_label.configure(text="")
                return

            redraw_preview(force=True)

        def render_gallery_page():
            visible = images[gallery_offset["value"]: gallery_offset["value"] + 5]

            for idx in range(5):
                img_label, date_label = thumb_slots[idx]
                if idx >= len(visible):
                    img_label.configure(image="", text="No image")
                    img_label.image = None
                    img_label.unbind("<Button-1>")
                    date_label.configure(text="")
                    continue

                img_path = visible[idx]
                img_label.configure(text="")
                img_label.bind("<Button-1>", lambda e, p=img_path: self.on_thumbnail_click(p))

                date_str = img_path.stem.split('_')[-1]
                try:
                    parsed_date = datetime.datetime.strptime(date_str, "%Y%m%d")
                    display_name = parsed_date.strftime("%b %d, %Y")
                except ValueError:
                    display_name = img_path.stem
                date_label.configure(text=display_name)

            render_state["thumb_size"] = None
            redraw_thumbnails(force=True)
            update_gallery_scrollbar()

        self._refresh_preview_ui = refresh_current_preview
        self._render_gallery_page = render_gallery_page

        def redraw_preview(_event=None, force=False):
            source = get_preview_source(self.current_image_path)
            if source is None:
                return
            avail_w = max(preview_frame.winfo_width() - 16, 220)
            avail_h = max(preview_frame.winfo_height() - 50, 180)
            size_key = (avail_w, avail_h)
            if not force and render_state["preview_size"] == size_key:
                return
            render_state["preview_size"] = size_key

            resized = source.copy()
            resized.thumbnail((avail_w, avail_h))
            tk_img = ImageTk.PhotoImage(resized)
            preview_label.configure(image=tk_img, text="")
            preview_label.image = tk_img

        def redraw_thumbnails(_event=None, force=False):
            container_w = max(latest_frame.winfo_width() - 30, 500)
            container_h = max(latest_frame.winfo_height() - 42, layout_state["thumb_container_min"])
            per_width = max(container_w // 5 - 12, 80)
            per_height = max(container_h - 24, layout_state["thumb_min_height"])
            size_key = (per_width, per_height, gallery_offset["value"])
            if not force and render_state["thumb_size"] == size_key:
                return
            render_state["thumb_size"] = size_key

            visible = images[gallery_offset["value"]: gallery_offset["value"] + 5]
            for idx, img_path in enumerate(visible):
                original = get_thumb_source(img_path)
                if original is None:
                    continue
                view = original.copy()
                view.thumbnail((per_width, per_height))
                tk_img = ImageTk.PhotoImage(view)
                img_label, date_label = thumb_slots[idx]
                img_label.configure(image=tk_img, text="")
                img_label.image = tk_img

        def schedule_preview_redraw(_event=None):
            if render_state["preview_after"]:
                win.after_cancel(render_state["preview_after"])
            render_state["preview_after"] = win.after(45, lambda: redraw_preview())

        def schedule_thumb_redraw(_event=None):
            if render_state["thumb_after"]:
                win.after_cancel(render_state["thumb_after"])
            render_state["thumb_after"] = win.after(60, lambda: redraw_thumbnails())

        def apply_adaptive_layout(_event=None):
            if _event is not None and _event.widget is not win:
                return
            h = max(win.winfo_height(), 560)
            if h < 680:
                preview_weight = 2
                latest_weight = 3
                name_font_size = 9
                thumb_font_size = 7
                layout_state["thumb_min_height"] = 74
                layout_state["thumb_container_min"] = 150
            elif h > 980:
                preview_weight = 6
                latest_weight = 2
                name_font_size = 11
                thumb_font_size = 9
                layout_state["thumb_min_height"] = 64
                layout_state["thumb_container_min"] = 130
            else:
                preview_weight = 5
                latest_weight = 2
                name_font_size = 10
                thumb_font_size = 8
                layout_state["thumb_min_height"] = 60
                layout_state["thumb_container_min"] = 120

            main_frame.grid_rowconfigure(2, weight=preview_weight)
            main_frame.grid_rowconfigure(3, weight=latest_weight)
            preview_name_label.configure(font=("Segoe UI", name_font_size, "bold"))

            for idx, _ in enumerate(status_rows):
                value_labels[idx].configure(wraplength=max(status_frame.winfo_width() - 200, 200))

            for _, date_label in thumb_slots:
                date_label.configure(font=("Consolas", thumb_font_size))

            schedule_preview_redraw()
            schedule_thumb_redraw()

        preview_frame.bind("<Configure>", schedule_preview_redraw)
        latest_frame.bind("<Configure>", schedule_thumb_redraw)
        win.bind("<Configure>", apply_adaptive_layout)
        refresh_current_preview(force_reload=True)
        render_gallery_page()
        win.after(20, schedule_preview_redraw)
        win.after(20, schedule_thumb_redraw)
        win.after(20, apply_adaptive_layout)

    def create_thumbnail(self, parent, img_path):
        try:
            pil_img = Image.open(img_path)
            pil_img.thumbnail((150, 100))
            tk_img = ImageTk.PhotoImage(pil_img)
            
            f = tk.Frame(parent, bd=2, relief="groove")
            f.pack(side=tk.LEFT, padx=5)
            
            lbl = tk.Label(f, image=tk_img, cursor="hand2")
            lbl.image = tk_img
            lbl.pack()
            
            lbl.bind("<Button-1>", lambda e, p=img_path: self.set_wallpaper(p))
            
            # Extract '20260408' from 'bing_20260408.jpg'
            date_str = img_path.stem.split('_')[-1] 
            try:
                # Convert 'YYYYMMDD' to 'Apr 08, 2026'
                parsed_date = datetime.datetime.strptime(date_str, "%Y%m%d")
                display_name = parsed_date.strftime("%b %d, %Y")
            except ValueError:
                display_name = img_path.stem

            tk.Label(f, text=display_name, font=("Consolas", 8)).pack()
            
        except Exception: pass

    def run(self):
        t = threading.Thread(target=self.background_loop, daemon=True)
        t.start()
        
        try:
            icon_img = Image.new('RGB', (64, 64), color=(0, 120, 215))
            self.icon = pystray.Icon(APP_NAME, icon_img, "Bing Wallpaper", self.create_menu())
            threading.Thread(target=self.icon.run, daemon=True).start()
            self.create_root()
            self.root.withdraw()
            self.root.mainloop()
        except KeyboardInterrupt:
            self.running = False

if __name__ == "__main__":
    app_mutex = check_single_instance()
    app = BingTrayApp()
    app.run()