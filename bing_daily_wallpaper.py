# bing_daily_wallpaper.py
import os
import sys
import ctypes
import datetime
import time
import threading
import logging
import subprocess
import webbrowser
import tempfile
from pathlib import Path
from logging.handlers import RotatingFileHandler
from PIL import Image, ImageTk, UnidentifiedImageError
import pystray
from pystray import MenuItem as item
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog
from requests.exceptions import SSLError
import bing_core as core
from bing_core import (
    DEFAULT_MARKET, MARKET_PRESETS, MAX_BACKFILL_DAYS,
)
from bing_service import WallpaperService, INTERVAL_PRESETS, RELEASES_API, log_msg
from platform_windows import WindowsPlatform

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

# Import centralized version
try:
    from _version import __version__ as VERSION
except ImportError:
    VERSION = "1.4.4"

# Configuration
# Bing API, retry and retention constants live in bing_core.py.
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

class BingTrayApp(WallpaperService):
    """Windows tray front end. All download/backfill/retention/config logic
    lives in WallpaperService; this class adds Tk, pystray and the updater."""

    def __init__(self):
        self.icon = None
        self.root = None
        self.status_vars = {}
        self.update_in_progress = False
        self._refresh_preview_ui = None
        self._render_gallery_page = None
        platform = WindowsPlatform(
            DATA_DIR, STARTUP_SHORTCUT,
            script_path=None if getattr(sys, 'frozen', False) else Path(__file__),
        )
        super().__init__(platform, CONFIG_FILE, IMAGE_DIR, APP_NAME, VERSION)

    # ---- service hooks (may be called from worker threads)
    def on_state_changed(self):
        self.update_status_panel()

    def on_wallpaper_applied(self, image_path):
        self.update_tray_icon(image_path)

    def on_images_changed(self):
        self.run_on_ui_thread(self._refresh_preview_if_visible)

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
        return self.download_release_asset(url, installer_name, info.get("sha256"), temp_dir, timeout=90)

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
                detail = str(e)
                self.root.after(0, lambda: messagebox.showerror("Update", f"Failed to install update: {detail}"))
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


    def set_interval(self, minutes, label):
        try:
            log_msg(f"Setting interval to: {label} ({minutes} minutes)")
            WallpaperService.set_interval(self, minutes)
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
        settings_win.geometry("620x600")
        settings_win.minsize(620, 600)
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


    def update_tray_icon(self, image_path):
        if self.icon:
            try:
                with Image.open(image_path) as img:
                    thumb = img.copy()
                    thumb.thumbnail((64, 64))
                    self.icon.icon = thumb
            except Exception: pass


    def _refresh_preview_if_visible(self):
        if self.root and self.root.winfo_viewable():
            self.setup_ui(self.root)


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

        images = self.list_images_newest_first()
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