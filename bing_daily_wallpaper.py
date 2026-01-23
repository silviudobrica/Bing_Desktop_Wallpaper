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
import shutil
from pathlib import Path
from logging.handlers import RotatingFileHandler
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from PIL import Image, ImageTk, UnidentifiedImageError
import pystray
from pystray import MenuItem as item
import tkinter as tk
from tkinter import ttk, simpledialog
import winreg
import re

# Import centralized version
try:
    from _version import __version__ as VERSION
except ImportError:
    VERSION = "1.3.2"

# Configuration
BING_API = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US"
APP_NAME = "BingWallpaper"

# Paths
DATA_DIR = Path(os.environ["LOCALAPPDATA"]) / "Programs" / APP_NAME
LOG_DIR = DATA_DIR / "logs"
CONFIG_FILE = DATA_DIR / "config.json"
IMAGE_DIR = Path(os.environ["USERPROFILE"]) / "Pictures" / "Bing"

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
        self.current_image_path = None
        self.running = True
        self.session = self._create_retry_session()
        
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

    def load_config(self):
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
        except Exception as e:
            log_msg(f"Error loading config: {e}", "error")
        return {}

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            log_msg(f"Error saving config: {e}", "error")

    def set_interval(self, minutes, label):
        try:
            log_msg(f"Setting interval to: {label} ({minutes} minutes)")
            self.config["check_interval_minutes"] = minutes
            self.save_config()
            
            self.check_interval = minutes * 60 if minutes > 0 else 0
            self.last_check = 0 
            
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

    def get_proxy_dict(self):
        url = self.config.get("proxy_url", "").strip()
        port = self.config.get("proxy_port", "").strip()
        if url and port:
            proxy = f"http://{url}:{port}"
            return {"http": proxy, "https": proxy}
        return None

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
            resp = self.session.get(pac_url, timeout=5)
            if resp.status_code == 200:
                match = re.search(r'PROXY\s+([a-zA-Z0-9.-]+:\d+)', resp.text)
                if match: return match.group(1)
        except Exception: pass
        return None

    def parse_proxy_string(self, proxy_str):
        if not proxy_str: return None, None
        parts = proxy_str.split(':')
        return (parts[0], parts[1]) if len(parts) > 1 else (parts[0], "80")

    def get_bing_image_info(self):
        try:
            resp = self.session.get(BING_API, timeout=10, proxies=self.get_proxy_dict())
            resp.raise_for_status()
            data = resp.json()
            if not data.get("images"): return None
            img_data = data["images"][0]
            return ("https://www.bing.com" + img_data["url"], img_data["startdate"])
        except Exception as e:
            log_msg(f"API Fetch Error: {e}", "error")
            return None

    def download_image(self, url, date_str):
        filename = f"bing_{date_str}.jpg"
        file_path = IMAGE_DIR / filename
        
        if file_path.exists() and file_path.stat().st_size > 0:
            return file_path
        
        try:
            resp = self.session.get(url, timeout=30, proxies=self.get_proxy_dict(), stream=True)
            resp.raise_for_status()
            
            if 'image' not in resp.headers.get('Content-Type', ''):
                return None

            temp_path = file_path.with_suffix(".tmp")
            with open(temp_path, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            try:
                with Image.open(temp_path) as img:
                    img.verify()
                shutil.move(temp_path, file_path)
                return file_path
            except Exception:
                if temp_path.exists(): os.remove(temp_path)
                return None

        except Exception as e:
            log_msg(f"Download Error: {e}", "error")
            return None

    def set_wallpaper(self, image_path):
        if not image_path or not image_path.exists():
            return
        try:
            log_msg(f"Setting wallpaper: {image_path.name}")
            ctypes.windll.user32.SystemParametersInfoW(20, 0, str(image_path), 3)
            self.current_image_path = image_path
            self.update_tray_icon(image_path)
        except Exception as e:
            log_msg(f"Wallpaper Set Error: {e}", "error")

    def update_tray_icon(self, image_path):
        if self.icon:
            try:
                with Image.open(image_path) as img:
                    thumb = img.copy()
                    thumb.thumbnail((64, 64))
                    self.icon.icon = thumb
            except Exception: pass

    def check_and_update(self, force=False):
        try:
            info = self.get_bing_image_info()
            if info:
                url, date_str = info
                path = self.download_image(url, date_str)
                if path:
                    if force or self.current_image_path != path:
                        self.set_wallpaper(path)
                    elif self.icon and self.icon.icon is None:
                        self.update_tray_icon(path)
            
            if self.root and self.root.winfo_viewable():
                self.root.after(0, lambda: self.setup_ui(self.root))
                
        except Exception as e:
            log_msg(f"Update Loop Error: {e}", "error")
        
        self.last_check = time.time()

    def background_loop(self):
        while self.running:
            try:
                if self.check_interval > 0:
                    elapsed = time.time() - self.last_check
                    if elapsed > self.check_interval:
                        self.check_and_update(force=False)
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
            item('Check Now', lambda i, it: threading.Thread(target=self.check_and_update, args=(True,)).start()),
            pystray.Menu.SEPARATOR,
            item('Interval', pystray.Menu(*sub_items)),
            pystray.Menu.SEPARATOR,
            item('Exit', self.on_exit)
        )

    def on_open_preview(self, icon, item):
        if self.root: self.root.after(0, self.show_preview_window)

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
        self.root = tk.Tk()
        self.root.title(f"Bing Wallpaper v{VERSION}")
        self.root.geometry("800x600")
        self.root.protocol("WM_DELETE_WINDOW", self.root.withdraw)
        self.setup_ui(self.root)

    # --- PREVIEW UI WITH THUMBNAILS RESTORED ---
    def setup_ui(self, win):
        for w in win.winfo_children(): w.destroy()
        
        # Main Container
        main_frame = tk.Frame(win)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 1. Main Preview Area
        preview_frame = tk.Frame(main_frame)
        preview_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        if self.current_image_path and self.current_image_path.exists():
            try:
                img = Image.open(self.current_image_path)
                img.thumbnail((780, 400))
                tk_img = ImageTk.PhotoImage(img)
                lbl = tk.Label(preview_frame, image=tk_img)
                lbl.image = tk_img 
                lbl.pack(pady=5)
                tk.Label(preview_frame, text=self.current_image_path.name, font=("Segoe UI", 10, "bold")).pack()
            except Exception:
                tk.Label(preview_frame, text="Error displaying image").pack()
        else:
            tk.Label(preview_frame, text="No wallpaper set.").pack(pady=50)

        # 2. Horizontal Scroll List (Restored)
        list_frame = tk.Frame(main_frame, height=150)
        list_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        canvas = tk.Canvas(list_frame, height=130)
        scrollbar = ttk.Scrollbar(list_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(xscrollcommand=scrollbar.set)

        # Populating images with Thumbnails
        images = sorted(IMAGE_DIR.glob("bing_*.jpg"), key=os.path.getmtime, reverse=True)
        for img_path in images[:15]: 
            self.create_thumbnail(scrollable_frame, img_path)

        canvas.pack(side="top", fill="both", expand=True)
        scrollbar.pack(side="bottom", fill="x")

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
            
            tk.Label(f, text=img_path.name[-12:], font=("Consolas", 8)).pack() 
            
        except Exception: pass

    def run(self):
        t = threading.Thread(target=self.background_loop, daemon=True)
        t.start()
        threading.Thread(target=self.check_and_update, args=(True,), daemon=True).start()
        
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
    app = BingTrayApp()
    app.run()