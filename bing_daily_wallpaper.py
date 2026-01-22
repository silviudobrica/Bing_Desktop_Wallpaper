import os
import sys
import ctypes
import requests
import datetime
import time
import threading
import logging
import json
from pathlib import Path
from PIL import Image, ImageTk
import pystray
from pystray import MenuItem as item
import tkinter as tk
from tkinter import ttk, simpledialog
import winreg
import urllib.request
import re

# Configuration
BING_API = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US"
SAVE_DIR = Path(os.environ["USERPROFILE"]) / "Pictures" / "Bing"
APP_DIR = Path(os.environ["LOCALAPPDATA"]) / "Programs" / "BingWallpaper"
CONFIG_FILE = APP_DIR / "config.json"
VERSION = "1.3.0"

# Interval presets in minutes
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

# Setup Logging
SAVE_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = SAVE_DIR / "bing_wallpaper.log"
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def log_msg(msg):
    print(msg)
    logging.info(msg)

class BingTrayApp:
    def __init__(self):
        self.icon = None
        self.root = None
        self.last_check = 0
        self.current_image_path = None
        self.running = True
        
        # Ensure directories exist
        SAVE_DIR.mkdir(parents=True, exist_ok=True)
        APP_DIR.mkdir(parents=True, exist_ok=True)
        
        # Load config and set check interval
        config = self.load_config()
        interval_minutes = config.get("check_interval_minutes", 720)  # Default 12 hours
        self.check_interval = interval_minutes * 60 if interval_minutes > 0 else 0
        
        log_msg(f"Initializing Bing Wallpaper App v{VERSION}")
        log_msg(f"Check interval: {interval_minutes} minutes")

        # Auto-detect proxy if not configured
        if not self.get_proxy_dict():
            self.detect_and_save_proxy()

    def load_config(self):
        """Load configuration from JSON file"""
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, 'r') as f:
                    return json.load(f)
        except Exception as e:
            log_msg(f"Error loading config: {e}")
        return {}

    def save_config(self, config):
        """Save configuration to JSON file"""
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
            log_msg(f"Config saved: {config}")
        except Exception as e:
            log_msg(f"Error saving config: {e}")

    def set_interval(self, minutes, label):
        """Set the check interval and save to config"""
        try:
            log_msg(f"Setting interval to: {label} ({minutes} minutes)")
            config = self.load_config()
            config["check_interval_minutes"] = minutes
            self.save_config(config)
            
            # Update runtime interval
            self.check_interval = minutes * 60 if minutes > 0 else 0
            self.last_check = 0  # Reset to trigger check soon
            
            # Rebuild menu to update checkmarks
            if self.icon:
                self.icon.menu = self.create_menu()
        except Exception as e:
            log_msg(f"Error setting interval: {e}")

    def show_custom_interval_dialog(self):
        """Show dialog to input custom interval in minutes"""
        if not self.root:
            self.create_root()
        
        # Temporarily show root to make dialog visible
        was_hidden = not self.root.winfo_viewable()
        if was_hidden:
            self.root.deiconify()
            self.root.update()  # Force window to appear
        
        result = simpledialog.askinteger(
            "Custom Interval",
            "Enter check interval in minutes:",
            parent=self.root,
            minvalue=1,
            maxvalue=10080  # 1 week max
        )
        
        # Hide root again if it was hidden
        if was_hidden:
            self.root.withdraw()
        
        if result:
            self.set_interval(result, f"Custom ({result} min)")

    def get_proxy_dict(self):
        """Get proxy configuration for requests library"""
        config = self.load_config()
        proxy_url = config.get("proxy_url", "").strip()
        proxy_port = config.get("proxy_port", "").strip()
        
        if proxy_url and proxy_port:
            proxy = f"http://{proxy_url}:{proxy_port}"
            log_msg(f"Using proxy: {proxy}")
            return {"http": proxy, "https": proxy}
        return None



    def detect_and_save_proxy(self):
        """Detect proxy via PAC and save to config if found"""
        log_msg("Attempting to detect proxy via PAC...")
        try:
            pac_url = self.get_pac_url_from_registry()
            if pac_url:
                log_msg(f"Found PAC URL: {pac_url}")
                raw_proxy = self.get_external_proxy_string(pac_url)
                if raw_proxy:
                    host, port = self.parse_proxy_string(raw_proxy)
                    if host:
                        log_msg(f"Detected Proxy: {host}:{port}")
                        config = self.load_config()
                        config['proxy_url'] = host
                        config['proxy_port'] = port
                        self.save_config(config)
            else:
                log_msg("No PAC URL found in registry.")
        except Exception as e:
            log_msg(f"Proxy detection error: {e}")

    def get_pac_url_from_registry(self):
        """Scans Windows Registry for AutoConfigURL"""
        locations = [
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"),
            (winreg.HKEY_CURRENT_USER, r"Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"),
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

    def get_external_proxy_string(self, pac_url):
        """Downloads PAC and extracts proxy string"""
        try:
            with urllib.request.urlopen(pac_url) as response:
                pac_content = response.read().decode('utf-8')
            
            # Try pypac
            try:
                from pypac.parser import PACFile
                pac = PACFile(pac_content)
                return pac.find_proxy_for_url("https://www.google.com", "www.google.com")
            except ImportError:
                pass 
                
            # Regex fallback
            match = re.search(r'strPxlProxy\s*=\s*"(.*?)";', pac_content)
            if match:
                return match.group(1)
        except Exception as e:
            log_msg(f"PAC processing error: {e}")
        return None

    def parse_proxy_string(self, proxy_string):
        """Parses 'PROXY host:port'"""
        if not proxy_string or "DIRECT" in proxy_string:
            return None, None
        clean_str = proxy_string.replace("PROXY ", "").split(";")[0].strip()
        if ":" in clean_str:
            host, port = clean_str.split(":")
            return host, port
        else:
            return clean_str, "80"

    def get_bing_image_info(self):
        try:
            response = requests.get(BING_API, timeout=10, proxies=self.get_proxy_dict())
            response.raise_for_status()
            data = response.json()
            
            if not data.get("images"):
                log_msg("Error: No images found in API response.")
                return None
                
            image_data = data["images"][0]
            base_url = "https://www.bing.com"
            image_url = base_url + image_data["url"]
            start_date = image_data["startdate"]
            
            return image_url, start_date
        except Exception as e:
            log_msg(f"Error fetching API: {e}")
            return None

    def download_image(self, url, date_str):
        try:
            filename = f"bing_{date_str}.jpg"
            file_path = SAVE_DIR / filename
            
            if file_path.exists():
                return file_path
                
            log_msg(f"Downloading new image to {file_path}...")
            response = requests.get(url, timeout=20, proxies=self.get_proxy_dict())
            response.raise_for_status()
            
            with open(file_path, "wb") as f:
                f.write(response.content)
                
            return file_path
        except Exception as e:
            log_msg(f"Error downloading image: {e}")
            return None

    def set_wallpaper(self, image_path):
        try:
            log_msg(f"Setting wallpaper to: {image_path}")
            ctypes.windll.user32.SystemParametersInfoW(20, 0, str(image_path), 3)
            self.current_image_path = image_path
            self.update_tray_icon(image_path)
            log_msg("Wallpaper updated successfully.")
        except Exception as e:
            log_msg(f"Error setting wallpaper: {e}")

    def update_tray_icon(self, image_path):
        if self.icon:
            try:
                if image_path and os.path.exists(image_path):
                    image = Image.open(image_path)
                    self.icon.icon = image
            except Exception as e:
                log_msg(f"Failed to update tray icon: {e}")

    def check_and_update(self, force=False):
        log_msg(f"Checking for updates (Force={force})...")
        try:
            info = self.get_bing_image_info()
            if info:
                url, date_str = info
                image_path = self.download_image(url, date_str)
                if image_path:
                    # Logic: If it's the daily image, set it.
                    # Also set it if we haven't set anything yet this session.
                    # FORCE: Always set if force=True (Manual check or Startup)
                    if force or self.current_image_path != image_path:
                         self.set_wallpaper(image_path)
                    
                    # Ensure tray icon is updated even if wallpaper didn't change
                    if self.icon and self.icon.icon is None:
                        self.update_tray_icon(image_path)
            
            # Refresh UI if open
            if self.root and self.root.winfo_viewable():
                 self.root.after(0, lambda: self.setup_ui(self.root))

        except Exception as e:
            log_msg(f"Update check failed: {e}")
            
        self.last_check = time.time()

    def background_loop(self):
        while self.running:
            # Skip automatic checks if interval is disabled (0)
            if self.check_interval > 0 and time.time() - self.last_check > self.check_interval:
                self.check_and_update(force=False)
            time.sleep(1)

    # --- UI Logic ---
    
    def show_preview_window(self):
        log_msg("Showing preview window...")
        if not self.root:
             self.create_root()
        
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.setup_ui(self.root)

    def hide_window(self):
        log_msg("Hiding window...")
        if self.root:
            self.root.withdraw()

    def create_root(self):
        self.root = tk.Tk()
        self.root.title(f"Bing Wallpaper Preview v{VERSION}")
        self.root.geometry("800x600")
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        
        # Bring to center
        # self.root.eval('tk::PlaceWindow . center') # Doesn't always work for main window
        
        self.setup_ui(self.root)

    def setup_ui(self, window):
        # Clear existing
        for widget in window.winfo_children():
            widget.destroy()
            
        # Main container
        main_frame = tk.Frame(window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 1. Main Preview Area
        preview_frame = tk.Frame(main_frame)
        preview_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        if self.current_image_path and os.path.exists(self.current_image_path):
             self.display_main_preview(preview_frame, self.current_image_path)
        else:
             lbl = tk.Label(preview_frame, text="No image loaded yet.")
             lbl.pack(pady=50)

        # 2. Horizontal Scroll List
        list_frame = tk.Frame(main_frame, height=150)
        list_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        
        # Scroll canvas
        canvas = tk.Canvas(list_frame, height=120)
        scrollbar = ttk.Scrollbar(list_frame, orient="horizontal", command=canvas.xview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(xscrollcommand=scrollbar.set)

        # Lay out images
        images = sorted(SAVE_DIR.glob("bing_*.jpg"), key=os.path.getmtime, reverse=True)
        log_msg(f"Found {len(images)} images in {SAVE_DIR}")
        
        for img_path in images[:15]: # Limit to reasonable number
            self.create_thumbnail(scrollable_frame, img_path)

        canvas.pack(side="top", fill="both", expand=True)
        scrollbar.pack(side="bottom", fill="x")

    def display_main_preview(self, parent, img_path):
        try:
            pil_img = Image.open(img_path)
            # Resize logic to keep aspect ratio
            # Max size 780x400
            pil_img.thumbnail((780, 400))
            tk_img = ImageTk.PhotoImage(pil_img)
            
            lbl = tk.Label(parent, image=tk_img)
            lbl.image = tk_img 
            lbl.pack(pady=5)
            
            tk.Label(parent, text=img_path.name, font=("Segoe UI", 10, "bold")).pack()
            tk.Label(parent, text=f"Path: {img_path}", font=("Segoe UI", 8)).pack()
            
        except Exception as e:
            tk.Label(parent, text=f"Error display image: {e}").pack()

    def create_thumbnail(self, parent, img_path):
        try:
            pil_img = Image.open(img_path)
            pil_img.thumbnail((150, 100))
            tk_img = ImageTk.PhotoImage(pil_img)
            
            # Clickable frame
            f = tk.Frame(parent, bd=2, relief="groove")
            f.pack(side=tk.LEFT, padx=5)
            
            lbl = tk.Label(f, image=tk_img, cursor="hand2")
            lbl.image = tk_img
            lbl.pack()
            
            lbl.bind("<Button-1>", lambda e, p=img_path: self.on_thumbnail_click(p))
            
            # Label
            tk.Label(f, text=img_path.name[-12:], font=("Consolas", 8)).pack() 
            
        except Exception as e:
            log_msg(f"Thumbnail error for {img_path.name}: {e}")

    def on_thumbnail_click(self, img_path):
        self.set_wallpaper(img_path)
        # Refresh to show new main image
        if self.root:
            self.setup_ui(self.root)

    # --- Tray Callbacks ---

    def on_open_preview(self, icon, item):
        # Schedule show_preview_window on the main thread
        if self.root:
            self.root.after(0, self.show_preview_window)

    def on_update_now(self, icon, item):
        threading.Thread(target=self.check_and_update, args=(True,)).start()
        
    def on_exit(self, icon, item):
        log_msg("Exiting...")
        self.running = False
        icon.stop()
        if self.root:
            self.root.quit()

    def create_menu(self):
        """Create the tray icon menu with interval submenu"""
        # Get current interval in minutes
        config = self.load_config()
        current_interval = config.get("check_interval_minutes", 720)
        
        # Helper function to create interval setter
        def make_interval_setter(minutes, label):
            def setter(icon, item):
                self.set_interval(minutes, label)
            return setter
        
        # Create interval menu items
        interval_items = []
        for label, minutes in INTERVAL_PRESETS.items():
            prefix = "● " if current_interval == minutes else "○ "
            interval_items.append(
                item(prefix + label, make_interval_setter(minutes, label))
            )
        
        # Add separator and custom option
        interval_items.append(item('─────────', None, enabled=False))
        
        # Check if current interval is a custom value
        is_custom = current_interval not in INTERVAL_PRESETS.values()
        custom_label = f"Custom ({current_interval} min)" if is_custom else "Custom..."
        prefix = "● " if is_custom else "○ "
        
        def custom_setter(icon, item):
            # Must run Tkinter dialog on main thread
            if self.root:
                self.root.after(0, self.show_custom_interval_dialog)
        
        interval_items.append(
            item(prefix + custom_label, custom_setter)
        )
        
        # Build main menu
        return pystray.Menu(
            item('Preview / Gallery', self.on_open_preview, default=True),
            item('Check for Updates', self.on_update_now),
            item('─────────', None, enabled=False),
            item('Check Interval', pystray.Menu(*interval_items)),
            item('─────────', None, enabled=False),
            item('Exit', self.on_exit)
        )


    def run(self):
        log_msg("Starting Bing Wallpaper App...")
        
        # Start background checker
        threading.Thread(target=self.background_loop, daemon=True).start()
        
        # Initial check - Force update on startup
        threading.Thread(target=self.check_and_update, args=(True,), daemon=True).start()
        
        # Pystray Setup
        default_img = Image.new('RGB', (64, 64), color = (73, 109, 137))
        self.icon = pystray.Icon("BingWallpaper", default_img, "Bing Wallpaper", self.create_menu())
        
        # Run pystray in a separate thread because it blocks
        threading.Thread(target=self.icon.run, daemon=True).start()
        
        # Main Thread: Tkinter
        # We need a root for 'after' calls even if hidden
        self.create_root()
        self.root.withdraw()
        self.root.mainloop()

if __name__ == "__main__":
    app = BingTrayApp()
    app.run()
