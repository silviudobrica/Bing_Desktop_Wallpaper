import os
import sys
import ctypes
import requests
import datetime
import time
import threading
import logging
from pathlib import Path
from PIL import Image, ImageTk
import pystray
from pystray import MenuItem as item
import tkinter as tk
from tkinter import ttk

# Configuration
BING_API = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US"
SAVE_DIR = Path(os.environ["USERPROFILE"]) / "Pictures" / "Bing"
APP_DIR = Path(os.environ["ProgramFiles"]) / "BingWallpaper"
VERSION = "1.0.0"

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
        self.check_interval = 15 * 60  # 15 minutes
        self.current_image_path = None
        self.running = True
        
        # Ensure save directory exists
        SAVE_DIR.mkdir(parents=True, exist_ok=True)
        log_msg(f"Initializing Bing Wallpaper App v{VERSION}")

    def get_bing_image_info(self):
        try:
            response = requests.get(BING_API, timeout=10)
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
            response = requests.get(url, timeout=20)
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
            if time.time() - self.last_check > self.check_interval:
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

    def run(self):
        log_msg("Starting Bing Wallpaper App...")
        
        # Start background checker
        threading.Thread(target=self.background_loop, daemon=True).start()
        
        # Initial check - Force update on startup
        threading.Thread(target=self.check_and_update, args=(True,), daemon=True).start()
        
        # Pystray Setup
        default_img = Image.new('RGB', (64, 64), color = (73, 109, 137))
        menu = pystray.Menu(
            item('Preview / Gallery', self.on_open_preview, default=True),
            item('Check for Updates', self.on_update_now),
            item('Exit', self.on_exit)
        )
        self.icon = pystray.Icon("BingWallpaper", default_img, "Bing Wallpaper", menu)
        
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
