# installer.py
import os
import sys
import shutil
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path
import winreg
import urllib.request
import re
import json
import subprocess

try:
    from _version import __version__ as VERSION
except ImportError:
    VERSION = "1.3.3"

APP_NAME = "BingWallpaper"
# Default Path: C:\Users\<User>\AppData\Local\Programs\BingWallpaper
INSTALL_DIR = Path(os.environ["LOCALAPPDATA"]) / "Programs" / APP_NAME

class SimpleInstaller(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Install Bing Wallpaper v{VERSION}")
        self.geometry("420x550")
        self.resizable(False, False)
        self.eval('tk::PlaceWindow . center')
        
        self.create_ui()

    def create_ui(self):
        # Header
        ttk.Label(self, text=f"Bing Wallpaper v{VERSION}", font=("Segoe UI", 14, "bold")).pack(pady=10)
        
        status = "Installed" if INSTALL_DIR.exists() else "Not Installed"
        status_color = "green" if INSTALL_DIR.exists() else "red"
        ttk.Label(self, text=f"Status: {status}", foreground=status_color).pack()
        
        # Options
        opts_frame = ttk.LabelFrame(self, text="Installation Options", padding=10)
        opts_frame.pack(fill="x", padx=10, pady=10)
        
        self.startup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts_frame, text="Run at Startup", variable=self.startup_var).pack(anchor="w")
        
        self.desktop_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts_frame, text="Create Desktop Shortcut", variable=self.desktop_var).pack(anchor="w")
        
        # Network
        proxy_frame = ttk.LabelFrame(self, text="Network / Proxy (Optional)", padding=10)
        proxy_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(proxy_frame, text="Host:").grid(row=0, column=0, padx=5, sticky="w")
        self.proxy_host_var = tk.StringVar()
        ttk.Entry(proxy_frame, textvariable=self.proxy_host_var, width=20).grid(row=0, column=1, padx=5)
        
        ttk.Label(proxy_frame, text="Port:").grid(row=0, column=2, padx=5, sticky="w")
        self.proxy_port_var = tk.StringVar()
        ttk.Entry(proxy_frame, textvariable=self.proxy_port_var, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Button(proxy_frame, text="Auto-Detect Proxy", command=self.detect_proxy).grid(row=1, column=0, columnspan=4, sticky="ew", pady=(10, 0))

        # Actions
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="Install / Update", command=self.install).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Uninstall", command=self.uninstall).pack(side=tk.LEFT, padx=5)
        
        self.btn_open = ttk.Button(btn_frame, text="Open Folder", command=self.open_folder, state=tk.DISABLED)
        self.btn_open.pack(side=tk.LEFT, padx=5)
        
        # Log Window
        log_frame = ttk.LabelFrame(self, text="Log", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = tk.Text(log_frame, height=8, width=40, font=("Consolas", 8), state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.update()

    def detect_proxy(self):
        self.log("Detecting proxy...")
        try:
            pac_url = self.get_pac_url()
            if pac_url:
                proxy = self.get_proxy_from_pac(pac_url)
                if proxy:
                    self.fill_proxy(proxy)
                    return
            
            env_proxy = os.environ.get("HTTP_PROXY") or os.environ.get("HTTPS_PROXY")
            if env_proxy:
                self.fill_proxy(env_proxy)
                return

            self.log("No proxy detected.")
        except Exception as e:
            self.log(f"Detection error: {e}")

    def get_pac_url(self):
        locations = [
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"),
        ]
        for hive, subkey in locations:
            try:
                with winreg.OpenKey(hive, subkey) as key:
                    url, _ = winreg.QueryValueEx(key, "AutoConfigURL")
                    if url: return url
            except: continue
        return None

    def get_proxy_from_pac(self, pac_url):
        try:
            with urllib.request.urlopen(pac_url, timeout=5) as response:
                content = response.read().decode('utf-8')
                match = re.search(r'PROXY\s+([a-zA-Z0-9.-]+:\d+)', content)
                if match: return match.group(1)
        except: pass
        return None

    def fill_proxy(self, proxy_str):
        clean = proxy_str.replace("http://", "").replace("https://", "").split("/")[0]
        if ":" in clean:
            host, port = clean.split(":")
            self.proxy_host_var.set(host)
            self.proxy_port_var.set(port)
            self.log(f"Detected: {host}:{port}")
        else:
            self.proxy_host_var.set(clean)
            self.proxy_port_var.set("80")

    def install(self):
        self.log("Starting installation...")
        self.log(f"Target: {INSTALL_DIR}")
        
        try:
            INSTALL_DIR.mkdir(parents=True, exist_ok=True)
            
            # 1. Source Detection
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
                src_exe = Path(base_path) / "BingWallpaper.exe"
            else:
                src_exe = Path("dist/BingWallpaper.exe") 
            
            if not src_exe.exists(): src_exe = Path("BingWallpaper.exe")

            if not src_exe.exists():
                messagebox.showerror("Error", f"Source file missing:\n{src_exe}")
                self.log("Source file missing.")
                return

            # 2. File Copy
            dst_exe = INSTALL_DIR / "BingWallpaper.exe"
            self.log(f"Copying to {dst_exe}...")
            shutil.copy2(src_exe, dst_exe)
            
            if not dst_exe.exists():
                raise Exception("Copy failed - File not found at destination.")

            # 3. Config
            config_data = {
                "check_interval_minutes": 720,
                "proxy_url": self.proxy_host_var.get().strip(),
                "proxy_port": self.proxy_port_var.get().strip()
            }
            config_path = INSTALL_DIR / "config.json"
            if not config_path.exists() or (config_data["proxy_url"]):
                with open(config_path, 'w') as f:
                    json.dump(config_data, f, indent=2)

            # 4. Shortcuts (Using subprocess for safety)
            if self.desktop_var.get():
                self.create_shortcut(dst_exe, "Bing Wallpaper", "Desktop")
            
            if self.startup_var.get():
                self.create_shortcut(dst_exe, "Bing Wallpaper", "Startup")
            
            # 5. Success State
            self.log("Installation Successful!")
            self.btn_open.config(state=tk.NORMAL)
            
            # Launch
            self.log("Launching app...")
            os.startfile(dst_exe)
            
            messagebox.showinfo("Success", "Installation Complete!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Install failed: {str(e)}")
            self.log(f"Error: {str(e)}")

    def create_shortcut(self, target, name, folder):
        try:
            if folder == "Startup":
                link_dir = Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
            else:
                link_dir = Path(os.environ["USERPROFILE"]) / "Desktop"
            
            link_dir.mkdir(parents=True, exist_ok=True)
            link_path = link_dir / f"{name}.lnk"
            
            self.log(f"Creating {folder} shortcut...")
            
            # Safe PowerShell command using Subprocess
            ps_script = f"""
            $ws = New-Object -ComObject WScript.Shell
            $s = $ws.CreateShortcut('{str(link_path)}')
            $s.TargetPath = '{str(target)}'
            $s.WorkingDirectory = '{str(INSTALL_DIR)}'
            $s.Save()
            """
            
            subprocess.run(["powershell", "-Command", ps_script], check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
        except Exception as e:
            self.log(f"Shortcut Error ({folder}): {e}")

    def uninstall(self):
        try:
            os.system('taskkill /F /IM "BingWallpaper.exe" >nul 2>&1')
            if INSTALL_DIR.exists():
                shutil.rmtree(INSTALL_DIR)
            
            startup = Path(os.getenv("APPDATA")) / "Microsoft/Windows/Start Menu/Programs/Startup/Bing Wallpaper.lnk"
            desktop = Path(os.environ["USERPROFILE"]) / "Desktop" / "Bing Wallpaper.lnk"
            
            if startup.exists(): startup.unlink()
            if desktop.exists(): desktop.unlink()
                
            self.log("Uninstalled.")
            self.btn_open.config(state=tk.DISABLED)
            messagebox.showinfo("Success", "Uninstalled successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def open_folder(self):
        if INSTALL_DIR.exists():
            os.startfile(INSTALL_DIR)

if __name__ == "__main__":
    app = SimpleInstaller()
    app.mainloop()