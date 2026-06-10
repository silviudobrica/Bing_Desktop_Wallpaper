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
import time
import tempfile
import stat

try:
    from _version import __version__ as VERSION
except ImportError:
    VERSION = "1.3.6"

APP_NAME = "BingWallpaper"
PUBLISHER = "Silviu Dobrica"
# Default Path: C:\Users\<User>\AppData\Local\Programs\BingWallpaper
INSTALL_DIR = Path(os.environ["LOCALAPPDATA"]) / "Programs" / APP_NAME
UNINSTALL_REG_PATH = fr"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_NAME}"


def remove_shortcuts(log_fn=print):
    startup = Path(os.getenv("APPDATA")) / "Microsoft/Windows/Start Menu/Programs/Startup/Bing Wallpaper.lnk"
    start_menu = Path(os.getenv("APPDATA")) / "Microsoft/Windows/Start Menu/Programs/Bing Wallpaper/Bing Wallpaper.lnk"
    # Legacy cleanup path in case a shortcut was created under Desktop\Programs.
    desktop_legacy = Path(os.environ["USERPROFILE"]) / "Desktop" / "Programs" / "Bing Wallpaper.lnk"
    desktop = Path(os.environ["USERPROFILE"]) / "Desktop" / "Bing Wallpaper.lnk"

    for link in [startup, start_menu, desktop_legacy, desktop]:
        try:
            if link.exists():
                link.unlink()
                log_fn(f"Removed shortcut: {link}")
        except Exception as e:
            log_fn(f"Shortcut cleanup warning ({link}): {e}")

    for folder in [start_menu.parent, desktop_legacy.parent]:
        try:
            if folder.exists() and not any(folder.iterdir()):
                folder.rmdir()
        except Exception:
            pass


def relaunch_uninstall_worker_if_needed(log_fn=print):
    """If running from INSTALL_DIR, relaunch uninstall from temp so install dir can be deleted."""
    if not getattr(sys, 'frozen', False):
        return False

    current_exe = Path(sys.executable).resolve()
    if current_exe.parent != INSTALL_DIR:
        return False

    try:
        temp_worker_dir = Path(tempfile.gettempdir()) / APP_NAME
        temp_worker_dir.mkdir(parents=True, exist_ok=True)
        worker_exe = temp_worker_dir / "UninstallBingWallpaper.exe"

        shutil.copy2(current_exe, worker_exe)
        subprocess.Popen(
            [str(worker_exe), "--uninstall", "--worker"],
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        log_fn(f"Relaunching uninstaller worker: {worker_exe}")
        return True
    except Exception as e:
        log_fn(f"Failed to relaunch uninstall worker: {e}")
        return False


def _rmtree_onerror(log_fn, func, path, exc_info):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception as e:
        log_fn(f"Failed to remove locked path {path}: {e}")


def remove_install_dir(log_fn=print):
    if not INSTALL_DIR.exists():
        return True

    try:
        # Avoid holding a handle in the folder we are trying to remove.
        os.chdir(tempfile.gettempdir())
    except Exception:
        pass

    for attempt in range(1, 6):
        try:
            shutil.rmtree(INSTALL_DIR, onerror=lambda f, p, e: _rmtree_onerror(log_fn, f, p, e))
            if not INSTALL_DIR.exists():
                return True
        except Exception as e:
            log_fn(f"Install folder delete attempt {attempt} failed: {e}")
        time.sleep(0.5)

    # Last resort: spawn detached delayed cleanup after this process exits.
    try:
        cmd = f'cmd /c timeout /t 2 /nobreak >nul & rmdir /s /q "{str(INSTALL_DIR)}"'
        subprocess.Popen(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
        log_fn("Scheduled delayed cleanup for install folder.")
    except Exception as e:
        log_fn(f"Failed to schedule delayed cleanup: {e}")

    return not INSTALL_DIR.exists()

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

        self.start_menu_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts_frame, text="Create Start Menu Shortcut", variable=self.start_menu_var).pack(anchor="w")
        
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
            self.stop_running_app()
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
            self.copy_with_retries(src_exe, dst_exe)

            uninstall_exe = None
            if getattr(sys, 'frozen', False):
                uninstall_exe = INSTALL_DIR / "UninstallBingWallpaper.exe"
                shutil.copy2(Path(sys.executable), uninstall_exe)
            
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

            if self.start_menu_var.get():
                self.create_shortcut(dst_exe, "Bing Wallpaper", "StartMenu")
            
            if self.startup_var.get():
                self.create_shortcut(dst_exe, "Bing Wallpaper", "Startup")

            self.register_uninstall_entry(uninstall_exe)
            
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

    def copy_with_retries(self, src_exe, dst_exe, retries=8, delay=0.5):
        last_error = None
        temp_target = dst_exe.with_suffix(".new")

        for attempt in range(1, retries + 1):
            try:
                if temp_target.exists():
                    temp_target.unlink()
                shutil.copy2(src_exe, temp_target)
                os.replace(temp_target, dst_exe)
                return
            except PermissionError as e:
                last_error = e
                self.log(f"File is locked (attempt {attempt}/{retries}). Retrying...")
                self.stop_running_app()
                time.sleep(delay)
            except Exception as e:
                last_error = e
                break

        if temp_target.exists():
            try:
                temp_target.unlink()
            except Exception:
                pass

        raise Exception(f"Unable to replace app executable after {retries} attempts: {last_error}")

    def create_shortcut(self, target, name, folder):
        try:
            if folder == "Startup":
                link_dir = Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
            elif folder == "StartMenu":
                link_dir = Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Bing Wallpaper"
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

    def register_uninstall_entry(self, uninstall_exe):
        try:
            if uninstall_exe and uninstall_exe.exists():
                uninstall_cmd = f'"{str(uninstall_exe)}" --uninstall'
                icon_path = str(INSTALL_DIR / "BingWallpaper.exe")
            elif getattr(sys, 'frozen', False):
                uninstall_cmd = f'"{str(Path(sys.executable))}" --uninstall'
                icon_path = str(INSTALL_DIR / "BingWallpaper.exe")
            else:
                script_path = Path(__file__).resolve()
                uninstall_cmd = f'"{sys.executable}" "{str(script_path)}" --uninstall'
                icon_path = str(script_path)

            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, UNINSTALL_REG_PATH)
            with key:
                winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "Bing Wallpaper")
                winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, VERSION)
                winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
                winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(INSTALL_DIR))
                winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, uninstall_cmd)
                winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, icon_path)
                winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)

                app_exe = INSTALL_DIR / "BingWallpaper.exe"
                if app_exe.exists():
                    winreg.SetValueEx(key, "EstimatedSize", 0, winreg.REG_DWORD, max(1, app_exe.stat().st_size // 1024))
            self.log("Registered app in Windows Installed Apps.")
        except Exception as e:
            self.log(f"Uninstall registration error: {e}")

    def remove_uninstall_entry(self):
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, UNINSTALL_REG_PATH)
            self.log("Removed Installed Apps registration.")
        except FileNotFoundError:
            pass
        except Exception as e:
            self.log(f"Uninstall registry cleanup error: {e}")

    def stop_running_app(self):
        try:
            # Try graceful stop first.
            subprocess.run(
                ["taskkill", "/IM", "BingWallpaper.exe"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW,
                check=False
            )

            for _ in range(10):
                check = subprocess.run(
                    ["tasklist", "/FI", "IMAGENAME eq BingWallpaper.exe"],
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    check=False
                )
                if "BingWallpaper.exe" not in check.stdout:
                    return
                time.sleep(0.2)

            # Fall back to force kill if still running.
            subprocess.run(
                ["taskkill", "/F", "/IM", "BingWallpaper.exe"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW,
                check=False
            )
        except Exception as e:
            self.log(f"Process stop warning: {e}")

    def uninstall(self):
        try:
            self.stop_running_app()

            remove_shortcuts(self.log)

            self.remove_uninstall_entry()

            removed = remove_install_dir(self.log)
            if not removed:
                self.log("Install folder cleanup is pending.")
                
            self.log("Uninstalled.")
            self.btn_open.config(state=tk.DISABLED)
            messagebox.showinfo("Success", "Uninstalled successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def open_folder(self):
        if INSTALL_DIR.exists():
            os.startfile(INSTALL_DIR)

if __name__ == "__main__":
    if "--uninstall" in sys.argv:
        class _HeadlessInstaller:
            def log(self, msg):
                print(msg)

            stop_running_app = SimpleInstaller.stop_running_app
            remove_uninstall_entry = SimpleInstaller.remove_uninstall_entry

        headless = _HeadlessInstaller()
        try:
            if "--worker" not in sys.argv and relaunch_uninstall_worker_if_needed(headless.log):
                sys.exit(0)

            headless.stop_running_app()

            remove_shortcuts(headless.log)

            headless.remove_uninstall_entry()

            remove_install_dir(headless.log)
        except Exception as e:
            print(f"Uninstall failed: {e}")
        sys.exit(0)

    app = SimpleInstaller()
    app.mainloop()