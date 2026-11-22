# platform_linux.py
"""Linux implementation of the service's Platform interface, KDE Plasma
first, GNOME second, anything with XDG portals third. Works both from a
Flatpak sandbox and from a plain checkout.

Wallpaper   KDE: Plasma shell scripting over D-Bus (silent).
            GNOME outside Flatpak: gsettings. Everything else: the XDG
            Wallpaper portal (the desktop may ask for confirmation).
Lock screen KDE: kscreenlockerrc. GNOME outside Flatpak: gsettings.
            Else: the Wallpaper portal with set-on=lockscreen.
Proxies     Looked up fresh before every request: the XDG ProxyResolver
            portal (what the desktop's proxy settings say right now), then
            KDE's kioslaverc, then GNOME's gsettings, then the environment.
Autostart   XDG Background portal in Flatpak, ~/.config/autostart otherwise.
Notify      XDG Notification portal, notify-send as a fallback.

Only jeepney (pure Python) is needed for D-Bus; no Qt here, so this module
is testable headless and reusable by any front end.
"""
import configparser
import os
import re
import secrets
import shutil
import subprocess
import urllib.parse
from dataclasses import dataclass
from pathlib import Path

import bing_core as core
from bing_service import Platform, log_msg

APP_ID = "io.github.silviudobrica.BingWallpaper"
APP_DISPLAY_NAME = "Bing Wallpaper"
APP_DIR_NAME = "bing-wallpaper"

PORTAL_BUS = "org.freedesktop.portal.Desktop"
PORTAL_PATH = "/org/freedesktop/portal/desktop"
PLASMA_BUS = "org.kde.plasmashell"

# Plasma shell script: switch every desktop containment to the image plugin
# and point it at the file. Same thing plasma-apply-wallpaperimage does.
PLASMA_WALLPAPER_SCRIPT = """
var all = desktops();
for (var i = 0; i < all.length; i++) {
    var d = all[i];
    d.wallpaperPlugin = "org.kde.image";
    d.currentConfigGroup = Array("Wallpaper", "org.kde.image", "General");
    d.writeConfig("Image", %(uri)s);
}
"""


@dataclass
class LinuxPaths:
    data_dir: Path
    config_file: Path
    log_dir: Path
    image_dir: Path


def _xdg(var, default):
    value = os.environ.get(var, "").strip()
    return Path(value) if value else Path.home() / default


def xdg_pictures_dir():
    """XDG_PICTURES_DIR from user-dirs.dirs (Flatpak copies the host's file
    into the sandbox config dir), else ~/Pictures."""
    candidates = [_xdg("XDG_CONFIG_HOME", ".config") / "user-dirs.dirs", Path.home() / ".config" / "user-dirs.dirs"]
    for candidate in candidates:
        try:
            for line in candidate.read_text(encoding="utf-8").splitlines():
                m = re.match(r'\s*XDG_PICTURES_DIR\s*=\s*"?([^"]*)"?', line)
                if m:
                    return Path(os.path.expandvars(m.group(1).replace("$HOME", str(Path.home()))))
        except OSError:
            continue
    return Path.home() / "Pictures"


def linux_paths():
    data_dir = _xdg("XDG_DATA_HOME", ".local/share") / APP_DIR_NAME
    return LinuxPaths(
        data_dir=data_dir,
        config_file=_xdg("XDG_CONFIG_HOME", ".config") / APP_DIR_NAME / "config.json",
        log_dir=data_dir / "logs",
        image_dir=xdg_pictures_dir() / "Bing",
    )


def detect_desktop(environ=None):
    environ = os.environ if environ is None else environ
    current = (environ.get("XDG_CURRENT_DESKTOP", "") + ":" + environ.get("DESKTOP_SESSION", "")).upper()
    if "KDE" in current or "PLASMA" in current:
        return "kde"
    if "GNOME" in current:
        return "gnome"
    return "other"


def in_flatpak(environ=None):
    """FLATPAK_ID is set inside the sandbox; /.flatpak-info is checked only
    for the real process environment (tests pass their own environ)."""
    if environ is not None:
        return bool(environ.get("FLATPAK_ID"))
    return bool(os.environ.get("FLATPAK_ID")) or Path("/.flatpak-info").exists()


# ---------------------------------------------------------------- proxies
def proxies_from_environment(environ=None):
    environ = os.environ if environ is None else environ
    result = {}
    for scheme in ("http", "https"):
        for key in (f"{scheme}_proxy", f"{scheme.upper()}_PROXY", "all_proxy", "ALL_PROXY"):
            value = environ.get(key, "").strip()
            if value:
                result[scheme] = value
                break
    return result


def proxies_from_portal_result(entries):
    """ProxyResolver.Lookup returns e.g. ['http://proxy:3128'] or
    ['direct://']. Only the first usable entry is taken."""
    for entry in entries or []:
        entry = str(entry).strip()
        if not entry or entry.startswith("direct"):
            continue
        if entry.startswith(("http://", "https://", "socks5://", "socks5h://", "socks4://")):
            return {"http": entry, "https": entry}
        if entry.startswith("socks://"):
            fixed = "socks5h://" + entry[len("socks://"):]
            return {"http": fixed, "https": fixed}
    return {}


def _normalize_proxy_url(value):
    value = str(value or "").strip()
    if not value:
        return ""
    if "://" not in value:
        value = "http://" + value
    # KDE stores "http://host 8080" when typed in the dialog with a space.
    value = re.sub(r"\s+(\d+)$", r":\1", value)
    return value


def proxies_from_kioslaverc(text, environ=None, pac_fetch=None):
    """Parse KDE's [Proxy Settings] group. ProxyType: 0 none, 1 manual,
    2 PAC script, 3 auto-detect (WPAD, unsupported), 4 environment variable
    names. pac_fetch(url) -> script text, used for type 2."""
    environ = os.environ if environ is None else environ
    parser = configparser.ConfigParser(interpolation=None, strict=False, allow_no_value=True)
    parser.optionxform = str
    try:
        parser.read_string(text)
    except configparser.Error:
        return {}
    if not parser.has_section("Proxy Settings"):
        return {}
    group = parser["Proxy Settings"]
    proxy_type = str(group.get("ProxyType", "0")).strip()
    if proxy_type == "1":
        result = {}
        for scheme, key in (("http", "httpProxy"), ("https", "httpsProxy")):
            url = _normalize_proxy_url(group.get(key, ""))
            if url:
                result[scheme] = url
        if "https" not in result and "http" in result:
            result["https"] = result["http"]
        return result
    if proxy_type == "4":
        result = {}
        for scheme, key in (("http", "httpProxy"), ("https", "httpsProxy")):
            var = str(group.get(key, "")).strip()
            url = _normalize_proxy_url(environ.get(var, "")) if var else ""
            if url:
                result[scheme] = url
        return result
    if proxy_type == "2" and pac_fetch:
        pac_url = str(group.get("Proxy Config Script", "")).strip()
        if pac_url:
            try:
                proxy = core.extract_pac_proxy(pac_fetch(pac_url) or "")
            except Exception:
                proxy = None
            if proxy:
                url = _normalize_proxy_url(proxy)
                return {"http": url, "https": url}
    return {}


def proxies_from_gsettings(get):
    """get(schema, key) -> raw `gsettings get` output. GNOME manual mode."""
    try:
        mode = get("org.gnome.system.proxy", "mode").strip().strip("'")
        if mode != "manual":
            return {}
        result = {}
        for scheme in ("http", "https"):
            host = get(f"org.gnome.system.proxy.{scheme}", "host").strip().strip("'")
            port = get(f"org.gnome.system.proxy.{scheme}", "port").strip()
            if host and port and port != "0":
                result[scheme] = f"http://{host}:{port}"
        if "https" not in result and "http" in result:
            result["https"] = result["http"]
        return result
    except Exception:
        return {}


def proxy_host_port(proxies):
    """(host, port) of the https (or http) proxy for the Settings dialog."""
    url = proxies.get("https") or proxies.get("http") or ""
    if not url:
        return "", ""
    parsed = urllib.parse.urlsplit(url if "://" in url else "http://" + url)
    return parsed.hostname or "", str(parsed.port) if parsed.port else ""


# ---------------------------------------------------------------- D-Bus
class DBus:
    """Thin jeepney wrapper: plain calls and portal request/response."""

    def __init__(self, bus="SESSION"):
        self.bus = bus

    def call(self, bus_name, path, interface, method, signature="", args=(), timeout=30):
        from jeepney import DBusAddress, MessageType, new_method_call
        from jeepney.io.blocking import open_dbus_connection
        with open_dbus_connection(bus=self.bus) as conn:
            addr = DBusAddress(path, bus_name=bus_name, interface=interface)
            reply = conn.send_and_get_reply(new_method_call(addr, method, signature, tuple(args)), timeout=timeout)
            if reply.header.message_type == MessageType.error:
                raise RuntimeError(reply.body[0] if reply.body else "D-Bus error")
            return reply.body

    def portal_request(self, interface, method, signature, build_args, timeout=180):
        """Call a portal method that answers through a Request object.
        build_args(handle_token) -> argument tuple. Returns (code, results):
        code 0 success, 1 cancelled by the user, 2 other error."""
        from jeepney import DBusAddress, MessageType, new_method_call
        from jeepney.bus_messages import MatchRule, message_bus
        from jeepney.io.blocking import Proxy, open_dbus_connection
        with open_dbus_connection(bus=self.bus) as conn:
            token = "bingwallpaper" + secrets.token_hex(8)
            sender = conn.unique_name[1:].replace(".", "_")
            handle = f"/org/freedesktop/portal/desktop/request/{sender}/{token}"
            rule = MatchRule(type="signal", interface="org.freedesktop.portal.Request", member="Response", path=handle)
            Proxy(message_bus, conn).AddMatch(rule)
            with conn.filter(rule) as queue:
                addr = DBusAddress(PORTAL_PATH, bus_name=PORTAL_BUS, interface=interface)
                reply = conn.send_and_get_reply(new_method_call(addr, method, signature, tuple(build_args(token))), timeout=30)
                if reply.header.message_type == MessageType.error:
                    raise RuntimeError(reply.body[0] if reply.body else "D-Bus error")
                signal = conn.recv_until_filtered(queue, timeout=timeout)
                code, results = signal.body
                return int(code), dict(results)


# ---------------------------------------------------------------- platform
class LinuxPlatform(Platform):
    autosave_detected_proxy = False
    supports_lock_screen = True

    def __init__(self, paths, launch_command=None, dbus=None, environ=None):
        super().__init__()
        self.paths = paths
        self.launch_command = list(launch_command or [APP_DIR_NAME, "--hidden"])
        self.environ = os.environ if environ is None else environ
        self.desktop = detect_desktop(self.environ)
        self.in_flatpak = in_flatpak(self.environ)
        self.dbus = dbus or DBus()
        self.name = f"linux-{self.desktop}" + ("-flatpak" if self.in_flatpak else "")
        self._portal_lock_screen_done_for = None
        self._autostart_flag = Path(paths.data_dir) / "autostart.enabled"

    def describe(self):
        label = {"kde": "KDE Plasma", "gnome": "GNOME"}.get(self.desktop, "Linux desktop")
        return f"{label} ({'Flatpak' if self.in_flatpak else 'native'})"

    # ------------------------------------------------------------ wallpaper
    def set_wallpaper(self, image_path):
        uri = Path(image_path).resolve().as_uri()
        if self.desktop == "kde":
            self._set_plasma_wallpaper(uri)
            return
        if self.desktop == "gnome" and not self.in_flatpak and shutil.which("gsettings"):
            for key in ("picture-uri", "picture-uri-dark"):
                subprocess.run(["gsettings", "set", "org.gnome.desktop.background", key, uri], check=True, timeout=15)
            subprocess.run(["gsettings", "set", "org.gnome.desktop.background", "picture-options", "zoom"], timeout=15)
            return
        self._portal_set_wallpaper(uri, with_lock_screen=self._lock_screen_wanted())

    def _set_plasma_wallpaper(self, uri):
        script = PLASMA_WALLPAPER_SCRIPT % {"uri": _js_string(uri)}
        try:
            self.dbus.call(PLASMA_BUS, "/PlasmaShell", "org.kde.PlasmaShell", "evaluateScript", "s", (script,))
            return
        except Exception as e:
            if self.in_flatpak or not shutil.which("plasma-apply-wallpaperimage"):
                raise RuntimeError(f"Plasma shell did not accept the wallpaper over D-Bus: {e}") from e
            log_msg(f"Plasma D-Bus failed ({e}); falling back to plasma-apply-wallpaperimage")
        subprocess.run(["plasma-apply-wallpaperimage", urllib.parse.unquote(urllib.parse.urlsplit(uri).path)],
                       check=True, capture_output=True, timeout=30)

    def _lock_screen_wanted(self):
        return bool(self.service and self.service.config.get("set_lock_screen", False))

    def _portal_set_wallpaper(self, uri, with_lock_screen=False):
        target = "both" if with_lock_screen else "background"
        code, _ = self.dbus.portal_request(
            "org.freedesktop.portal.Wallpaper", "SetWallpaperURI", "ssa{sv}",
            lambda token: ("", uri, {"show-preview": ("b", False), "set-on": ("s", target), "handle_token": ("s", token)}),
        )
        if code == 1:
            raise RuntimeError("the desktop's wallpaper dialog was cancelled")
        if code != 0:
            raise RuntimeError(f"the Wallpaper portal refused the image (code {code})")
        self._portal_lock_screen_done_for = uri if with_lock_screen else None

    # ------------------------------------------------------------ lock screen
    def set_lock_screen(self, image_path):
        uri = Path(image_path).resolve().as_uri()
        if self.desktop == "kde":
            try:
                self._set_kde_lock_screen(uri)
            except Exception as e:
                return False, str(e)
            return True, "kscreenlockerrc"
        if self.desktop == "gnome" and not self.in_flatpak and shutil.which("gsettings"):
            subprocess.run(["gsettings", "set", "org.gnome.desktop.screensaver", "picture-uri", uri], check=True, timeout=15)
            return True, "gsettings"
        if self._portal_lock_screen_done_for == uri:
            return True, "Wallpaper portal (set together with the background)"
        try:
            code, _ = self.dbus.portal_request(
                "org.freedesktop.portal.Wallpaper", "SetWallpaperURI", "ssa{sv}",
                lambda token: ("", uri, {"show-preview": ("b", False), "set-on": ("s", "lockscreen"), "handle_token": ("s", token)}),
            )
        except Exception as e:
            return False, str(e)
        if code == 0:
            return True, "Wallpaper portal"
        return False, "cancelled in the desktop's dialog" if code == 1 else f"Wallpaper portal refused it (code {code})"

    def host_config_dir(self):
        """The desktop's own config dir. Inside Flatpak XDG_CONFIG_HOME points
        at the sandbox's private ~/.var/app/<id>/config, which Plasma never
        reads; files granted with --filesystem=xdg-config/<file> appear at
        their real host path, ~/.config/<file>."""
        if self.in_flatpak:
            return Path.home() / ".config"
        return _xdg("XDG_CONFIG_HOME", ".config")

    def kscreenlockerrc_path(self):
        return self.host_config_dir() / "kscreenlockerrc"

    def _set_kde_lock_screen(self, uri):
        path = self.kscreenlockerrc_path()
        if not self.in_flatpak and shutil.which("kwriteconfig6"):
            base = ["kwriteconfig6", "--file", str(path)]
            subprocess.run(base + ["--group", "Greeter", "--key", "WallpaperPlugin", "org.kde.image"], check=True, timeout=15)
            subprocess.run(base + ["--group", "Greeter", "--group", "Wallpaper", "--group", "org.kde.image",
                                   "--group", "General", "--key", "Image", uri], check=True, timeout=15)
            return
        if self.in_flatpak and not path.exists():
            raise RuntimeError(
                "kscreenlockerrc does not exist yet; open System Settings > Screen Locking > "
                "Appearance once so Plasma creates it, then try again")
        text = path.read_text(encoding="utf-8") if path.exists() else ""
        content = kde_lock_screen_config(text, uri)
        if self.in_flatpak:
            # The file is a single-file bind mount from the host: it can be
            # rewritten in place, but a rename over it fails with EBUSY.
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
                f.flush()
                os.fsync(f.fileno())
            return
        path.parent.mkdir(parents=True, exist_ok=True)
        temp = path.with_name(path.name + ".tmp")
        temp.write_text(content, encoding="utf-8")
        os.replace(temp, path)

    # ------------------------------------------------------------ proxies
    def system_proxies(self, url=None):
        url = url or core.BING_IMAGE_HOST + "/"
        try:
            (entries,) = self.dbus.call(PORTAL_BUS, PORTAL_PATH, "org.freedesktop.portal.ProxyResolver", "Lookup", "s", (url,), timeout=5)
            found = proxies_from_portal_result(entries)
            if found or entries:
                return found  # the portal answered, 'direct://' included
        except Exception:
            pass  # no portal on this session; read the desktop's settings directly
        if self.desktop == "kde":
            found = self._proxies_from_kde()
            if found:
                return found
        if self.desktop == "gnome" and shutil.which("gsettings"):
            found = proxies_from_gsettings(self._gsettings_get)
            if found:
                return found
        return proxies_from_environment(self.environ)

    def _proxies_from_kde(self):
        for candidate in (self.host_config_dir() / "kioslaverc", Path.home() / ".config" / "kioslaverc"):
            try:
                text = candidate.read_text(encoding="utf-8")
            except OSError:
                continue
            fetch = self.service.fetch_text if self.service else None
            return proxies_from_kioslaverc(text, self.environ, fetch)
        return {}

    @staticmethod
    def _gsettings_get(schema, key):
        return subprocess.run(["gsettings", "get", schema, key], capture_output=True, text=True, timeout=10).stdout

    def detect_proxy_values(self):
        return proxy_host_port(self.system_proxies())

    # ------------------------------------------------------------ notify
    def notify(self, title, message):
        """Show a desktop notification; True when it was handed over.
        Priority is "normal": Plasma shows no popup for "low" by default."""
        try:
            self.dbus.call(PORTAL_BUS, PORTAL_PATH, "org.freedesktop.portal.Notification", "AddNotification", "sa{sv}",
                           ("bing-wallpaper", {"title": ("s", title), "body": ("s", message), "priority": ("s", "normal")}), timeout=5)
            return True
        except Exception as e:
            log_msg(f"Notification portal failed: {e}", "warning")
        if shutil.which("notify-send"):
            done = subprocess.run(["notify-send", "--app-name", APP_DISPLAY_NAME, title, message], timeout=10)
            return done.returncode == 0
        return False

    # ------------------------------------------------------------ autostart
    def autostart_file(self):
        return _xdg("XDG_CONFIG_HOME", ".config") / "autostart" / f"{APP_ID}.desktop"

    def is_startup_enabled(self):
        if self.in_flatpak:
            return self._autostart_flag.exists()
        return self.autostart_file().exists()

    def set_startup_enabled(self, enabled):
        if self.in_flatpak:
            code, results = self.dbus.portal_request(
                "org.freedesktop.portal.Background", "RequestBackground", "sa{sv}",
                lambda token: ("", {
                    "reason": ("s", f"{APP_DISPLAY_NAME} updates the wallpaper daily in the background"),
                    "autostart": ("b", bool(enabled)),
                    "commandline": ("as", self.launch_command),
                    "dbus-activatable": ("b", False),
                    "handle_token": ("s", token),
                }),
            )
            if code != 0:
                raise RuntimeError("the desktop declined the autostart request" if code == 1 else f"Background portal error (code {code})")
            granted = bool(results.get("autostart", (None, enabled))[1] if isinstance(results.get("autostart"), tuple) else results.get("autostart", enabled))
            if enabled and not granted:
                raise RuntimeError("the desktop did not allow autostart; check Settings > Autostart")
            self._autostart_flag.parent.mkdir(parents=True, exist_ok=True)
            if enabled:
                self._autostart_flag.write_text("1", encoding="utf-8")
            elif self._autostart_flag.exists():
                self._autostart_flag.unlink()
            return
        path = self.autostart_file()
        if enabled:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(autostart_desktop_entry(self.launch_command), encoding="utf-8")
        elif path.exists():
            path.unlink()

    # ------------------------------------------------------------ self-update
    @property
    def supports_self_update(self):
        """A sandboxed app cannot replace itself, but with the
        org.freedesktop.Flatpak permission it may ask the host to run
        `flatpak install` on a bundle it downloaded (the manifest grants it)."""
        return self.in_flatpak and shutil.which("flatpak-spawn") is not None

    def flatpak_installation(self):
        """'user' or 'system': where this app is installed, from /.flatpak-info."""
        try:
            return flatpak_installation_from_info(Path("/.flatpak-info").read_text(encoding="utf-8"), Path.home())
        except OSError:
            return "user"

    def install_update_bundle(self, bundle_path, runner=None):
        """Install a downloaded .flatpak bundle over this app via the host.
        Returns (ok, detail)."""
        runner = runner or self._run_host_command
        cmd = ["flatpak", "install", f"--{self.flatpak_installation()}", "-y", "--noninteractive",
               "--reinstall", str(bundle_path)]
        code, output = runner(cmd, 900)
        tail = "\n".join(line for line in output.strip().splitlines()[-5:])
        if code == 0:
            return True, tail or "installed"
        return False, tail or f"flatpak install exited with {code}"

    def _run_host_command(self, cmd, timeout):
        proc = subprocess.run(["flatpak-spawn", "--host"] + list(cmd), capture_output=True, text=True, timeout=timeout)
        return proc.returncode, (proc.stdout or "") + (proc.stderr or "")

    def restart_app(self):
        """Relaunch the (updated) app from the host after this process exits."""
        subprocess.Popen(
            ["flatpak-spawn", "--host", "sh", "-c", f"sleep 2; exec flatpak run {APP_ID} --hidden"],
            start_new_session=True, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )

    # ------------------------------------------------------------ misc
    def open_path(self, path):
        subprocess.Popen(["xdg-open", str(path)])

    def open_url(self, url):
        subprocess.Popen(["xdg-open", str(url)])


# ---------------------------------------------------------------- helpers
def flatpak_installation_from_info(info_text, home):
    """'user' when /.flatpak-info's app-path is under the home directory,
    else 'system'."""
    parser = configparser.ConfigParser(interpolation=None, strict=False)
    parser.optionxform = str
    try:
        parser.read_string(info_text)
        app_path = parser.get("Instance", "app-path", fallback="")
    except configparser.Error:
        app_path = ""
    return "user" if app_path.startswith(str(home)) else "system"


def _js_string(value):
    return '"' + str(value).replace("\\", "\\\\").replace('"', '\\"') + '"'


def kde_lock_screen_config(text, uri):
    """Return kscreenlockerrc content with the image plugin pointed at uri.
    Other groups and keys are preserved."""
    parser = configparser.ConfigParser(interpolation=None, strict=False, allow_no_value=True)
    parser.optionxform = str
    parser.read_string(text)
    greeter = "Greeter"
    image = "Greeter][Wallpaper][org.kde.image][General"
    for section in (greeter, image):
        if not parser.has_section(section):
            parser.add_section(section)
    parser[greeter]["WallpaperPlugin"] = "org.kde.image"
    parser[image]["Image"] = uri
    parser[image].setdefault("FillMode", "2")
    parser[image]["PreviewImage"] = urllib.parse.unquote(urllib.parse.urlsplit(uri).path)  # KDE stores a plain path here
    import io
    out = io.StringIO()
    parser.write(out, space_around_delimiters=False)
    return out.getvalue().rstrip("\n") + "\n"


def autostart_desktop_entry(command):
    exec_line = " ".join(_shell_quote(part) for part in command)
    return (
        "[Desktop Entry]\n"
        "Type=Application\n"
        f"Name={APP_DISPLAY_NAME}\n"
        f"Exec={exec_line}\n"
        f"Icon={APP_ID}\n"
        "X-GNOME-Autostart-enabled=true\n"
        "X-KDE-autostart-after=panel\n"
    )


def _shell_quote(part):
    return part if re.fullmatch(r"[A-Za-z0-9_./=:-]+", part) else '"' + part.replace('"', '\\"') + '"'
