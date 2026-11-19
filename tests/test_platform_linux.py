"""Pure-function tests for the Linux platform (no D-Bus, no desktop)."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import platform_linux as pl  # noqa: E402


@pytest.fixture(autouse=True)
def isolated_home(tmp_path, monkeypatch):
    """No test may touch the real home: Flatpak paths resolve via HOME."""
    home = tmp_path / "home"
    home.mkdir()
    monkeypatch.setenv("HOME", str(home))
    monkeypatch.setenv("XDG_CONFIG_HOME", str(home / ".config"))
    monkeypatch.setenv("XDG_DATA_HOME", str(home / ".local" / "share"))
    return home

# Linux desktop integration: path conventions (HOME, XDG) are POSIX-only.
pytestmark = pytest.mark.skipif(sys.platform != "linux", reason="Linux platform tests")


# ---------------------------------------------------------------- environment
def test_detect_desktop():
    assert pl.detect_desktop({"XDG_CURRENT_DESKTOP": "KDE"}) == "kde"
    assert pl.detect_desktop({"XDG_CURRENT_DESKTOP": "ubuntu:GNOME"}) == "gnome"
    assert pl.detect_desktop({"DESKTOP_SESSION": "plasmawayland"}) == "kde"
    assert pl.detect_desktop({"XDG_CURRENT_DESKTOP": "sway"}) == "other"
    assert pl.detect_desktop({}) == "other"


def test_in_flatpak_from_env():
    assert pl.in_flatpak({"FLATPAK_ID": "x"}) is True


def test_linux_paths_follow_xdg(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "cfg"))
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "data"))
    monkeypatch.setenv("HOME", str(tmp_path))
    (tmp_path / "cfg").mkdir()
    (tmp_path / "cfg" / "user-dirs.dirs").write_text('XDG_PICTURES_DIR="$HOME/Bilder"\n')
    paths = pl.linux_paths()
    assert paths.config_file == tmp_path / "cfg" / "bing-wallpaper" / "config.json"
    assert paths.log_dir == tmp_path / "data" / "bing-wallpaper" / "logs"
    assert paths.image_dir == tmp_path / "Bilder" / "Bing"


def test_pictures_dir_falls_back_to_home(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "nothing"))
    monkeypatch.setenv("HOME", str(tmp_path))
    assert pl.xdg_pictures_dir() == tmp_path / "Pictures"


# ---------------------------------------------------------------- proxies
def test_proxies_from_portal_result():
    assert pl.proxies_from_portal_result(["direct://"]) == {}
    assert pl.proxies_from_portal_result(["http://proxy.corp:3128", "direct://"]) == {
        "http": "http://proxy.corp:3128", "https": "http://proxy.corp:3128"}
    assert pl.proxies_from_portal_result(["socks://h:1080"]) == {"http": "socks5h://h:1080", "https": "socks5h://h:1080"}
    assert pl.proxies_from_portal_result([]) == {}


def test_proxies_from_environment_prefers_scheme_specific():
    env = {"http_proxy": "http://a:1", "HTTPS_PROXY": "http://b:2", "all_proxy": "http://c:3"}
    assert pl.proxies_from_environment(env) == {"http": "http://a:1", "https": "http://b:2"}
    assert pl.proxies_from_environment({"ALL_PROXY": "http://c:3"}) == {"http": "http://c:3", "https": "http://c:3"}
    assert pl.proxies_from_environment({}) == {}


KIO_MANUAL = """[Proxy Settings]
ProxyType=1
httpProxy=http://proxy.corp 8080
httpsProxy=https-proxy.corp:8443
NoProxyFor=localhost
"""


def test_kioslaverc_manual():
    assert pl.proxies_from_kioslaverc(KIO_MANUAL) == {"http": "http://proxy.corp:8080", "https": "http://https-proxy.corp:8443"}


def test_kioslaverc_manual_http_only_covers_https():
    text = "[Proxy Settings]\nProxyType=1\nhttpProxy=http://p:1\n"
    assert pl.proxies_from_kioslaverc(text) == {"http": "http://p:1", "https": "http://p:1"}


def test_kioslaverc_none_and_wpad_and_missing():
    assert pl.proxies_from_kioslaverc("[Proxy Settings]\nProxyType=0\nhttpProxy=http://p:1\n") == {}
    assert pl.proxies_from_kioslaverc("[Proxy Settings]\nProxyType=3\n") == {}
    assert pl.proxies_from_kioslaverc("[Other]\nx=1\n") == {}
    assert pl.proxies_from_kioslaverc("not an ini [[[") == {}


def test_kioslaverc_env_variable_names():
    text = "[Proxy Settings]\nProxyType=4\nhttpProxy=HTTP_PROXY\nhttpsProxy=HTTPS_PROXY\n"
    env = {"HTTP_PROXY": "http://e:1", "HTTPS_PROXY": "e2:2"}
    assert pl.proxies_from_kioslaverc(text, env) == {"http": "http://e:1", "https": "http://e2:2"}
    assert pl.proxies_from_kioslaverc(text, {}) == {}


def test_kioslaverc_pac_script_is_fetched_fresh():
    text = "[Proxy Settings]\nProxyType=2\nProxy Config Script=http://corp/proxy.pac\n"
    fetched = []

    def fetch(url):
        fetched.append(url)
        return 'function FindProxyForURL(u, h) { return "PROXY pac.corp:3128; DIRECT"; }'

    assert pl.proxies_from_kioslaverc(text, {}, fetch) == {"http": "http://pac.corp:3128", "https": "http://pac.corp:3128"}
    assert fetched == ["http://corp/proxy.pac"]
    assert pl.proxies_from_kioslaverc(text, {}, None) == {}


def test_proxies_from_gsettings():
    values = {("org.gnome.system.proxy", "mode"): "'manual'",
              ("org.gnome.system.proxy.http", "host"): "'gp'", ("org.gnome.system.proxy.http", "port"): "3128",
              ("org.gnome.system.proxy.https", "host"): "''", ("org.gnome.system.proxy.https", "port"): "0"}
    assert pl.proxies_from_gsettings(lambda s, k: values[(s, k)]) == {"http": "http://gp:3128", "https": "http://gp:3128"}
    values[("org.gnome.system.proxy", "mode")] = "'none'"
    assert pl.proxies_from_gsettings(lambda s, k: values[(s, k)]) == {}


def test_proxy_host_port():
    assert pl.proxy_host_port({"https": "http://proxy.corp:3128"}) == ("proxy.corp", "3128")
    assert pl.proxy_host_port({"http": "proxy.corp:80"}) == ("proxy.corp", "80")
    assert pl.proxy_host_port({}) == ("", "")


class FakeDBus:
    def __init__(self, lookup=None, fail=False):
        self.lookup = lookup
        self.fail = fail
        self.calls = []

    def call(self, bus_name, path, interface, method, signature="", args=(), timeout=30):
        self.calls.append((interface, method, args))
        if self.fail:
            raise RuntimeError("no portal")
        if method == "Lookup":
            return (self.lookup,)
        return ()

    def portal_request(self, interface, method, signature, build_args, timeout=180):
        self.calls.append((interface, method, build_args("tok")))
        return 0, {}


def make_platform(tmp_path, desktop="kde", flatpak=True, dbus=None, environ=None):
    env = {"XDG_CURRENT_DESKTOP": {"kde": "KDE", "gnome": "GNOME", "other": "sway"}[desktop]}
    if flatpak:
        env["FLATPAK_ID"] = pl.APP_ID
    env.update(environ or {})
    paths = pl.LinuxPaths(tmp_path / "data", tmp_path / "cfg" / "config.json", tmp_path / "logs", tmp_path / "pics")
    return pl.LinuxPlatform(paths, launch_command=["bing-wallpaper", "--hidden"], dbus=dbus or FakeDBus(), environ=env)


def test_system_proxies_asks_the_portal_every_time(tmp_path):
    dbus = FakeDBus(lookup=["http://live.corp:3128"])
    p = make_platform(tmp_path, dbus=dbus)
    assert p.system_proxies("https://www.bing.com/") == {"http": "http://live.corp:3128", "https": "http://live.corp:3128"}
    dbus.lookup = ["direct://"]
    assert p.system_proxies("https://www.bing.com/") == {}
    assert len([c for c in dbus.calls if c[1] == "Lookup"]) == 2
    assert p.detect_proxy_values() == ("", "")


def test_system_proxies_falls_back_to_environment_without_portal(tmp_path):
    p = make_platform(tmp_path, desktop="other", flatpak=False, dbus=FakeDBus(fail=True),
                      environ={"https_proxy": "http://env:9"})
    assert p.system_proxies() == {"https": "http://env:9"}


# ---------------------------------------------------------------- wallpaper
def test_kde_wallpaper_goes_through_plasma_shell(tmp_path):
    dbus = FakeDBus()
    p = make_platform(tmp_path, dbus=dbus)
    img = tmp_path / "bing_20261119.jpg"
    img.write_bytes(b"x")
    p.set_wallpaper(img)
    interface, method, args = dbus.calls[-1]
    assert (interface, method) == ("org.kde.PlasmaShell", "evaluateScript")
    assert 'wallpaperPlugin = "org.kde.image"' in args[0] and img.resolve().as_uri() in args[0]


def test_portal_wallpaper_sets_both_when_lock_screen_enabled(tmp_path):
    dbus = FakeDBus()
    p = make_platform(tmp_path, desktop="other", dbus=dbus)

    class Svc:
        config = {"set_lock_screen": True}
    p.bind(Svc())
    img = tmp_path / "bing_20261119.jpg"
    img.write_bytes(b"x")
    p.set_wallpaper(img)
    interface, method, args = dbus.calls[-1]
    assert (interface, method) == ("org.freedesktop.portal.Wallpaper", "SetWallpaperURI")
    assert args[2]["set-on"] == ("s", "both") and args[2]["show-preview"] == ("b", False)
    # the lock screen is then reported as done without a second prompt
    assert p.set_lock_screen(img)[0] is True
    assert len(dbus.calls) == 1


# ---------------------------------------------------------------- lock screen
def test_kde_lock_screen_config_preserves_other_groups():
    text = "[Daemon]\nAutolock=true\n\n[Greeter]\nWallpaperPlugin=org.kde.slideshow\n\n[Greeter][Wallpaper][org.kde.slideshow][General]\nSlidePaths=/x\n"
    out = pl.kde_lock_screen_config(text, "file:///p/bing_20261119.jpg")
    assert "[Daemon]\nAutolock=true" in out
    assert "[Greeter]\nWallpaperPlugin=org.kde.image" in out
    assert "[Greeter][Wallpaper][org.kde.image][General]\nImage=file:///p/bing_20261119.jpg" in out
    assert "SlidePaths=/x" in out
    assert pl.kde_lock_screen_config("", "file:///a.jpg").count("\nImage=file:///a.jpg") == 1


def test_kde_lock_screen_written_atomically(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "cfg"))
    p = make_platform(tmp_path, flatpak=False)
    img = tmp_path / "bing_20261119.jpg"
    img.write_bytes(b"x")
    ok, detail = p.set_lock_screen(img)
    assert ok and detail == "kscreenlockerrc"
    rc = (tmp_path / "cfg" / "kscreenlockerrc").read_text()
    assert f"Image={img.resolve().as_uri()}" in rc
    assert not (tmp_path / "cfg" / "kscreenlockerrc.tmp").exists()


# ---------------------------------------------------------------- autostart
def test_autostart_native_writes_desktop_entry(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "cfg"))
    p = make_platform(tmp_path, flatpak=False)
    assert p.is_startup_enabled() is False
    p.set_startup_enabled(True)
    entry = (tmp_path / "cfg" / "autostart" / f"{pl.APP_ID}.desktop").read_text()
    assert "Exec=bing-wallpaper --hidden" in entry and "Type=Application" in entry
    assert p.is_startup_enabled() is True
    p.set_startup_enabled(False)
    assert p.is_startup_enabled() is False


def test_autostart_flatpak_uses_background_portal(tmp_path):
    dbus = FakeDBus()
    p = make_platform(tmp_path, dbus=dbus)
    p.set_startup_enabled(True)
    interface, method, args = dbus.calls[-1]
    assert (interface, method) == ("org.freedesktop.portal.Background", "RequestBackground")
    assert args[1]["autostart"] == ("b", True) and args[1]["commandline"] == ("as", ["bing-wallpaper", "--hidden"])
    assert p.is_startup_enabled() is True
    p.set_startup_enabled(False)
    assert p.is_startup_enabled() is False


def test_describe(tmp_path):
    assert make_platform(tmp_path).describe() == "KDE Plasma (Flatpak)"
    assert make_platform(tmp_path, desktop="gnome", flatpak=False).describe() == "GNOME (native)"


def test_flatpak_lock_screen_targets_host_config_and_writes_in_place(tmp_path, monkeypatch):
    # Inside the sandbox XDG_CONFIG_HOME is private; Plasma reads ~/.config.
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / ".var" / "app" / "x" / "config"))
    rc = tmp_path / ".config" / "kscreenlockerrc"
    rc.parent.mkdir(parents=True)
    rc.write_text("[Greeter]\nWallpaperPlugin=org.kde.image\n")
    inode = rc.stat().st_ino
    p = make_platform(tmp_path, flatpak=True)
    assert p.kscreenlockerrc_path() == rc
    img = tmp_path / "bing_20261119.jpg"
    img.write_bytes(b"x")
    assert p.set_lock_screen(img) == (True, "kscreenlockerrc")
    assert rc.stat().st_ino == inode  # rewritten in place, no rename
    assert f"Image={img.resolve().as_uri()}" in rc.read_text()
    assert "FillMode=2" in rc.read_text()
    assert not (tmp_path / ".var").exists()


def test_flatpak_lock_screen_needs_existing_file(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    p = make_platform(tmp_path, flatpak=True)
    img = tmp_path / "bing_20261119.jpg"
    img.write_bytes(b"x")
    ok, detail = p.set_lock_screen(img)
    assert ok is False and "System Settings" in detail


def test_native_config_dir_follows_xdg(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "cfg"))
    assert make_platform(tmp_path, flatpak=False).host_config_dir() == tmp_path / "cfg"


# ---------------------------------------------------------------- self-update
def test_flatpak_installation_from_info():
    home = Path("/home/u")
    assert pl.flatpak_installation_from_info("[Instance]\napp-path=/home/u/.local/share/flatpak/app/x/current/active/files\n", home) == "user"
    assert pl.flatpak_installation_from_info("[Instance]\napp-path=/var/lib/flatpak/app/x/current/active/files\n", home) == "system"
    assert pl.flatpak_installation_from_info("garbage [[[", home) == "system"


def test_install_update_bundle_runs_flatpak_on_the_host(tmp_path, monkeypatch):
    p = make_platform(tmp_path, flatpak=True)
    monkeypatch.setattr(p, "flatpak_installation", lambda: "user")
    calls = []

    def runner(cmd, timeout):
        calls.append((cmd, timeout))
        return 0, "Installing...\nInstallation complete.\n"

    ok, detail = p.install_update_bundle(tmp_path / "BingWallpaper.flatpak", runner=runner)
    assert ok and detail.endswith("Installation complete.")
    assert calls[0][0] == ["flatpak", "install", "--user", "-y", "--noninteractive", "--reinstall", str(tmp_path / "BingWallpaper.flatpak")]
    ok, detail = p.install_update_bundle(tmp_path / "b.flatpak", runner=lambda c, t: (1, "error: no remote\n"))
    assert ok is False and "no remote" in detail
    assert make_platform(tmp_path, flatpak=False).supports_self_update is False
