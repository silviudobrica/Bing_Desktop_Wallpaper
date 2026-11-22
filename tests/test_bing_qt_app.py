"""Offscreen tests for the PySide6 front end. The real platform_linux module
is used with a fake D-Bus and a fixed environment, so nothing touches the
desktop."""
import io
import os
import sys
import time
import types
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

pytest.importorskip("PySide6")
from PIL import Image  # noqa: E402
from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

import bing_core as core  # noqa: E402


# ---------------------------------------------------------------- real platform, fake D-Bus
import platform_linux as pl  # noqa: E402
import bing_qt_app as qt_app  # noqa: E402


class FakeDBus:
    """Stands in for platform_linux.DBus so no session bus is touched."""

    def __init__(self):
        self.calls = []

    def call(self, *args, **kwargs):
        self.calls.append(("call", args, kwargs))
        return ()

    def portal_request(self, *args, **kwargs):
        self.calls.append(("portal_request", args, kwargs))
        return 0, {}


def make_platform(paths):
    """Real LinuxPlatform (KDE, Flatpak) with the OS-touching methods
    replaced by recorders on the instance."""
    platform = pl.LinuxPlatform(
        paths, launch_command=["bing-wallpaper", "--hidden"], dbus=FakeDBus(),
        environ={"XDG_CURRENT_DESKTOP": "KDE", "FLATPAK_ID": pl.APP_ID},
    )
    platform.wallpapers = []
    platform.notifications = []
    platform.opened = []
    platform.startup = False
    platform.set_wallpaper = lambda path: platform.wallpapers.append(Path(path))
    platform.set_lock_screen = lambda path: (True, "fake")
    platform.notify = lambda title, message: platform.notifications.append((title, message))
    platform.open_url = lambda url: platform.opened.append(url)
    platform.open_path = lambda path: platform.opened.append(str(path))
    platform.is_startup_enabled = lambda: platform.startup
    platform.set_startup_enabled = lambda enabled: setattr(platform, "startup", bool(enabled))
    return platform


# ---------------------------------------------------------------- helpers
def jpeg_bytes(color):
    buf = io.BytesIO()
    Image.new("RGB", (16, 9), color).save(buf, format="JPEG")
    return buf.getvalue()


class FakeResponse:
    def __init__(self, status=200, body=b"", headers=None, payload=None):
        self.status_code = status
        self.headers = headers or {}
        self._body = body
        self._payload = payload

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def json(self):
        return self._payload

    def iter_content(self, chunk_size=8192):
        yield self._body


class FakeSession:
    def __init__(self):
        self.routes = {}

    def get(self, url, **kwargs):
        if url not in self.routes:
            raise ConnectionError(url)
        return self.routes[url]


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture
def env(tmp_path, qapp):
    paths = pl.LinuxPaths(data_dir=tmp_path, config_file=tmp_path / "cfg" / "config.json",
                          log_dir=tmp_path / "logs", image_dir=tmp_path / "Bing")
    paths.image_dir.mkdir(parents=True)
    platform = make_platform(paths)
    service = qt_app.QtWallpaperService(platform, paths)
    service.session = FakeSession()
    return types.SimpleNamespace(paths=paths, platform=platform, service=service, app=qapp)


def make_images(image_dir, days):
    paths = []
    for i in range(days):
        name = core.image_filename(f"202611{19 - i:02d}")
        p = image_dir / name
        p.write_bytes(jpeg_bytes((i * 20 % 255, 50, 100)))
        paths.append(p)
    return paths


def wait_until(predicate, timeout=5.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        QApplication.processEvents()
        if predicate():
            return True
        time.sleep(0.02)
    return predicate()


# ---------------------------------------------------------------- tests
def test_window_builds_and_status_rows_populate(env):
    api_url = env.service.bing_api_url(1 + env.service.get_int_setting("backfill_days", 7, 0, 7))
    env.service.session.routes[api_url] = FakeResponse(payload={"images": [
        {"startdate": "20261119", "urlbase": "/th?id=OHR.X", "title": "t", "copyright": "c"}]})
    env.service.session.routes[f"{core.BING_IMAGE_HOST}/th?id=OHR.X_UHD.jpg"] = FakeResponse(
        body=jpeg_bytes((1, 2, 3)), headers={"Content-Type": "image/jpeg"})

    window = qt_app.MainWindow(env.service, env.platform, env.paths, tray_available=True)
    window.resize(840, 760)
    window.show()
    QApplication.processEvents()
    assert window.windowTitle() == f"Bing Wallpaper v{qt_app.VERSION}"
    assert window.status_labels["status"].text() == "Waiting for first check"
    assert window.status_labels["desktop"].text() == "KDE Plasma (Flatpak)"
    assert window.status_labels["interval"].text() == "720 minutes"

    env.service.check_and_update(force=True)
    QApplication.processEvents()
    assert env.platform.wallpapers == [env.paths.image_dir / "bing_20261119.jpg"]
    assert window.status_labels["status"].text() == "Wallpaper updated: bing_20261119.jpg"
    assert window.status_labels["current_image"].text() == "bing_20261119.jpg"
    assert window.status_labels["last_error"].text() == "None"
    assert window.status_labels["last_success"].text() != "Never"
    assert window.preview_name.text() == "bing_20261119.jpg"
    assert not window.preview_label.pixmap().isNull()
    assert window.gallery_counter.text() == "1/1"
    window.close()


def test_gallery_pages_newest_first(env):
    images = make_images(env.paths.image_dir, 8)
    window = qt_app.MainWindow(env.service, env.platform, env.paths, tray_available=True)
    window.show()
    QApplication.processEvents()

    assert [p.name for p in window.images] == [p.name for p in images]  # newest first
    assert window.gallery_counter.text() == "1-5/8"
    assert window.thumb_slots[0][1].text() == "Nov 19, 2026"
    assert window.thumb_slots[4][1].text() == "Nov 15, 2026"
    assert window.gallery_scroll.maximum() == 3
    assert all(not thumb.pixmap().isNull() for thumb, _ in window.thumb_slots)

    window.gallery_scroll.setValue(3)
    QApplication.processEvents()
    assert window.gallery_counter.text() == "4-8/8"
    assert window.thumb_slots[0][1].text() == "Nov 16, 2026"
    assert window.thumb_slots[4][1].text() == "Nov 12, 2026"

    # A new image on disk refreshes the strip and keeps the page in range.
    (env.paths.image_dir / "bing_20261120.jpg").write_bytes(jpeg_bytes((9, 9, 9)))
    env.service.signals.images_changed.emit()
    QApplication.processEvents()
    assert window.images[0].name == "bing_20261120.jpg"
    assert window.gallery_counter.text() == "4-8/9"
    window.close()


def test_thumbnail_click_applies_wallpaper(env):
    images = make_images(env.paths.image_dir, 6)
    window = qt_app.MainWindow(env.service, env.platform, env.paths, tray_available=True)
    window.show()
    QApplication.processEvents()

    window.thumb_slots[2][0].clicked.emit()
    assert wait_until(lambda: env.platform.wallpapers == [images[2]])
    assert wait_until(lambda: window.status_labels["status"].text() == f"Wallpaper updated: {images[2].name}")
    assert window.preview_name.text() == images[2].name
    window.close()


def test_gallery_without_images_and_small_pages(env):
    window = qt_app.MainWindow(env.service, env.platform, env.paths, tray_available=True)
    assert window.gallery_counter.text() == "0/0"
    assert window.thumb_slots[0][0].text() == "No image"
    make_images(env.paths.image_dir, 3)
    window.reload_gallery()
    assert window.gallery_counter.text() == "3/3"
    assert window.gallery_scroll.maximum() == 0


def test_settings_dialog_round_trips_config(env):
    make_images(env.paths.image_dir, 1)
    dialog = qt_app.SettingsDialog(env.service)
    assert dialog.interval_combo.currentText() == "12 Hours"
    dialog.interval_combo.setCurrentText("1 Hour")
    dialog.custom_interval.setText("")
    dialog.proxy_host.setText("proxy.corp")
    dialog.proxy_port.setText("8080")
    dialog.ca_bundle.setText("/etc/ca.pem")
    dialog.market.setCurrentText("de-DE")
    dialog.keep_days.setText("30")
    dialog.backfill_days.setText("3")
    dialog.lock_screen.setChecked(True)
    dialog.startup.setChecked(True)
    dialog.save()
    assert dialog.result() == qt_app.QDialog.Accepted

    saved = env.service.load_config()
    assert saved["check_interval_minutes"] == 60
    assert saved["proxy_url"] == "proxy.corp" and saved["proxy_port"] == "8080"
    assert saved["ca_bundle_path"] == "/etc/ca.pem"
    assert saved["market"] == "de-DE"
    assert saved["keep_images_days"] == 30 and saved["backfill_days"] == 3
    assert saved["set_lock_screen"] is True
    assert env.platform.startup is True
    assert env.service.last_status_message == "Settings saved"
    assert env.service.check_interval == 3600

    reopened = qt_app.SettingsDialog(env.service)
    assert reopened.interval_combo.currentText() == "1 Hour"
    assert reopened.market.currentText() == "de-DE"
    assert reopened.lock_screen.isChecked() and reopened.startup.isChecked()
    reopened.custom_interval.setText("90")
    reopened.save()
    assert env.service.load_config()["check_interval_minutes"] == 90
    third = qt_app.SettingsDialog(env.service)
    assert third.custom_interval.text() == "90"


def test_settings_dialog_has_check_for_updates(env, monkeypatch):
    checks = []
    monkeypatch.setattr(env.service, "check_for_updates", lambda manual=False: checks.append(manual))
    dialog = qt_app.SettingsDialog(env.service)
    button = next(b for b in dialog.findChildren(qt_app.QPushButton) if b.text() == "Check for Updates")
    button.click()
    for _ in range(50):
        if checks:
            break
        time.sleep(0.02)
    assert checks == [True]
    assert not button.isDefault()  # Enter still saves


def test_settings_dialog_rejects_bad_input(env, monkeypatch):
    errors = []
    monkeypatch.setattr(qt_app.QMessageBox, "critical", lambda *a: errors.append(a[2]))
    dialog = qt_app.SettingsDialog(env.service)
    dialog.market.setCurrentText("germany")
    dialog.save()
    assert errors and "Market must look like en-US" in errors[-1]
    dialog.market.setCurrentText("en-US")
    dialog.backfill_days.setText("12")
    dialog.save()
    assert "Backfill days must be between 0 and" in errors[-1]
    dialog.backfill_days.setText("2")
    dialog.custom_interval.setText("-5")
    dialog.save()
    assert "Custom interval must be positive" in errors[-1]
    assert dialog.result() != qt_app.QDialog.Accepted


def test_escape_hides_window_and_close_hides_with_tray(env):
    window = qt_app.MainWindow(env.service, env.platform, env.paths, tray_available=True)
    window.show()
    QApplication.processEvents()
    assert window.isVisible()
    QTest.keyClick(window, Qt.Key_Escape)
    QApplication.processEvents()
    assert not window.isVisible()
    window.show()
    window.close()
    QApplication.processEvents()
    assert not window.isVisible()


def test_update_check_reports_through_signal(env, monkeypatch):
    results = []
    env.service.signals.update_result.connect(lambda info, manual: results.append((info, manual)))
    monkeypatch.setattr(env.service, "fetch_latest_release",
                        lambda: {"tag_name": "v99.0.0", "html_url": "https://example/rel"})
    env.service.check_for_updates(manual=True)
    QApplication.processEvents()
    assert results == [({"version": "99.0.0", "html_url": "https://example/rel", "newer": True, "asset": None}, True)]
    assert env.service.last_status_message == "Update available: v99.0.0"
    assert env.service.last_update_check > 0

    monkeypatch.setattr(env.service, "fetch_latest_release", lambda: (_ for _ in ()).throw(RuntimeError("down")))
    env.service.check_for_updates(manual=True)
    QApplication.processEvents()
    assert results[-1] == ({}, True)


def test_update_dialog_opens_release_page(env, monkeypatch):
    controller = qt_app.BingQtApp(env.app, env.service, env.platform, env.paths)
    monkeypatch.setattr(qt_app.QMessageBox, "question", lambda *a, **k: qt_app.QMessageBox.Yes)
    controller.on_update_result({"version": "99.0.0", "html_url": "https://example/rel", "newer": True}, True)
    assert env.platform.opened == ["https://example/rel"]
    infos = []
    monkeypatch.setattr(qt_app.QMessageBox, "information", lambda *a: infos.append(a[2]))
    controller.on_update_result({"version": "0.0.1", "html_url": "u", "newer": False}, True)
    assert infos[-1] == f"You are up to date (v{qt_app.VERSION})."
    controller.on_update_result({"version": "99.0.0", "html_url": "u", "newer": True}, False)
    assert env.platform.opened == ["https://example/rel"]  # silent when automatic
    controller.window.close()


def test_in_app_update_downloads_installs_and_restarts(env, monkeypatch):
    """Flatpak with a bundle asset: Yes -> download (verified) -> host install -> restart + quit."""
    controller = qt_app.BingQtApp(env.app, env.service, env.platform, env.paths)
    monkeypatch.setattr(type(env.platform), "supports_self_update", property(lambda self: True))
    monkeypatch.setattr(qt_app.QMessageBox, "question", lambda *a, **k: qt_app.QMessageBox.Yes)
    infos, installed, restarted, quit_calls = [], [], [], []
    monkeypatch.setattr(qt_app.QMessageBox, "information", lambda *a: infos.append(a[2]))
    monkeypatch.setattr(env.service, "download_release_asset",
                        lambda url, name, sha, target_dir, timeout=0: (Path(target_dir).mkdir(parents=True, exist_ok=True) or Path(target_dir) / name))
    monkeypatch.setattr(env.platform, "install_update_bundle", lambda bundle: installed.append(bundle) or (True, "Installation complete."))
    monkeypatch.setattr(env.platform, "restart_app", lambda: restarted.append(True))
    monkeypatch.setattr(env.app, "quit", lambda: quit_calls.append(True))
    info = {"version": "99.0.0", "html_url": "u", "newer": True,
            "asset": {"name": "BingWallpaper.flatpak", "url": "https://x/BingWallpaper.flatpak", "sha256": None}}
    controller.on_update_result(info, True)
    for _ in range(50):
        QApplication.processEvents()
        if quit_calls:
            break
        time.sleep(0.05)
    assert installed == [env.paths.data_dir / "updates" / "BingWallpaper.flatpak"]
    assert restarted == [True] and quit_calls == [True]
    assert infos[-1] == "Update v99.0.0 installed. The app will now restart."
    assert env.service.last_status_message.startswith("Update v99.0.0 installed")
    controller.window.close()


def test_in_app_update_failure_offers_release_page(env, monkeypatch):
    controller = qt_app.BingQtApp(env.app, env.service, env.platform, env.paths)
    monkeypatch.setattr(type(env.platform), "supports_self_update", property(lambda self: True))
    monkeypatch.setattr(qt_app.QMessageBox, "question", lambda *a, **k: qt_app.QMessageBox.Yes)
    errors = []
    monkeypatch.setattr(qt_app.QMessageBox, "critical", lambda *a, **k: errors.append(a[2]) or qt_app.QMessageBox.Yes)
    monkeypatch.setattr(env.service, "download_release_asset", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("checksum mismatch")))
    info = {"version": "99.0.0", "html_url": "https://example/rel", "newer": True,
            "asset": {"name": "BingWallpaper.flatpak", "url": "https://x/b", "sha256": "0" * 64}}
    controller.on_update_result(info, True)
    for _ in range(50):
        QApplication.processEvents()
        if errors:
            break
        time.sleep(0.05)
    assert "checksum mismatch" in errors[-1]
    assert env.platform.opened[-1] == "https://example/rel"
    assert env.service.last_error == "Update install failed: checksum mismatch"
    controller.window.close()


def test_notification_falls_back_to_platform_without_tray(env):
    controller = qt_app.BingQtApp(env.app, env.service, env.platform, env.paths)
    controller.tray = None
    controller.on_notification("Bing Wallpaper", "hello")
    assert env.platform.notifications == [("Bing Wallpaper", "hello")]
    controller.window.close()


class FakeTray:
    def __init__(self):
        self.messages = []

    def notify(self, title, message):
        self.messages.append((title, message))
        return True


def test_notification_in_flatpak_uses_portal_not_tray(env):
    # The sandbox cannot reach org.freedesktop.Notifications, which is what
    # the tray's showMessage() uses; it would be dropped silently.
    controller = qt_app.BingQtApp(env.app, env.service, env.platform, env.paths)
    controller.tray = FakeTray()
    controller.on_notification("Bing Wallpaper", "hello")
    assert env.platform.notifications == [("Bing Wallpaper", "hello")]
    assert controller.tray.messages == []
    controller.tray = None
    controller.window.close()


def test_notification_outside_flatpak_prefers_tray(env):
    controller = qt_app.BingQtApp(env.app, env.service, env.platform, env.paths)
    env.platform.in_flatpak = False
    controller.tray = FakeTray()
    controller.on_notification("Bing Wallpaper", "hello")
    assert controller.tray.messages == [("Bing Wallpaper", "hello")]
    assert env.platform.notifications == []
    controller.tray = None
    controller.window.close()


def test_launch_command_and_display_date():
    assert qt_app.launch_command(True) == ["bing-wallpaper", "--hidden"]
    assert qt_app.launch_command(False)[-1] == "--hidden"
    assert qt_app.display_date(Path("bing_20261112.jpg")) == "Nov 12, 2026"
    assert qt_app.display_date(Path("bing_custom.jpg")) == "bing_custom"
