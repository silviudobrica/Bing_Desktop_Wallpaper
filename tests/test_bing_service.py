"""WallpaperService tests with a fake platform and a fake HTTP session.
No network, no desktop, no threads beyond the ones the service starts."""
import io
import json
import sys
import threading
from pathlib import Path

import pytest
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import bing_core as core  # noqa: E402
from bing_service import Platform, WallpaperService  # noqa: E402


def jpeg_bytes(color=(10, 20, 30)):
    buf = io.BytesIO()
    Image.new("RGB", (4, 4), color).save(buf, format="JPEG")
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

    @property
    def text(self):
        return self._body.decode("utf-8", "replace")

    def iter_content(self, chunk_size=8192):
        for i in range(0, len(self._body), chunk_size):
            yield self._body[i:i + chunk_size]


class FakeSession:
    """Routes URLs to canned responses and records every call."""

    def __init__(self):
        self.routes = {}
        self.calls = []

    def route(self, url, response):
        self.routes[url] = response

    def get(self, url, **kwargs):
        self.calls.append((url, kwargs))
        if url not in self.routes:
            raise ConnectionError(f"no route for {url}")
        resp = self.routes[url]
        return resp() if callable(resp) else resp


class FakePlatform(Platform):
    name = "fake"
    supports_lock_screen = True

    def __init__(self):
        super().__init__()
        self.wallpapers = []
        self.lock_screens = []
        self.notifications = []
        self.proxy_lookups = 0
        self.proxies = {}
        self.policy = None
        self.lock_result = (True, "FakeAPI")
        self.detected = ("", "")
        self.startup = False

    def set_wallpaper(self, image_path):
        self.wallpapers.append(Path(image_path))

    def set_lock_screen(self, image_path):
        self.lock_screens.append(Path(image_path))
        return self.lock_result

    def lock_screen_policy(self):
        return self.policy

    def system_proxies(self, url=None):
        self.proxy_lookups += 1
        return dict(self.proxies)

    def detect_proxy_values(self):
        return self.detected

    def notify(self, title, message):
        self.notifications.append((title, message))

    def is_startup_enabled(self):
        return self.startup

    def set_startup_enabled(self, enabled):
        self.startup = bool(enabled)


class RecordingService(WallpaperService):
    def __init__(self, *args, **kwargs):
        self.events = []
        super().__init__(*args, **kwargs)

    def on_state_changed(self):
        self.events.append("state")

    def on_wallpaper_applied(self, image_path):
        self.events.append(("applied", image_path.name))

    def on_images_changed(self):
        self.events.append("images")


def bing_payload(dates):
    return {"images": [
        {"startdate": d, "urlbase": f"/th?id=OHR.Pic{d}", "title": f"Pic {d}", "copyright": "c"}
        for d in dates
    ]}


@pytest.fixture
def svc(tmp_path):
    platform = FakePlatform()
    service = RecordingService(platform, tmp_path / "cfg" / "config.json", tmp_path / "Bing", "TestApp", "1.0.0")
    service.session = FakeSession()
    service.config["backfill_days"] = 2
    return service


def route_images(service, dates, resolution="_UHD.jpg"):
    api_url = service.bing_api_url(1 + service.get_int_setting("backfill_days", 7, 0, 7))
    service.session.route(api_url, FakeResponse(payload=bing_payload(dates)))
    for d in dates:
        url = f"{core.BING_IMAGE_HOST}/th?id=OHR.Pic{d}{resolution}"
        service.session.route(url, FakeResponse(body=jpeg_bytes(), headers={"Content-Type": "image/jpeg"}))


# ---------------------------------------------------------------- check loop
def test_first_check_sets_wallpaper_and_backfills(svc):
    route_images(svc, ["20261119", "20261118", "20261117"])
    svc.check_and_update(force=True)

    assert svc.platform.wallpapers == [svc.image_dir / "bing_20261119.jpg"]
    assert svc.current_image_path == svc.image_dir / "bing_20261119.jpg"
    assert sorted(p.name for p in svc.image_dir.glob("bing_*.jpg")) == [
        "bing_20261117.jpg", "bing_20261118.jpg", "bing_20261119.jpg"]
    assert svc.last_success > 0 and svc.last_error == "None" and svc.consecutive_failures == 0
    assert svc.last_status_message == "Wallpaper updated: bing_20261119.jpg"
    assert ("applied", "bing_20261119.jpg") in svc.events and "images" in svc.events
    assert svc.platform.notifications == [("TestApp", "Wallpaper updated: bing_20261119.jpg")]
    assert svc.list_images_newest_first()[0].name == "bing_20261119.jpg"


def test_unforced_check_with_same_image_does_not_reapply(svc):
    route_images(svc, ["20261119"])
    svc.check_and_update(force=True)
    svc.check_and_update(force=False)
    assert len(svc.platform.wallpapers) == 1
    assert svc.last_status_message == "Wallpaper available: bing_20261119.jpg"


def test_resolution_fallback_on_404(svc):
    api_url = svc.bing_api_url(3)
    svc.session.route(api_url, FakeResponse(payload=bing_payload(["20261119"])))
    svc.session.route(f"{core.BING_IMAGE_HOST}/th?id=OHR.Pic20261119_UHD.jpg", FakeResponse(status=404))
    svc.session.route(f"{core.BING_IMAGE_HOST}/th?id=OHR.Pic20261119_1920x1080.jpg",
                      FakeResponse(body=jpeg_bytes(), headers={"Content-Type": "image/jpeg"}))
    svc.check_and_update(force=True)
    assert svc.current_image_path.name == "bing_20261119.jpg"


def test_failed_check_backs_off_and_stays_quiet(svc):
    svc.check_interval = 3600
    svc.check_and_update(force=True)  # no routes -> API failure
    assert svc.consecutive_failures == 1
    assert svc.last_error == "No image info returned from Bing API"
    assert svc.last_status_message == "Last operation failed"
    assert svc._next_check_at > 0
    assert svc.platform.notifications == []  # quiet for the first failures
    for _ in range(core.QUIET_FAILURE_COUNT - 1):
        svc.check_and_update(force=True)
    assert svc.platform.notifications[-1][1].startswith("Operation failed")


def test_corrupt_download_is_rejected(svc):
    api_url = svc.bing_api_url(3)
    svc.session.route(api_url, FakeResponse(payload=bing_payload(["20261119"])))
    for res in core.IMAGE_RESOLUTION_SUFFIXES:
        svc.session.route(f"{core.BING_IMAGE_HOST}/th?id=OHR.Pic20261119{res}",
                          FakeResponse(body=b"not an image", headers={"Content-Type": "image/jpeg"}))
    svc.check_and_update(force=True)
    assert svc.last_error == "Image download failed"
    assert list(svc.image_dir.glob("*")) == []  # temp file cleaned up


# ---------------------------------------------------------------- proxies
def test_explicit_proxy_wins_over_system_proxy(svc):
    svc.platform.proxies = {"https": "http://sys:3128"}
    assert svc.get_proxy_dict("https://x") == {"https": "http://sys:3128"}
    svc.config["proxy_url"] = "corp"
    svc.config["proxy_port"] = "8080"
    assert svc.get_proxy_dict("https://x") == core.build_proxy_dict("corp", "8080")


def test_system_proxy_is_looked_up_fresh_for_every_request(svc):
    route_images(svc, ["20261119", "20261118", "20261117"])
    svc.check_and_update(force=True)
    requests_made = len(svc.session.calls)
    assert requests_made >= 4
    assert svc.platform.proxy_lookups == requests_made


def test_detected_proxy_is_persisted_only_when_platform_says_so(tmp_path):
    platform = FakePlatform()
    platform.detected = ("pac-proxy", "3128")
    service = WallpaperService(platform, tmp_path / "config.json", tmp_path / "img", "T", "1")
    assert "proxy_url" not in service.config

    platform.autosave_detected_proxy = True
    service = WallpaperService(platform, tmp_path / "config.json", tmp_path / "img", "T", "1")
    assert service.config["proxy_url"] == "pac-proxy"
    assert json.loads((tmp_path / "config.json").read_text())["proxy_port"] == "3128"


# ---------------------------------------------------------------- lock screen
def test_lock_screen_applied_after_wallpaper_when_enabled(svc):
    svc.config["set_lock_screen"] = True
    image = svc.image_dir / "bing_20261119.jpg"
    image.write_bytes(jpeg_bytes())
    svc._set_lock_screen_blocking(image)
    assert svc.platform.lock_screens == [image]
    assert svc.last_error == "None"


def test_lock_screen_policy_blocks_and_reports(svc):
    svc.platform.policy = "LockScreenImage policy (x)"
    image = svc.image_dir / "bing_20261119.jpg"
    image.write_bytes(jpeg_bytes())
    assert svc._set_lock_screen_blocking(image) is False
    assert svc.platform.lock_screens == []
    assert "managed by your organization" in svc.last_error


def test_lock_screen_failure_detail_is_recorded(svc):
    svc.platform.lock_result = (False, "portal refused")
    image = svc.image_dir / "bing_20261119.jpg"
    image.write_bytes(jpeg_bytes())
    svc._set_lock_screen_blocking(image)
    assert svc.last_error == "Lock screen update failed: portal refused"


def test_apply_lock_screen_from_main_thread_uses_worker(svc):
    image = svc.image_dir / "bing_20261119.jpg"
    image.write_bytes(jpeg_bytes())
    before = threading.active_count()
    svc.apply_lock_screen(image)
    for t in threading.enumerate():
        if t is not threading.main_thread():
            t.join(timeout=5)
    assert svc.platform.lock_screens == [image]
    assert threading.active_count() <= before + 1


# ---------------------------------------------------------------- housekeeping
def test_retention_keeps_current_image(svc):
    for d in ("20200101", "20200102"):
        (svc.image_dir / f"bing_{d}.jpg").write_bytes(jpeg_bytes())
    svc.current_image_path = svc.image_dir / "bing_20200102.jpg"
    svc.config["keep_images_days"] = 30
    assert svc.cleanup_old_images() == 1
    assert [p.name for p in svc.image_dir.glob("*.jpg")] == ["bing_20200102.jpg"]


def test_config_roundtrip_is_atomic(svc):
    svc.config["market"] = "de-DE"
    svc.save_config()
    assert not svc.config_file.with_name("config.json.tmp").exists()
    assert svc.load_config()["market"] == "de-DE"
    assert svc.get_market() == "de-DE"


def test_set_interval_schedules_immediate_check(svc):
    svc._next_check_at = 10 ** 12
    svc.set_interval(60)
    assert svc.check_interval == 3600 and svc._next_check_at == 0
    assert svc.load_config()["check_interval_minutes"] == 60
    svc.set_interval(0)
    assert svc.check_interval == 0


def test_platform_errors_surface_as_recorded_errors(svc):
    class Broken(FakePlatform):
        def set_startup_enabled(self, enabled):
            raise OSError("no shortcut")
    svc.platform = Broken()
    svc.set_startup_enabled(True)
    assert svc.last_error == "Startup toggle failed: no shortcut"
    assert svc.is_startup_enabled() is False


# ---------------------------------------------------------------- release downloads
def test_download_release_asset_verifies_sha256(svc, tmp_path):
    import hashlib
    body = b"flatpak-bundle-bytes"
    svc.session.route("https://x/BingWallpaper.flatpak", FakeResponse(body=body))
    target = svc.download_release_asset("https://x/BingWallpaper.flatpak", "BingWallpaper.flatpak",
                                        hashlib.sha256(body).hexdigest(), tmp_path / "updates")
    assert target == tmp_path / "updates" / "BingWallpaper.flatpak" and target.read_bytes() == body
    assert not (tmp_path / "updates" / "BingWallpaper.flatpak.download").exists()


def test_download_release_asset_discards_on_mismatch(svc, tmp_path):
    svc.session.route("https://x/b.flatpak", FakeResponse(body=b"tampered"))
    with pytest.raises(RuntimeError, match="checksum mismatch"):
        svc.download_release_asset("https://x/b.flatpak", "b.flatpak", "0" * 64, tmp_path / "updates")
    assert list((tmp_path / "updates").glob("*")) == []
