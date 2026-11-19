import datetime
import hashlib
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import bing_core as core  # noqa: E402

TODAY = datetime.date(2026, 11, 19)


# ---------------------------------------------------------------- versions
@pytest.mark.parametrize("raw, expected", [
    ("1.3.8", (1, 3, 8)),
    ("v1.3.8", (1, 3, 8)),
    ("1.3", (1, 3, 0)),
    ("1.3.8-beta2", (1, 3, 8)),
    ("", (0, 0, 0)),
    (None, (0, 0, 0)),
])
def test_version_tuple(raw, expected):
    assert core.version_tuple(raw) == expected


@pytest.mark.parametrize("candidate, current, newer", [
    ("1.3.9", "1.3.8", True),
    ("v1.4.0", "1.3.8", True),
    ("1.3.8", "1.3.8", False),
    ("1.3.7", "1.3.8", False),
    ("2.0", "1.9.9", True),
])
def test_is_newer_version(candidate, current, newer):
    assert core.is_newer_version(candidate, current) is newer


# ---------------------------------------------------------------- settings
@pytest.mark.parametrize("value, expected", [
    ("5", 5), (5, 5), (" 7 ", 7), ("abc", 3), (None, 3), ("-1", 0), ("999", 10),
])
def test_clamp_int(value, expected):
    assert core.clamp_int(value, default=3, minimum=0, maximum=10) == expected


@pytest.mark.parametrize("market, expected", [
    ("fr-FR", "fr-FR"), (" de-DE ", "de-DE"), ("", "en-US"), (None, "en-US"),
    ("garbage", "en-US"), ("en_US", "en-US"), ("en-USA", "en-US"),
])
def test_normalize_market(market, expected):
    assert core.normalize_market(market) == expected


def test_bing_api_url_clamps_count_and_uses_market():
    assert core.bing_api_url("fr-FR", 8) == core.BING_API_BASE + "?format=js&idx=0&n=8&mkt=fr-FR"
    assert "n=8&" in core.bing_api_url("en-US", 50)
    assert "n=1&" in core.bing_api_url("en-US", 0)
    assert "mkt=en-US" in core.bing_api_url("nope", 1)


# ---------------------------------------------------------------- bing payload
def test_parse_bing_images_builds_candidates_newest_first():
    data = {"images": [
        {"startdate": "20261119", "urlbase": "/th?id=OHR.A", "title": "A"},
        {"startdate": "20261118", "urlbase": "/th?id=OHR.B", "copyright": "B (c)"},
        {"startdate": "20261117"},                       # no urlbase -> skipped
        {"urlbase": "/th?id=OHR.D"},                     # no date -> skipped
    ]}
    entries = core.parse_bing_images(data)
    assert [e["date"] for e in entries] == ["20261119", "20261118"]
    assert entries[0]["urls"] == [
        "https://www.bing.com/th?id=OHR.A_UHD.jpg",
        "https://www.bing.com/th?id=OHR.A_1920x1080.jpg",
    ]
    assert entries[0]["title"] == "A"
    assert entries[1]["title"] == "B (c)"


def test_parse_bing_images_tolerates_empty_payloads():
    assert core.parse_bing_images({}) == []
    assert core.parse_bing_images(None) == []
    assert core.parse_bing_images({"images": None}) == []


def test_image_filename_and_date_roundtrip():
    name = core.image_filename("20261119")
    assert name == "bing_20261119.jpg"
    assert core.parse_image_date(Path("C:/x") / name) == TODAY
    assert core.parse_image_date("bing_notadate.jpg") is None
    assert core.parse_bing_date("2026-11-19") is None


# ---------------------------------------------------------------- proxy
@pytest.mark.parametrize("raw, expected", [
    ("proxy.corp:9090", ("proxy.corp", "9090")),
    ("http://proxy.corp:8080/", ("proxy.corp", "8080")),
    ("https://proxy.corp", ("proxy.corp", "80")),
    ("proxy.corp", ("proxy.corp", "80")),
    ("", (None, None)),
    (None, (None, None)),
])
def test_parse_proxy_string(raw, expected):
    assert core.parse_proxy_string(raw) == expected


def test_build_proxy_dict():
    assert core.build_proxy_dict("h", "1") == {"http": "http://h:1", "https": "http://h:1"}
    assert core.build_proxy_dict("h", "") is None
    assert core.build_proxy_dict("", "1") is None
    assert core.build_proxy_dict(None, None) is None


def test_extract_pac_proxy():
    pac = 'function FindProxyForURL(u,h){ if (x) return "DIRECT"; return "PROXY proxyshwg.onelux.net:9090; DIRECT"; }'
    assert core.extract_pac_proxy(pac) == "proxyshwg.onelux.net:9090"
    assert core.extract_pac_proxy('return "DIRECT";') is None
    assert core.extract_pac_proxy(None) is None


# ---------------------------------------------------------------- failures
def test_retry_delay_backoff_sequence_and_cap():
    minutes = [core.retry_delay_seconds(n, 24 * 3600) // 60 for n in range(1, 8)]
    assert minutes == [2, 5, 15, 30, 60, 60, 60]
    assert core.retry_delay_seconds(5, 600) == 600          # capped at the interval
    assert core.retry_delay_seconds(0, 3600) == 120         # defensive: 0 failures
    assert core.retry_delay_seconds(3, 0) == 15 * 60        # disabled interval: raw backoff


def test_should_notify_failure_is_quiet_then_once_then_on_change():
    assert core.should_notify_failure(1, "e", "None") is False
    assert core.should_notify_failure(2, "e", "e") is False
    assert core.should_notify_failure(3, "e", "e") is True
    assert core.should_notify_failure(4, "e", "e") is False
    assert core.should_notify_failure(4, "different", "e") is True


# ---------------------------------------------------------------- retention
def _img(days_ago):
    return Path("C:/Pictures/Bing") / core.image_filename((TODAY - datetime.timedelta(days=days_ago)).strftime("%Y%m%d"))


def test_select_old_images_respects_cutoff_and_current():
    paths = [_img(0), _img(4), _img(5), _img(30), Path("C:/Pictures/Bing/bing_odd.jpg")]
    assert core.select_old_images(paths, keep_days=0, today=TODAY) == []
    assert core.select_old_images(paths, keep_days=5, today=TODAY) == [_img(5), _img(30)]
    assert core.select_old_images(paths, keep_days=5, current_path=_img(30), today=TODAY) == [_img(5)]


def test_within_retention():
    assert core.within_retention("20261119", 0, TODAY)
    assert core.within_retention("20261115", 5, TODAY)
    assert not core.within_retention("20261114", 5, TODAY)
    assert not core.within_retention("bad", 5, TODAY)


# ---------------------------------------------------------------- updater
HEX = "54253280287dfe145ddc23c4e9f0ef9e147e63a23712c0e3c9128c02fbd53bcd"


def test_parse_asset_digest():
    assert core.parse_asset_digest("sha256:" + HEX.upper()) == HEX
    assert core.parse_asset_digest("md5:abc") is None
    assert core.parse_asset_digest("sha256:short") is None
    assert core.parse_asset_digest(None) is None


def test_find_sha256_in_text_prefers_asset_line():
    other = "a" * 64
    body = f"Checksums\nOther.exe  {other}\nInstallBingWallpaper.exe  {HEX}\n"
    assert core.find_sha256_in_text(body, "InstallBingWallpaper.exe") == HEX
    assert core.find_sha256_in_text(body, "Missing.exe") == other   # falls back to first hash
    assert core.find_sha256_in_text("no hashes here", "x") is None


def test_select_installer_asset_uses_digest_then_body():
    assets = [
        {"name": "Source.zip", "browser_download_url": "u0"},
        {"name": "InstallBingWallpaper-1.3.8.exe", "browser_download_url": "u1", "digest": "sha256:" + HEX},
    ]
    picked = core.select_installer_asset(assets)
    assert picked == {"name": "InstallBingWallpaper-1.3.8.exe", "url": "u1", "sha256": HEX}

    assets[1].pop("digest")
    assert core.select_installer_asset(assets)["sha256"] is None
    assert core.select_installer_asset(assets, f"InstallBingWallpaper-1.3.8.exe {HEX}")["sha256"] == HEX
    assert core.select_installer_asset([{"name": "readme.txt"}]) is None
    assert core.select_installer_asset(None) is None


def test_sha256_of_file(tmp_path):
    f = tmp_path / "blob.bin"
    data = b"x" * (3 * 1024 * 1024 + 17)
    f.write_bytes(data)
    assert core.sha256_of_file(f) == hashlib.sha256(data).hexdigest()


# ---------------------------------------------------------------- lock screen
def test_lock_screen_copy_name_is_unique_per_timestamp():
    a = core.lock_screen_copy_name(Path("Pictures") / "Bing" / "bing_2026-11-19.jpg", 1700000000.7)
    b = core.lock_screen_copy_name("bing_2026-11-19.jpg", 1700000001)
    assert a == "lockscreen_bing_2026-11-19_1700000000.jpg"
    assert b == "lockscreen_bing_2026-11-19_1700000001.jpg"
    assert core.lock_screen_copy_name("noext", 5) == "lockscreen_noext_5.jpg"


def test_select_stale_lock_screen_copies_keeps_current():
    paths = [Path("d/lockscreen_a_1.jpg"), Path("d/lockscreen_a_2.jpg"), Path("d/lockscreen_b_3.jpg")]
    assert core.select_stale_lock_screen_copies(paths, Path("d/lockscreen_a_2.jpg")) == [paths[0], paths[2]]
    assert core.select_stale_lock_screen_copies(paths) == paths
    assert core.select_stale_lock_screen_copies([], Path("d/x.jpg")) == []


def test_lock_screen_picture_mode_values_turn_spotlight_off():
    names = {name: value for _, name, value in core.LOCK_SCREEN_PICTURE_MODE_VALUES}
    assert names["RotatingLockScreenEnabled"] == 0
    assert names["SlideshowEnabled"] == 0
    assert all(subkey.startswith("Software\\Microsoft\\Windows\\CurrentVersion") for subkey, _, _ in core.LOCK_SCREEN_PICTURE_MODE_VALUES)


# ---------------------------------------------------------------- gallery order
def test_sort_images_newest_first_uses_date_in_name_not_mtime():
    today, d1, d2, d3 = (_img(0), _img(1), _img(2), _img(3))
    # Backfill wrote the older days after today's image, so their mtimes are newer.
    mtimes = {today: 100, d3: 200, d2: 300, d1: 400}
    ordered = core.sort_images_newest_first([d2, today, d3, d1], mtimes.get)
    assert ordered == [today, d1, d2, d3]


def test_sort_images_newest_first_puts_undated_last_by_mtime():
    odd_new, odd_old = Path("C:/Pictures/Bing/bing_custom.jpg"), Path("C:/Pictures/Bing/bing_old.jpg")
    mtimes = {odd_new: 50, odd_old: 10, _img(0): 1, _img(1): 2}
    ordered = core.sort_images_newest_first([odd_old, _img(1), odd_new, _img(0)], mtimes.get)
    assert ordered == [_img(0), _img(1), odd_new, odd_old]
    assert core.sort_images_newest_first([]) == []


def test_select_flatpak_asset():
    assets = [
        {"name": "InstallBingWallpaper.exe", "browser_download_url": "u1", "digest": "sha256:" + HEX},
        {"name": "BingWallpaper.flatpak", "browser_download_url": "u2", "digest": "sha256:" + HEX},
        {"name": "BingWallpaper.flatpak.sha256", "browser_download_url": "u3"},
    ]
    assert core.select_flatpak_asset(assets) == {"name": "BingWallpaper.flatpak", "url": "u2", "sha256": HEX}
    assert core.select_installer_asset(assets)["name"] == "InstallBingWallpaper.exe"
    assert core.select_flatpak_asset(assets[:1]) is None
