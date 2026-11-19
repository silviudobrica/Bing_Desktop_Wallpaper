# bing_service.py
"""Platform-neutral application logic shared by the Windows tray app and the
Linux (KDE/Flatpak) app.

WallpaperService owns the configuration, the download/backfill/retention
pipeline, the check loop with backoff, and the status fields shown in the
UI. Everything OS-specific goes through a Platform object (wallpaper, lock
screen, proxies, notifications, autostart). UI front ends subclass the
service and override the on_* hooks; those hooks may be called from worker
threads, so a UI must marshal to its own thread.
"""
import datetime
import json
import tempfile
import logging
import os
import threading
import time
import webbrowser
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from PIL import Image

import bing_core as core
from bing_core import DEFAULT_MARKET, MAX_BACKFILL_DAYS

RELEASES_API = "https://api.github.com/repos/SilviuDobrica/Bing_Desktop_Wallpaper/releases/latest"
DEFAULT_INTERVAL_MINUTES = 720
UPDATE_CHECK_SECONDS = 86400

INTERVAL_PRESETS = {
    "15 Minutes": 15,
    "30 Minutes": 30,
    "1 Hour": 60,
    "4 Hours": 240,
    "6 Hours": 360,
    "12 Hours": 720,
    "24 Hours": 1440,
    "Disabled": 0,
}


def log_msg(msg, level="info"):
    print(msg)
    if level == "error":
        logging.error(msg)
    else:
        logging.info(msg)


class Platform:
    """What the service needs from the host OS. Methods raise or return
    (ok, detail) tuples; they never touch the UI. The service binds itself
    with bind() so a platform can reuse its HTTP session (PAC downloads)."""

    name = "generic"
    # Windows persists the PAC-detected proxy into config.json on first run.
    # Desktops whose proxy can change at any time (Linux) must not pin it.
    autosave_detected_proxy = False
    supports_lock_screen = False

    def __init__(self):
        self.service = None

    def bind(self, service):
        self.service = service

    def set_wallpaper(self, image_path):
        raise NotImplementedError

    def set_lock_screen(self, image_path):
        """Return (ok, detail). detail names the method used, or the error."""
        return False, "lock screen is not supported on this platform"

    def lock_screen_policy(self):
        """Description of a policy that pins the lock screen, or None."""
        return None

    def system_proxies(self, url=None):
        """Current system proxy settings as a requests proxies dict ({} when
        direct). Called before every request, never cached by the service."""
        return {}

    def detect_proxy_values(self):
        """(host, port) the Settings 'Auto-Detect Proxy' button fills in."""
        return "", ""

    def notify(self, title, message):
        pass

    def is_startup_enabled(self):
        return False

    def set_startup_enabled(self, enabled):
        raise NotImplementedError

    def open_path(self, path):
        webbrowser.open(Path(path).as_uri())

    def open_url(self, url):
        webbrowser.open(url)


class WallpaperService:
    def __init__(self, platform, config_file, image_dir, app_name, version):
        self.platform = platform
        self.config_file = Path(config_file)
        self.image_dir = Path(image_dir)
        self.app_name = app_name
        self.version = version

        self.last_check = 0
        self._next_check_at = 0
        self.consecutive_failures = 0
        self.current_image_path = None
        self.running = True
        self.session = self._create_retry_session()
        self.last_success = 0
        self.last_error = "None"
        self.last_error_time = 0
        self.last_status_message = "Waiting for first check"
        self.last_update_check = 0
        self.latest_release_version = None
        self.latest_release_url = None
        # Serializes wallpaper checks so startup, Check Now, F5 and the interval
        # loop never download the same image concurrently.
        self._check_lock = threading.Lock()

        self.image_dir.mkdir(parents=True, exist_ok=True)
        self.cleanup_partial_downloads()

        self.config = self.load_config()
        interval_minutes = self.config.get("check_interval_minutes", DEFAULT_INTERVAL_MINUTES)
        self.check_interval = interval_minutes * 60 if interval_minutes > 0 else 0

        platform.bind(self)
        log_msg(f"Initializing {app_name} v{version} ({platform.name})")

        if platform.autosave_detected_proxy and not self.config.get("proxy_url"):
            self.detect_and_save_proxy()

    # ------------------------------------------------------------ UI hooks
    def on_state_changed(self):
        """Status fields changed (any thread)."""

    def on_wallpaper_applied(self, image_path):
        """The desktop wallpaper was just set to image_path (any thread)."""

    def on_images_changed(self):
        """Images were added or removed on disk (any thread)."""

    def notify_user(self, title, message):
        try:
            self.platform.notify(title, message)
        except Exception:
            pass
        log_msg(f"{title}: {message}")

    def check_for_updates(self, manual=False):
        """Hook for the front end's update strategy; the base does nothing."""
        self.last_update_check = time.time()

    # ------------------------------------------------------------ network
    def _create_retry_session(self):
        session = requests.Session()
        retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
        session.mount('http://', HTTPAdapter(max_retries=retries))
        session.mount('https://', HTTPAdapter(max_retries=retries))
        return session

    def get_verify_option(self):
        ca_bundle_path = str(self.config.get("ca_bundle_path", "")).strip()
        if not ca_bundle_path:
            return True
        if Path(ca_bundle_path).exists():
            return ca_bundle_path
        log_msg(f"Configured CA bundle path not found: {ca_bundle_path}", "error")
        return True

    def session_get(self, url, **kwargs):
        kwargs.setdefault("verify", self.get_verify_option())
        return self.session.get(url, **kwargs)

    def get_proxy_dict(self, url=None):
        """Explicit proxy from Settings wins; otherwise ask the platform for
        the system proxy right now, so a changed proxy is picked up by the
        very next request."""
        explicit = core.build_proxy_dict(self.config.get("proxy_url", ""), self.config.get("proxy_port", ""))
        if explicit:
            return explicit
        try:
            return self.platform.system_proxies(url) or {}
        except Exception as e:
            log_msg(f"System proxy lookup failed: {e}", "error")
            return {}

    def detect_proxy_values(self):
        try:
            host, port = self.platform.detect_proxy_values()
            return host or "", port or ""
        except Exception as e:
            log_msg(f"Proxy detection failed: {e}", "error")
            return "", ""

    def detect_and_save_proxy(self):
        try:
            host, port = self.detect_proxy_values()
            if host:
                self.config['proxy_url'] = host
                self.config['proxy_port'] = port
                self.save_config()
        except Exception:
            pass

    def fetch_text(self, url, timeout=5):
        """Small helper for platforms (PAC scripts)."""
        resp = self.session_get(url, timeout=timeout)
        return resp.text if resp.status_code == 200 else None

    def fetch_latest_release(self):
        """Latest GitHub release metadata, or None."""
        headers = {"Accept": "application/vnd.github+json", "User-Agent": f"{self.app_name}/{self.version}"}
        resp = self.session_get(RELEASES_API, timeout=10, proxies=self.get_proxy_dict(RELEASES_API), headers=headers)
        resp.raise_for_status()
        return resp.json()

    def download_release_asset(self, url, name, sha256=None, target_dir=None, timeout=180):
        """Download a release asset to target_dir/name through the configured
        proxy and CA bundle, verify it against sha256 when one is known, and
        return the final path. A mismatch discards the file and raises."""
        target_dir = Path(target_dir) if target_dir else Path(tempfile.gettempdir()) / self.app_name / "updates"
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / name
        temp_target = target.with_suffix(target.suffix + ".download")
        log_msg(f"Downloading {name} from {url} to {target}")
        headers = {"User-Agent": f"{self.app_name}/{self.version}"}
        resp = self.session_get(url, timeout=timeout, proxies=self.get_proxy_dict(url), headers=headers, stream=True)
        resp.raise_for_status()
        try:
            with open(temp_target, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
            if sha256:
                actual = core.sha256_of_file(temp_target)
                if actual != sha256:
                    raise RuntimeError(
                        f"{name} checksum mismatch: the download does not match the hash published "
                        f"on GitHub (expected {sha256[:12]}..., got {actual[:12]}...). The file was discarded."
                    )
                log_msg(f"{name} SHA-256 verified against the release digest.")
            else:
                log_msg(f"No SHA-256 published for {name}; skipping checksum verification.")
            os.replace(temp_target, target)
        finally:
            try:
                if temp_target.exists():
                    temp_target.unlink()
            except Exception:
                pass
        log_msg(f"Downloaded successfully: {target}")
        return target

    # ------------------------------------------------------------ config
    def load_config(self):
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            log_msg(f"Error loading config: {e}", "error")
        return {}

    def save_config(self):
        # Write to a sibling temp file and atomically replace, so a crash or
        # power loss mid-write never leaves a truncated config.json behind.
        temp_path = self.config_file.with_name(self.config_file.name + ".tmp")
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
                f.flush()
                os.fsync(f.fileno())
            os.replace(temp_path, self.config_file)
        except Exception as e:
            log_msg(f"Error saving config: {e}", "error")
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except Exception:
                pass

    def get_int_setting(self, key, default, minimum, maximum):
        return core.clamp_int(self.config.get(key, default), default, minimum, maximum)

    def get_market(self):
        return core.normalize_market(self.config.get("market", DEFAULT_MARKET))

    def bing_api_url(self, count=1):
        return core.bing_api_url(self.get_market(), count)

    def interval_minutes(self):
        return self.config.get("check_interval_minutes", DEFAULT_INTERVAL_MINUTES)

    def set_interval(self, minutes):
        self.config["check_interval_minutes"] = minutes
        self.save_config()
        self.check_interval = minutes * 60 if minutes > 0 else 0
        self._next_check_at = 0  # run a check right away
        self.on_state_changed()

    # ------------------------------------------------------------ status
    def format_ts(self, timestamp):
        if not timestamp:
            return "Never"
        return datetime.datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")

    def is_newer_version(self, candidate):
        return core.is_newer_version(candidate, self.version)

    def record_error(self, message, notify=None):
        """Record a failure. notify=None toasts only when the message changed;
        True/False force or suppress the toast."""
        previous = self.last_error
        self.last_error = message
        self.last_error_time = time.time()
        self.last_status_message = "Last operation failed"
        self.on_state_changed()
        if notify is None:
            notify = message != previous
        if notify:
            self.notify_user(self.app_name, f"Operation failed: {message}")
        else:
            log_msg(f"Operation failed (notification suppressed): {message}")

    def set_status(self, message):
        self.last_status_message = message
        self.on_state_changed()

    # ------------------------------------------------------------ platform
    def is_startup_enabled(self):
        try:
            return bool(self.platform.is_startup_enabled())
        except Exception:
            return False

    def set_startup_enabled(self, enabled):
        try:
            self.platform.set_startup_enabled(bool(enabled))
        except Exception as e:
            self.record_error(f"Startup toggle failed: {e}")

    def open_path(self, path):
        try:
            self.platform.open_path(path)
        except Exception as e:
            self.record_error(f"Unable to open path: {e}")

    def open_url(self, url):
        try:
            self.platform.open_url(url)
        except Exception as e:
            self.record_error(f"Unable to open link: {e}")

    def lock_screen_policy_source(self):
        try:
            return self.platform.lock_screen_policy()
        except Exception:
            return None

    # ------------------------------------------------------------ images
    def get_bing_images(self, count=1):
        """Return the latest `count` Bing images, newest first, as dicts with
        'date' (YYYYMMDD), 'urls' (candidate download URLs, best first) and
        'title'. Returns None on failure."""
        try:
            url = self.bing_api_url(count)
            resp = self.session_get(url, timeout=10, proxies=self.get_proxy_dict(url))
            resp.raise_for_status()
            return core.parse_bing_images(resp.json()) or None
        except Exception as e:
            log_msg(f"API Fetch Error: {e}", "error")
            return None

    def cleanup_partial_downloads(self):
        """Remove leftover temp files from interrupted or failed downloads."""
        for pattern in ("bing_*.tmp", "bing_*.part"):
            for leftover in self.image_dir.glob(pattern):
                try:
                    leftover.unlink()
                    log_msg(f"Removed stale partial download: {leftover.name}")
                except Exception as e:
                    log_msg(f"Could not remove stale partial download {leftover.name}: {e}", "error")

    def download_image(self, urls, date_str):
        """Download the image for date_str, trying each URL in `urls` in order
        (a str is accepted too). A 404, wrong content type or corrupt file moves
        on to the next candidate; network errors abort."""
        if isinstance(urls, str):
            urls = [urls]
        filename = core.image_filename(date_str)
        file_path = self.image_dir / filename

        if file_path.exists() and file_path.stat().st_size > 0:
            return file_path

        # Unique per process/thread so two downloads can never collide on the
        # same temp file (this was the source of WinError 32 sharing violations).
        temp_path = self.image_dir / f"{filename}.{os.getpid()}.{threading.get_ident()}.part"
        try:
            for index, url in enumerate(urls):
                resp = self.session_get(url, timeout=30, proxies=self.get_proxy_dict(url), stream=True)
                if resp.status_code == 404 and index < len(urls) - 1:
                    log_msg(f"Not available at this resolution, trying next: {url}")
                    continue
                resp.raise_for_status()

                content_type = resp.headers.get('Content-Type', '')
                if 'image' not in content_type:
                    log_msg(f"Download Error: unexpected Content-Type '{content_type}' for {url}", "error")
                    continue

                with open(temp_path, 'wb') as f:
                    for chunk in resp.iter_content(chunk_size=8192):
                        f.write(chunk)

                try:
                    with Image.open(temp_path) as img:
                        img.verify()
                except Exception as e:
                    log_msg(f"Downloaded image failed verification ({url}): {e}", "error")
                    continue

                os.replace(temp_path, file_path)
                return file_path

            log_msg(f"Download Error: no usable image found for {date_str}", "error")
            return None

        except Exception as e:
            # If another check finished the same image meanwhile, use it.
            if file_path.exists() and file_path.stat().st_size > 0:
                return file_path
            log_msg(f"Download Error: {e}", "error")
            return None
        finally:
            try:
                if temp_path.exists():
                    temp_path.unlink()
            except Exception:
                pass

    def list_images_newest_first(self):
        return core.sort_images_newest_first(self.image_dir.glob("bing_*.jpg"), os.path.getmtime)

    def latest_local_image(self):
        """Most recent downloaded wallpaper on disk, or None."""
        images = self.list_images_newest_first()
        return images[0] if images else None

    def set_wallpaper(self, image_path):
        image_path = Path(image_path) if image_path else None
        if not image_path or not image_path.exists():
            return False
        try:
            log_msg(f"Setting wallpaper: {image_path.name}")
            self.platform.set_wallpaper(image_path)
        except Exception as e:
            log_msg(f"Wallpaper Set Error: {e}", "error")
            return False
        self.current_image_path = image_path
        self.on_wallpaper_applied(image_path)
        if self.config.get("set_lock_screen", False):
            self.apply_lock_screen(image_path)
        return True

    def apply_lock_screen(self, image_path):
        """Set the lock screen image. Never blocks the caller's (UI) thread."""
        if threading.current_thread() is threading.main_thread():
            threading.Thread(target=self._set_lock_screen_blocking, args=(image_path,), daemon=True).start()
        else:
            self._set_lock_screen_blocking(image_path)

    def _set_lock_screen_blocking(self, image_path):
        image_path = Path(image_path)
        policy = self.lock_screen_policy_source()
        if policy:
            self.record_error(f"Lock screen is managed by your organization ({policy}); the app cannot change it")
            return False
        try:
            ok, detail = self.platform.set_lock_screen(image_path)
        except Exception as e:
            ok, detail = False, str(e)
        if ok:
            log_msg(f"Lock screen updated: {image_path.name} via {detail}")
            return True
        self.record_error(f"Lock screen update failed: {detail}")
        return False

    def cleanup_old_images(self):
        """Delete wallpapers older than keep_images_days (0 = keep forever).
        The image currently in use is never deleted."""
        keep_days = self.get_int_setting("keep_images_days", 0, 0, 3650)
        if keep_days <= 0:
            return 0
        removed = 0
        for img in core.select_old_images(self.image_dir.glob("bing_*.jpg"), keep_days, self.current_image_path):
            try:
                img.unlink()
                removed += 1
            except Exception as e:
                log_msg(f"Could not delete old wallpaper {img.name}: {e}", "error")
        if removed:
            log_msg(f"Retention: removed {removed} wallpaper(s) older than {keep_days} days")
        return removed

    def backfill_images(self, entries):
        """Download any of the given (older) Bing entries that are missing
        locally. Returns the number of images downloaded."""
        keep_days = self.get_int_setting("keep_images_days", 0, 0, 3650)
        downloaded = 0
        for entry in entries:
            if not self.running:
                break
            if not core.within_retention(entry["date"], keep_days):
                continue
            file_path = self.image_dir / core.image_filename(entry["date"])
            if file_path.exists() and file_path.stat().st_size > 0:
                continue
            if self.download_image(entry["urls"], entry["date"]):
                downloaded += 1
                log_msg(f"Backfilled missed wallpaper: {file_path.name}")
        return downloaded

    # ------------------------------------------------------------ check loop
    def check_and_update(self, force=False):
        if not self._check_lock.acquire(blocking=False):
            log_msg("Wallpaper check already in progress; skipping duplicate request.")
            return
        try:
            self._check_and_update_locked(force)
        finally:
            self._check_lock.release()

    def _schedule_next_check(self, seconds):
        self._next_check_at = time.time() + max(0, seconds)

    def _on_check_failed(self, message):
        self.consecutive_failures += 1
        if self.check_interval > 0:
            delay = core.retry_delay_seconds(self.consecutive_failures, self.check_interval)
            self._schedule_next_check(delay)
            log_msg(f"Wallpaper check failed ({self.consecutive_failures} in a row); retrying in {delay // 60} min")
        # Stay quiet for the first few failures (network often is not up yet
        # right after login); toast once when the threshold is reached, and
        # again only if the error changes.
        notify = core.should_notify_failure(self.consecutive_failures, message, self.last_error)
        self.record_error(message, notify=notify)

    def _check_and_update_locked(self, force=False):
        self.last_check = time.time()
        try:
            backfill_days = self.get_int_setting("backfill_days", MAX_BACKFILL_DAYS, 0, MAX_BACKFILL_DAYS)
            images = self.get_bing_images(1 + backfill_days)
            if not images:
                self._on_check_failed("No image info returned from Bing API")
                return

            today = images[0]
            path = self.download_image(today["urls"], today["date"])
            if not path:
                self._on_check_failed("Image download failed")
                return

            if force or self.current_image_path != path:
                if self.set_wallpaper(path):
                    self.last_status_message = f"Wallpaper updated: {path.name}"
                    self.notify_user(self.app_name, f"Wallpaper updated: {path.name}")
                else:
                    self.last_status_message = f"Wallpaper available: {path.name}"
            else:
                self.last_status_message = f"Wallpaper available: {path.name}"

            self.last_success = time.time()
            self.last_error = "None"
            self.last_error_time = 0
            self.consecutive_failures = 0
            self._schedule_next_check(self.check_interval)

            self.on_images_changed()
            self.on_state_changed()

            # Housekeeping after today's wallpaper is in place.
            changed = self.backfill_images(images[1:])
            changed += self.cleanup_old_images()
            if changed:
                self.on_images_changed()

        except Exception as e:
            log_msg(f"Update Loop Error: {e}", "error")
            self._on_check_failed(str(e))

    def background_loop(self):
        # Initial forced check at startup. Running it here instead of in a
        # separate thread avoids racing the interval check below.
        self.check_and_update(force=True)
        while self.running:
            try:
                if self.check_interval > 0 and time.time() >= self._next_check_at:
                    self.check_and_update(force=False)

                # Auto-check updates once every 24 hours.
                if (time.time() - self.last_update_check) > UPDATE_CHECK_SECONDS:
                    self.check_for_updates(manual=False)
                time.sleep(5)
            except Exception:
                time.sleep(60)

    def stop(self):
        self.running = False
