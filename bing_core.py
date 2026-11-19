# bing_core.py
"""Pure, GUI-free logic shared by the tray app and the installer.

Everything here is deterministic and has no side effects (no network, no
registry, no Tk), so it can be unit-tested with plain pytest. See tests/.
"""
import datetime
import hashlib
import re
from pathlib import Path

BING_API_BASE = "https://www.bing.com/HPImageArchive.aspx"
BING_IMAGE_HOST = "https://www.bing.com"
DEFAULT_MARKET = "en-US"
MARKET_PRESETS = [
    "en-US", "en-GB", "en-CA", "en-AU", "en-IN", "de-DE", "fr-FR", "fr-CA",
    "es-ES", "it-IT", "nl-NL", "pt-BR", "ja-JP", "zh-CN",
]
# Resolutions to try in order; Bing does not publish UHD for every image.
IMAGE_RESOLUTION_SUFFIXES = ["_UHD.jpg", "_1920x1080.jpg"]
# After a failed check, retry after these delays (minutes) instead of
# waiting a full interval. The last value repeats.
RETRY_BACKOFF_MINUTES = [2, 5, 15, 30, 60]
# Do not toast the user until this many checks fail in a row (covers the
# network not being up yet right after login).
QUIET_FAILURE_COUNT = 3
MAX_BACKFILL_DAYS = 7  # Bing's archive API returns at most 8 entries.
MAX_API_COUNT = 8

_MARKET_RE = re.compile(r"[A-Za-z]{2}-[A-Za-z]{2}")
_SHA256_RE = re.compile(r"\b[0-9a-fA-F]{64}\b")
_PAC_PROXY_RE = re.compile(r"PROXY\s+([a-zA-Z0-9.-]+:\d+)")


# --------------------------------------------------------------------------
# Versions
# --------------------------------------------------------------------------
def version_tuple(version, parts=3):
    """'v1.3.8-beta' -> (1, 3, 8). Non-numeric chunks count as 0."""
    clean = str(version or "").strip().lower().lstrip("v")
    out = []
    for chunk in clean.split("."):
        match = re.match(r"\d+", chunk)   # leading digits only: '8-beta2' -> 8
        out.append(int(match.group(0)) if match else 0)
    while len(out) < parts:
        out.append(0)
    return tuple(out[:parts])


def is_newer_version(candidate, current):
    return version_tuple(candidate) > version_tuple(current)


# --------------------------------------------------------------------------
# Settings helpers
# --------------------------------------------------------------------------
def clamp_int(value, default, minimum, maximum):
    """Parse an int from config, falling back to default and clamping."""
    try:
        number = int(str(value).strip())
    except (TypeError, ValueError, AttributeError):
        return default
    return max(minimum, min(number, maximum))


def normalize_market(market):
    market = str(market or "").strip()
    return market if _MARKET_RE.fullmatch(market) else DEFAULT_MARKET


def bing_api_url(market=DEFAULT_MARKET, count=1):
    count = max(1, min(int(count), MAX_API_COUNT))
    return f"{BING_API_BASE}?format=js&idx=0&n={count}&mkt={normalize_market(market)}"


# --------------------------------------------------------------------------
# Bing API payload
# --------------------------------------------------------------------------
def parse_bing_date(date_str):
    """'20261119' -> date(2026, 11, 19), or None."""
    try:
        return datetime.datetime.strptime(str(date_str), "%Y%m%d").date()
    except (TypeError, ValueError):
        return None


def image_filename(date_str):
    return f"bing_{date_str}.jpg"


def parse_image_date(path):
    """Date encoded in a wallpaper file name (bing_YYYYMMDD.jpg), or None."""
    stem = Path(str(path)).stem
    return parse_bing_date(stem.split("_")[-1])


def parse_bing_images(data):
    """Turn the HPImageArchive JSON into a list of dicts, newest first:
    {'date': 'YYYYMMDD', 'urls': [...best first...], 'title': str}.
    Entries missing a date or urlbase are skipped."""
    entries = []
    for img in (data or {}).get("images", []) or []:
        urlbase = img.get("urlbase")
        date = img.get("startdate")
        if not urlbase or not date:
            continue
        entries.append({
            "date": str(date),
            "urls": [BING_IMAGE_HOST + urlbase + suffix for suffix in IMAGE_RESOLUTION_SUFFIXES],
            "title": img.get("title") or img.get("copyright") or "",
        })
    return entries


# --------------------------------------------------------------------------
# Proxy
# --------------------------------------------------------------------------
def parse_proxy_string(proxy_str):
    """'host:8080' -> ('host', '8080'); 'http://host' -> ('host', '80');
    empty -> (None, None)."""
    if not proxy_str:
        return None, None
    clean = str(proxy_str).strip()
    clean = re.sub(r"^[a-zA-Z]+://", "", clean).split("/")[0]
    if not clean:
        return None, None
    host, sep, port = clean.partition(":")
    return host, (port if sep and port else "80")


def build_proxy_dict(host, port):
    host = str(host or "").strip()
    port = str(port or "").strip()
    if host and port:
        proxy = f"http://{host}:{port}"
        return {"http": proxy, "https": proxy}
    return None


def extract_pac_proxy(pac_text):
    """First 'PROXY host:port' directive in a PAC file, or None."""
    match = _PAC_PROXY_RE.search(pac_text or "")
    return match.group(1) if match else None


# --------------------------------------------------------------------------
# Failure handling
# --------------------------------------------------------------------------
def retry_delay_seconds(consecutive_failures, interval_seconds):
    """Backoff for the next retry, never longer than the normal interval."""
    index = max(0, min(consecutive_failures, len(RETRY_BACKOFF_MINUTES)) - 1)
    delay = RETRY_BACKOFF_MINUTES[index] * 60
    return min(delay, interval_seconds) if interval_seconds > 0 else delay


def should_notify_failure(consecutive_failures, message, last_error):
    """Stay quiet for the first failures, toast once at the threshold, then
    only when the error message changes."""
    if consecutive_failures < QUIET_FAILURE_COUNT:
        return False
    if consecutive_failures == QUIET_FAILURE_COUNT:
        return True
    return message != last_error


# --------------------------------------------------------------------------
# Retention / backfill
# --------------------------------------------------------------------------
def select_old_images(paths, keep_days, current_path=None, today=None):
    """Paths whose file-name date is keep_days or more days old, i.e. the
    images within_retention() would reject. keep_days <= 0 keeps everything;
    the current wallpaper is never selected."""
    if keep_days <= 0:
        return []
    today = today or datetime.date.today()
    selected = []
    for path in paths:
        img_date = parse_image_date(path)
        if img_date is None or (today - img_date).days < keep_days:
            continue
        if current_path is not None and Path(path) == Path(current_path):
            continue
        selected.append(path)
    return selected


def within_retention(date_str, keep_days, today=None):
    """True when an image dated date_str would survive retention."""
    if keep_days <= 0:
        return True
    img_date = parse_bing_date(date_str)
    if img_date is None:
        return False
    today = today or datetime.date.today()
    return (today - img_date).days < keep_days


# --------------------------------------------------------------------------
# Updater
# --------------------------------------------------------------------------
def parse_asset_digest(digest):
    """GitHub asset 'digest' field ('sha256:<hex>') -> lowercase hex or None."""
    if not digest:
        return None
    algo, sep, value = str(digest).strip().partition(":")
    if sep and algo.lower() == "sha256" and _SHA256_RE.fullmatch(value):
        return value.lower()
    return None


def find_sha256_in_text(text, asset_name=None):
    """Find a SHA-256 published in release notes. Prefers a hash on the same
    line as asset_name; otherwise the first 64-hex token."""
    text = text or ""
    if asset_name:
        for line in text.splitlines():
            if asset_name.lower() in line.lower():
                match = _SHA256_RE.search(line)
                if match:
                    return match.group(0).lower()
    match = _SHA256_RE.search(text)
    return match.group(0).lower() if match else None


def select_installer_asset(assets, release_body=""):
    """Pick the installer .exe from a release's assets. Returns
    {'name', 'url', 'sha256'} (sha256 may be None) or None."""
    for asset in assets or []:
        name = str(asset.get("name", ""))
        lowered = name.lower()
        if lowered.endswith(".exe") and ("installbingwallpaper" in lowered or "bingwallpaper" in lowered):
            sha256 = parse_asset_digest(asset.get("digest")) or find_sha256_in_text(release_body, name)
            return {
                "name": name,
                "url": asset.get("browser_download_url"),
                "sha256": sha256,
            }
    return None


def sha256_of_file(path, chunk_size=1024 * 1024):
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()
