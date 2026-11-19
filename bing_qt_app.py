# bing_qt_app.py
"""Linux (KDE-first) front end for Bing Daily Wallpaper, built with PySide6.

The look and behaviour mirror the Windows Tk app (bing_daily_wallpaper.py):
a tray icon with the same menu, a Preview / Gallery window with the status
panel, a scaled preview and the "Latest 5" strip, and the same Settings
dialog. All download, backfill, retention and config logic lives in
bing_service.WallpaperService; the OS integration (wallpaper, lock screen,
proxies, autostart, notifications) lives in platform_linux.LinuxPlatform.

Threading: the service runs its check loop on a daemon thread and calls the
on_* hooks from there. QtWallpaperService turns those hooks into Qt signals;
the receivers live on the GUI thread, so Qt delivers them as queued calls.
"""
import argparse
import datetime
import logging
import os
import re
import sys
import threading
import time
from logging.handlers import RotatingFileHandler
from pathlib import Path

from PySide6.QtCore import QObject, QSize, Qt, QTimer, Signal
from PySide6.QtGui import (
    QAction, QActionGroup, QColor, QFont, QIcon, QImageReader, QKeySequence, QPixmap, QShortcut,
)
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QFileDialog, QFormLayout, QFrame, QGridLayout,
    QGroupBox, QHBoxLayout, QInputDialog, QLabel, QLineEdit, QMainWindow, QMenu, QMessageBox,
    QPushButton, QScrollBar, QSizePolicy, QSystemTrayIcon, QVBoxLayout, QWidget,
)

import bing_core as core
from bing_core import DEFAULT_MARKET, MARKET_PRESETS, MAX_BACKFILL_DAYS
from bing_service import INTERVAL_PRESETS, WallpaperService, log_msg
from platform_linux import APP_DISPLAY_NAME, APP_ID, LinuxPlatform, linux_paths

try:
    from _version import __version__ as VERSION
except ImportError:
    VERSION = "1.4.0"

# Name of the launcher script the Flatpak manifest installs into /app/bin.
LAUNCHER_NAME = "bing-wallpaper"
GALLERY_PAGE = 5
PREVIEW_MAX = QSize(1920, 1080)
THUMB_MAX = QSize(480, 270)
_MARKET_RE = re.compile(r"[A-Za-z]{2}-[A-Za-z]{2}")


def configure_logging(log_dir):
    """Rotating log file like the Windows app (5 MB, two backups)."""
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(log_dir / "bing_wallpaper.log", maxBytes=5 * 1024 * 1024,
                                  backupCount=2, encoding="utf-8")
    logging.basicConfig(
        handlers=[handler],
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - [%(threadName)s] - %(message)s",
    )


def launch_command(in_flatpak):
    """Argument vector the autostart entry runs to start us hidden."""
    if in_flatpak:
        return [LAUNCHER_NAME, "--hidden"]
    return [sys.executable, str(Path(__file__).resolve()), "--hidden"]


def display_date(path):
    """'bing_20261112.jpg' -> 'Nov 12, 2026' (the stem when it has no date)."""
    date_str = path.stem.split("_")[-1]
    try:
        return datetime.datetime.strptime(date_str, "%Y%m%d").strftime("%b %d, %Y")
    except ValueError:
        return path.stem


def load_scaled_pixmap(path, max_size):
    """Decode an image already scaled down (UHD wallpapers are 30 MB as raw
    pixmaps; the reader scales while decoding instead)."""
    reader = QImageReader(str(path))
    reader.setAutoTransform(True)
    size = reader.size()
    if size.isValid() and (size.width() > max_size.width() or size.height() > max_size.height()):
        reader.setScaledSize(size.scaled(max_size, Qt.KeepAspectRatio))
    image = reader.read()
    if image.isNull():
        return None
    return QPixmap.fromImage(image)


def fallback_icon():
    pixmap = QPixmap(64, 64)
    pixmap.fill(QColor(0, 120, 215))
    return QIcon(pixmap)


# ---------------------------------------------------------------- service bridge
class ServiceSignals(QObject):
    state_changed = Signal()
    wallpaper_applied = Signal(str)
    images_changed = Signal()
    notification = Signal(str, str)
    update_result = Signal(dict, bool)
    update_install_result = Signal(bool, str)


class QtWallpaperService(WallpaperService):
    """WallpaperService whose hooks are re-emitted as Qt signals."""

    def __init__(self, platform, paths):
        self.signals = ServiceSignals()
        super().__init__(platform, paths.config_file, paths.image_dir, APP_DISPLAY_NAME, VERSION)

    def on_state_changed(self):
        self.signals.state_changed.emit()

    def on_wallpaper_applied(self, image_path):
        self.signals.wallpaper_applied.emit(str(image_path))

    def on_images_changed(self):
        self.signals.images_changed.emit()

    def notify_user(self, title, message):
        log_msg(f"{title}: {message}")
        self.signals.notification.emit(title, message)

    def check_for_updates(self, manual=False):
        """Compare the latest GitHub release with VERSION. In Flatpak the
        bundle asset is offered for download and installation through the
        host; otherwise the UI offers the release page."""
        self.last_update_check = time.time()
        try:
            data = self.fetch_latest_release()
            version = str(data.get("tag_name", "")).lstrip("v")
            html_url = data.get("html_url", "")
            if not version:
                raise ValueError("release has no tag")
            asset = core.select_flatpak_asset(data.get("assets", []), data.get("body", ""))
        except Exception as e:
            log_msg(f"Update check failed: {e}", "error")
            if manual:
                self.signals.update_result.emit({}, True)
            return
        self.latest_release_version = version
        self.latest_release_url = html_url
        newer = self.is_newer_version(version)
        if newer:
            self.last_status_message = f"Update available: v{version}"
            self.on_state_changed()
            self.notify_user(self.app_name, f"Update available: v{version}")
        else:
            self.last_status_message = f"Up to date: v{self.version}"
            self.on_state_changed()
            log_msg(f"No newer version found. Current={self.version}, Latest={version}")
        self.signals.update_result.emit(
            {"version": version, "html_url": html_url, "newer": newer, "asset": asset}, manual)

    def perform_update_install(self, info, updates_dir):
        """Worker: download the bundle, verify it, install it via the host."""
        asset = info.get("asset") or {}
        try:
            if not asset.get("url"):
                raise RuntimeError("the release has no Flatpak bundle attached")
            self.set_status(f"Downloading update v{info['version']}...")
            bundle = self.download_release_asset(asset["url"], asset["name"], asset.get("sha256"), updates_dir, timeout=600)
            self.set_status(f"Installing update v{info['version']}...")
            ok, detail = self.platform.install_update_bundle(bundle)
            if not ok:
                raise RuntimeError(detail)
            try:
                bundle.unlink()
            except OSError:
                pass
            self.set_status(f"Update v{info['version']} installed; restarting")
            log_msg(f"Update installed: {detail}")
            self.signals.update_install_result.emit(True, detail)
        except Exception as e:
            self.record_error(f"Update install failed: {e}")
            self.signals.update_install_result.emit(False, str(e))


# ---------------------------------------------------------------- widgets
class ClickableLabel(QLabel):
    clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setCursor(Qt.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMinimumSize(80, 60)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.rect().contains(event.position().toPoint()):
            self.clicked.emit()
        super().mouseReleaseEvent(event)


class SettingsDialog(QDialog):
    """Same fields, defaults and validation as the Tk settings window."""

    def __init__(self, service, parent=None):
        super().__init__(parent)
        self.service = service
        self.setWindowTitle("Settings")
        self.setMinimumSize(540, 600)
        config = service.config

        layout = QVBoxLayout(self)
        bold = QFont()
        bold.setBold(True)

        # Interval
        title = QLabel("Check Interval")
        title.setFont(bold)
        layout.addWidget(title)
        row = QHBoxLayout()
        self.interval_combo = QComboBox()
        self.interval_combo.addItems(list(INTERVAL_PRESETS.keys()))
        current = config.get("check_interval_minutes", 720)
        self.interval_combo.setCurrentText("12 Hours")
        for label, minutes in INTERVAL_PRESETS.items():
            if minutes == current:
                self.interval_combo.setCurrentText(label)
                break
        row.addWidget(self.interval_combo, 1)
        row.addWidget(QLabel("Custom minutes (optional)"))
        self.custom_interval = QLineEdit("" if current in INTERVAL_PRESETS.values() else str(current))
        self.custom_interval.setMaximumWidth(90)
        row.addWidget(self.custom_interval)
        layout.addLayout(row)
        layout.addWidget(self._separator())

        # Proxy
        title = QLabel("Proxy")
        title.setFont(bold)
        layout.addWidget(title)
        form = QFormLayout()
        self.proxy_host = QLineEdit(str(config.get("proxy_url", "")))
        form.addRow("Host", self.proxy_host)
        port_row = QHBoxLayout()
        self.proxy_port = QLineEdit(str(config.get("proxy_port", "")))
        self.proxy_port.setMaximumWidth(90)
        port_row.addWidget(self.proxy_port)
        port_row.addStretch(1)
        detect_btn = QPushButton("Auto-Detect Proxy")
        detect_btn.clicked.connect(self.auto_detect_proxy)
        port_row.addWidget(detect_btn)
        form.addRow("Port", port_row)
        ca_row = QHBoxLayout()
        self.ca_bundle = QLineEdit(str(config.get("ca_bundle_path", "")))
        ca_row.addWidget(self.ca_bundle, 1)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_ca_bundle)
        ca_row.addWidget(browse_btn)
        form.addRow("CA Bundle (PEM, optional)", ca_row)
        layout.addLayout(form)
        layout.addWidget(self._separator())

        # Wallpaper
        title = QLabel("Wallpaper")
        title.setFont(bold)
        layout.addWidget(title)
        form = QFormLayout()
        market_row = QHBoxLayout()
        self.market = QComboBox()
        self.market.setEditable(True)
        self.market.addItems(MARKET_PRESETS)
        self.market.setCurrentText(str(config.get("market", DEFAULT_MARKET)))
        market_row.addWidget(self.market)
        market_row.addWidget(QLabel("(region of the daily image, e.g. en-US)"))
        market_row.addStretch(1)
        form.addRow("Bing market", market_row)
        keep_row = QHBoxLayout()
        self.keep_days = QLineEdit(str(service.get_int_setting("keep_images_days", 0, 0, 3650)))
        self.keep_days.setMaximumWidth(70)
        keep_row.addWidget(self.keep_days)
        keep_row.addWidget(QLabel("(0 = keep forever)"))
        keep_row.addStretch(1)
        form.addRow("Keep images for (days)", keep_row)
        backfill_row = QHBoxLayout()
        self.backfill_days = QLineEdit(str(service.get_int_setting("backfill_days", MAX_BACKFILL_DAYS, 0, MAX_BACKFILL_DAYS)))
        self.backfill_days.setMaximumWidth(70)
        backfill_row.addWidget(self.backfill_days)
        backfill_row.addWidget(QLabel(f"(0-{MAX_BACKFILL_DAYS}, downloads images you missed)"))
        backfill_row.addStretch(1)
        form.addRow("Backfill missed days", backfill_row)
        layout.addLayout(form)
        self.lock_screen = QCheckBox("Also set the lock screen image")
        self.lock_screen.setChecked(bool(config.get("set_lock_screen", False)))
        self.lock_screen.setEnabled(bool(getattr(service.platform, "supports_lock_screen", False)))
        if not self.lock_screen.isEnabled():
            self.lock_screen.setToolTip("Not supported on this desktop")
        layout.addWidget(self.lock_screen)
        layout.addWidget(self._separator())

        self.startup = QCheckBox("Run at Startup")
        self.startup.setChecked(service.is_startup_enabled())
        layout.addWidget(self.startup)
        layout.addStretch(1)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        save_btn = QPushButton("Save")
        save_btn.setDefault(True)  # Enter saves
        save_btn.clicked.connect(self.save)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        buttons.addWidget(save_btn)
        buttons.addWidget(cancel_btn)
        layout.addLayout(buttons)

    @staticmethod
    def _separator():
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        return line

    def auto_detect_proxy(self):
        host, port = self.service.detect_proxy_values()
        if host:
            self.proxy_host.setText(host)
            self.proxy_port.setText(port)
        else:
            QMessageBox.information(self, "Proxy Detection", "No proxy detected from system settings.")

    def browse_ca_bundle(self):
        selected, _ = QFileDialog.getOpenFileName(self, "Select CA Bundle (PEM)", "",
                                                  "PEM files (*.pem);;All files (*)")
        if selected:
            self.ca_bundle.setText(selected)

    def save(self):
        service = self.service
        try:
            custom_raw = self.custom_interval.text().strip()
            if custom_raw:
                custom_minutes = int(custom_raw)
                if custom_minutes <= 0:
                    raise ValueError("Custom interval must be positive")
                minutes = custom_minutes
            else:
                minutes = INTERVAL_PRESETS[self.interval_combo.currentText()]
            log_msg(f"Setting interval to: {minutes} minutes")
            service.set_interval(minutes)

            service.config["proxy_url"] = self.proxy_host.text().strip()
            service.config["proxy_port"] = self.proxy_port.text().strip()
            service.config["ca_bundle_path"] = self.ca_bundle.text().strip()

            market = self.market.currentText().strip()
            if market and not _MARKET_RE.fullmatch(market):
                raise ValueError("Market must look like en-US")
            service.config["market"] = market or DEFAULT_MARKET

            keep_days = int(self.keep_days.text().strip() or "0")
            if keep_days < 0:
                raise ValueError("Keep images days cannot be negative")
            service.config["keep_images_days"] = keep_days

            backfill_days = int(self.backfill_days.text().strip() or "0")
            if not 0 <= backfill_days <= MAX_BACKFILL_DAYS:
                raise ValueError(f"Backfill days must be between 0 and {MAX_BACKFILL_DAYS}")
            service.config["backfill_days"] = backfill_days

            lock_screen_enabled = self.lock_screen.isChecked()
            lock_screen_was_enabled = bool(service.config.get("set_lock_screen", False))
            service.config["set_lock_screen"] = lock_screen_enabled
            service.save_config()

            if lock_screen_enabled and not lock_screen_was_enabled:
                policy = service.lock_screen_policy_source()
                if policy:
                    QMessageBox.warning(
                        self, "Lock Screen",
                        "Your organization enforces the lock screen image on this PC\n"
                        f"({policy}).\n\nThe option is saved, but the enforced image will stay.",
                    )
                else:
                    image = service.current_image_path or service.latest_local_image()
                    if image:
                        service.apply_lock_screen(image)
            service.set_startup_enabled(self.startup.isChecked())

            service.set_status("Settings saved")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Settings", f"Unable to save settings: {e}")


class MainWindow(QMainWindow):
    """Preview / Gallery window: actions row, status panel, preview, Latest 5."""

    STATUS_ROWS = [
        ("State", "status"),
        ("Last Check", "last_check"),
        ("Last Success", "last_success"),
        ("Last Error", "last_error"),
        ("Current Image", "current_image"),
        ("Interval", "interval"),
        ("Startup", "startup"),
        ("Desktop", "desktop"),
    ]

    def __init__(self, service, platform, paths, tray_available=True):
        super().__init__()
        self.service = service
        self.platform = platform
        self.paths = paths
        self.tray_available = tray_available
        self.images = []
        self.gallery_offset = 0
        self.preview_source = None
        self.preview_path = None
        self.thumb_cache = {}
        self.thumb_slots = []
        self.status_labels = {}

        self.setWindowTitle(f"{APP_DISPLAY_NAME} v{VERSION}")
        self.resize(840, 760)
        self.setMinimumSize(620, 560)
        self._build_ui()
        self._bind_shortcuts()

        self.preview_timer = QTimer(self)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(45)
        self.preview_timer.timeout.connect(self.redraw_preview)
        self.thumb_timer = QTimer(self)
        self.thumb_timer.setSingleShot(True)
        self.thumb_timer.setInterval(60)
        self.thumb_timer.timeout.connect(self.redraw_thumbnails)

        signals = service.signals
        signals.state_changed.connect(self.refresh_status)
        signals.wallpaper_applied.connect(self.on_wallpaper_applied)
        signals.images_changed.connect(self.reload_gallery)

        self.refresh_status()
        self.reload_gallery()
        self.refresh_preview()

    # ---- construction
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)

        actions = QHBoxLayout()
        for text, slot in (("Check Now", self.run_check), ("Settings", self.show_settings),
                           ("Open Logs", self.open_logs)):
            button = QPushButton(text)
            button.clicked.connect(slot)
            actions.addWidget(button, 1)
        root.addLayout(actions)

        status_box = QGroupBox("Status")
        grid = QGridLayout(status_box)
        grid.setColumnStretch(1, 1)
        for row, (label, key) in enumerate(self.STATUS_ROWS):
            grid.addWidget(QLabel(f"{label}:"), row, 0)
            value = QLabel()
            value.setWordWrap(True)
            value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            grid.addWidget(value, row, 1)
            self.status_labels[key] = value
        root.addWidget(status_box)

        preview_box = QWidget()
        preview_layout = QVBoxLayout(preview_box)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        self.preview_label = QLabel("No wallpaper set.")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview_label.setMinimumSize(220, 180)
        preview_layout.addWidget(self.preview_label, 1)
        self.preview_name = QLabel("")
        self.preview_name.setAlignment(Qt.AlignCenter)
        name_font = QFont()
        name_font.setBold(True)
        name_font.setPointSize(name_font.pointSize() + 1)
        self.preview_name.setFont(name_font)
        preview_layout.addWidget(self.preview_name)
        root.addWidget(preview_box, 5)

        latest_box = QGroupBox(f"Latest {GALLERY_PAGE}")
        latest_layout = QGridLayout(latest_box)
        mono = QFont("Monospace")
        mono.setStyleHint(QFont.TypeWriter)
        mono.setPointSize(8)
        for col in range(GALLERY_PAGE):
            thumb = ClickableLabel()
            thumb.clicked.connect(lambda idx=col: self.on_thumbnail_clicked(idx))
            date_label = QLabel("")
            date_label.setAlignment(Qt.AlignCenter)
            date_label.setFont(mono)
            latest_layout.addWidget(thumb, 0, col)
            latest_layout.addWidget(date_label, 1, col)
            latest_layout.setColumnStretch(col, 1)
            self.thumb_slots.append((thumb, date_label))
        self.gallery_scroll = QScrollBar(Qt.Horizontal)
        self.gallery_scroll.setSingleStep(1)
        self.gallery_scroll.setPageStep(GALLERY_PAGE)
        self.gallery_scroll.valueChanged.connect(self.set_gallery_offset)
        latest_layout.addWidget(self.gallery_scroll, 2, 0, 1, GALLERY_PAGE)
        self.gallery_counter = QLabel("0/0")
        self.gallery_counter.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        latest_layout.addWidget(self.gallery_counter, 2, GALLERY_PAGE)
        latest_layout.setRowStretch(0, 1)
        root.addWidget(latest_box, 2)

    def _bind_shortcuts(self):
        QShortcut(QKeySequence(Qt.Key_Escape), self, activated=self.hide)
        QShortcut(QKeySequence(Qt.Key_F5), self, activated=self.run_check)
        QShortcut(QKeySequence("Ctrl+R"), self, activated=self.run_check)
        QShortcut(QKeySequence("Ctrl+,"), self, activated=self.show_settings)

    # ---- actions
    def run_check(self):
        threading.Thread(target=self.service.check_and_update, args=(True,), daemon=True).start()

    def show_settings(self):
        dialog = SettingsDialog(self.service, self)
        dialog.exec()

    def open_logs(self):
        self.service.open_path(self.paths.log_dir)

    def apply_wallpaper(self, path):
        """Apply a gallery image. The platform call may talk to D-Bus (and on
        GNOME the portal may show a confirmation), so keep it off the GUI."""
        def work():
            if self.service.set_wallpaper(path):
                self.service.set_status(f"Wallpaper updated: {Path(path).name}")
        threading.Thread(target=work, daemon=True).start()

    def on_thumbnail_clicked(self, slot_index):
        index = self.gallery_offset + slot_index
        if 0 <= index < len(self.images):
            self.apply_wallpaper(self.images[index])

    # ---- status
    def refresh_status(self):
        s = self.service
        minutes = s.interval_minutes()
        values = {
            "status": s.last_status_message,
            "last_check": s.format_ts(s.last_check),
            "last_success": s.format_ts(s.last_success),
            "last_error": (f"{s.format_ts(s.last_error_time)} - {s.last_error}"
                           if s.last_error_time else s.last_error),
            "current_image": s.current_image_path.name if s.current_image_path else "None",
            "interval": f"{minutes} minutes" if minutes > 0 else "Disabled",
            "startup": "Enabled" if s.is_startup_enabled() else "Disabled",
            "desktop": self.platform.describe(),
        }
        for key, label in self.status_labels.items():
            label.setText(values[key])

    # ---- preview
    def on_wallpaper_applied(self, path_str):
        self.refresh_preview(force_reload=True)
        self.refresh_status()

    def refresh_preview(self, force_reload=False):
        path = self.service.current_image_path
        if force_reload or path != self.preview_path:
            self.preview_source = None
            self.preview_path = path
        if not path or not Path(path).exists():
            self.preview_label.setPixmap(QPixmap())
            self.preview_label.setText("No wallpaper set.")
            self.preview_name.setText("")
            return
        if self.preview_source is None:
            self.preview_source = load_scaled_pixmap(path, PREVIEW_MAX)
        if self.preview_source is None:
            self.preview_label.setPixmap(QPixmap())
            self.preview_label.setText("Error displaying image")
            self.preview_name.setText("")
            return
        self.preview_name.setText(Path(path).name)
        self.redraw_preview()

    def redraw_preview(self):
        if self.preview_source is None:
            return
        avail = self.preview_label.size() - QSize(16, 8)
        if avail.width() < 50 or avail.height() < 50:
            return
        self.preview_label.setText("")
        self.preview_label.setPixmap(self.preview_source.scaled(avail, Qt.KeepAspectRatio, Qt.SmoothTransformation))

    # ---- gallery
    def reload_gallery(self):
        self.images = list(self.service.list_images_newest_first())
        stale = set(self.thumb_cache) - set(self.images)
        for path in stale:
            self.thumb_cache.pop(path, None)
        max_offset = max(len(self.images) - GALLERY_PAGE, 0)
        self.gallery_offset = min(self.gallery_offset, max_offset)
        self.gallery_scroll.blockSignals(True)
        self.gallery_scroll.setRange(0, max_offset)
        self.gallery_scroll.setValue(self.gallery_offset)
        self.gallery_scroll.blockSignals(False)
        self.render_gallery_page()

    def set_gallery_offset(self, offset):
        max_offset = max(len(self.images) - GALLERY_PAGE, 0)
        clamped = min(max(int(offset), 0), max_offset)
        if clamped != self.gallery_offset:
            self.gallery_offset = clamped
        if self.gallery_scroll.value() != clamped:
            self.gallery_scroll.setValue(clamped)
        self.render_gallery_page()

    def visible_images(self):
        return self.images[self.gallery_offset:self.gallery_offset + GALLERY_PAGE]

    def render_gallery_page(self):
        visible = self.visible_images()
        for idx, (thumb, date_label) in enumerate(self.thumb_slots):
            if idx >= len(visible):
                thumb.setPixmap(QPixmap())
                thumb.setText("No image")
                date_label.setText("")
                continue
            thumb.setText("")
            date_label.setText(display_date(visible[idx]))
        self.update_gallery_counter()
        self.redraw_thumbnails()

    def update_gallery_counter(self):
        total = len(self.images)
        if total == 0:
            self.gallery_counter.setText("0/0")
        elif total <= GALLERY_PAGE:
            self.gallery_counter.setText(f"{total}/{total}")
        else:
            start = self.gallery_offset + 1
            end = min(self.gallery_offset + GALLERY_PAGE, total)
            self.gallery_counter.setText(f"{start}-{end}/{total}")

    def thumb_source(self, path):
        if path not in self.thumb_cache:
            self.thumb_cache[path] = load_scaled_pixmap(path, THUMB_MAX)
        return self.thumb_cache[path]

    def redraw_thumbnails(self):
        for idx, path in enumerate(self.visible_images()):
            thumb, _ = self.thumb_slots[idx]
            source = self.thumb_source(path)
            if source is None:
                thumb.setPixmap(QPixmap())
                thumb.setText("Error")
                continue
            avail = thumb.size() - QSize(8, 8)
            if avail.width() < 20 or avail.height() < 20:
                continue
            thumb.setPixmap(source.scaled(avail, Qt.KeepAspectRatio, Qt.SmoothTransformation))

    # ---- window events
    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.preview_timer.start()
        self.thumb_timer.start()

    def showEvent(self, event):
        super().showEvent(event)
        self.preview_timer.start()
        self.thumb_timer.start()

    def closeEvent(self, event):
        # With a tray the window only hides (the tray keeps the app alive).
        # Without one (GNOME without an extension, odd sessions) the window
        # is the only handle on the app, so closing it quits for real.
        if self.tray_available:
            event.ignore()
            self.hide()
        else:
            event.accept()
            QApplication.instance().quit()


# ---------------------------------------------------------------- tray
class TrayController(QObject):
    """QSystemTrayIcon with the Windows menu: Preview / Gallery, Settings,
    Check Now, Check for Updates, Interval submenu, Exit."""

    def __init__(self, service, window, app):
        super().__init__()
        self.service = service
        self.window = window
        self.app = app
        self.tray = QSystemTrayIcon(fallback_icon(), self)
        self.tray.setToolTip(APP_DISPLAY_NAME)
        self.menu = QMenu()
        self.interval_actions = {}
        self._build_menu()
        self.tray.setContextMenu(self.menu)
        self.tray.activated.connect(self.on_activated)
        service.signals.wallpaper_applied.connect(self.update_icon)
        service.signals.state_changed.connect(self.sync_interval_checks)
        if service.current_image_path:
            self.update_icon(str(service.current_image_path))
        self.tray.show()

    def _build_menu(self):
        preview = QAction("Preview / Gallery", self.menu)
        preview.triggered.connect(self.show_window)
        self.menu.addAction(preview)
        self.menu.setDefaultAction(preview)
        settings = QAction("Settings", self.menu)
        settings.triggered.connect(self.show_settings)
        self.menu.addAction(settings)
        check = QAction("Check Now", self.menu)
        check.triggered.connect(self.window.run_check)
        self.menu.addAction(check)
        updates = QAction("Check for Updates", self.menu)
        updates.triggered.connect(self.check_updates)
        self.menu.addAction(updates)
        self.menu.addSeparator()

        interval_menu = self.menu.addMenu("Interval")
        group = QActionGroup(interval_menu)
        group.setExclusive(True)
        for label, minutes in INTERVAL_PRESETS.items():
            action = QAction(label, interval_menu)
            action.setCheckable(True)
            action.triggered.connect(lambda checked=False, m=minutes, l=label: self.set_interval(m, l))
            group.addAction(action)
            interval_menu.addAction(action)
            self.interval_actions[minutes] = action
        interval_menu.addSeparator()
        self.custom_action = QAction("Custom...", interval_menu)
        self.custom_action.setCheckable(True)
        self.custom_action.triggered.connect(self.custom_interval)
        group.addAction(self.custom_action)
        interval_menu.addAction(self.custom_action)
        self.sync_interval_checks()

        self.menu.addSeparator()
        quit_action = QAction("Exit", self.menu)
        quit_action.triggered.connect(self.exit_app)
        self.menu.addAction(quit_action)

    def sync_interval_checks(self):
        current = self.service.interval_minutes()
        is_custom = current not in self.interval_actions
        for minutes, action in self.interval_actions.items():
            action.setChecked(minutes == current)
        self.custom_action.setChecked(is_custom)
        self.custom_action.setText(f"Custom ({current} min)" if is_custom else "Custom...")

    def set_interval(self, minutes, label):
        log_msg(f"Setting interval to: {label} ({minutes} minutes)")
        self.service.set_interval(minutes)

    def custom_interval(self):
        self.sync_interval_checks()  # undo the click's check until confirmed
        minutes, ok = QInputDialog.getInt(None, "Custom Interval", "Enter check interval in minutes:",
                                          max(self.service.interval_minutes(), 1), 1, 10080)
        if ok:
            self.set_interval(minutes, f"Custom ({minutes} min)")

    def show_window(self):
        self.window.show()
        self.window.raise_()
        self.window.activateWindow()

    def show_settings(self):
        self.show_window()
        self.window.show_settings()

    def check_updates(self):
        threading.Thread(target=self.service.check_for_updates, args=(True,), daemon=True).start()

    def on_activated(self, reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self.show_window()

    def update_icon(self, path_str):
        pixmap = load_scaled_pixmap(Path(path_str), QSize(64, 64))
        if pixmap is not None:
            self.tray.setIcon(QIcon(pixmap))

    def notify(self, title, message):
        if self.tray.isVisible() and self.tray.supportsMessages():
            self.tray.showMessage(title, message, QSystemTrayIcon.Information, 5000)
            return True
        return False

    def exit_app(self):
        self.service.stop()
        self.tray.hide()
        self.app.quit()


# ---------------------------------------------------------------- application
class BingQtApp(QObject):
    """Wires service, window, tray, notifications and the update dialog."""

    def __init__(self, app, service, platform, paths):
        super().__init__()
        self.app = app
        self.service = service
        self.platform = platform
        self.paths = paths
        self.tray_available = QSystemTrayIcon.isSystemTrayAvailable()
        self.window = MainWindow(service, platform, paths, tray_available=self.tray_available)
        self.tray = TrayController(service, self.window, app) if self.tray_available else None
        service.signals.notification.connect(self.on_notification)
        service.signals.update_result.connect(self.on_update_result)
        service.signals.update_install_result.connect(self.on_update_install_result)
        self._pending_update = None

    def on_notification(self, title, message):
        if self.tray and self.tray.notify(title, message):
            return
        try:
            self.platform.notify(title, message)
        except Exception as e:
            log_msg(f"Notification failed: {e}", "error")

    def on_update_result(self, info, manual):
        if not manual:
            return
        if not info:
            QMessageBox.warning(self.window, "Update Check", "Unable to fetch release information.")
            return
        version = info["version"]
        if not info.get("newer"):
            QMessageBox.information(self.window, "Update Check", f"You are up to date (v{VERSION}).")
            return
        can_install = getattr(self.platform, "supports_self_update", False) and bool((info.get("asset") or {}).get("url"))
        if can_install:
            text = (f"A new version (v{version}) is available.\n\n"
                    "Download the Flatpak bundle, verify it and install it now? "
                    "The app restarts when the update is done.")
        elif self.platform.in_flatpak:
            text = (f"A new version (v{version}) is available, but this release has no Flatpak bundle "
                    "attached.\n\nOpen the release page to read what changed?")
        else:
            text = (f"A new version (v{version}) is available.\n\n"
                    "Open the release page to download it?")
        answer = QMessageBox.question(self.window, "Update Available", text,
                                      QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if answer != QMessageBox.Yes:
            return
        if can_install:
            self._pending_update = info
            updates_dir = Path(self.paths.data_dir) / "updates"
            threading.Thread(target=self.service.perform_update_install, args=(info, updates_dir),
                             name="update-install", daemon=True).start()
        elif info.get("html_url"):
            self.service.open_url(info["html_url"])

    def on_update_install_result(self, ok, detail):
        info = self._pending_update or {}
        self._pending_update = None
        if ok:
            QMessageBox.information(self.window, "Update",
                                    f"Update v{info.get('version', '')} installed. The app will now restart.")
            try:
                self.platform.restart_app()
            except Exception as e:
                log_msg(f"Could not schedule the restart: {e}", "error")
            self.service.stop()
            self.app.quit()
            return
        answer = QMessageBox.critical(self.window, "Update",
                                      f"The update could not be installed:\n{detail}\n\nOpen the release page instead?",
                                      QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if answer == QMessageBox.Yes and info.get("html_url"):
            self.service.open_url(info["html_url"])

    def start(self, hidden):
        threading.Thread(target=self.service.background_loop, name="check-loop", daemon=True).start()
        # No tray: the window is the only way to reach the app, so never start hidden.
        if hidden and self.tray_available:
            self.window.hide()
        else:
            self.window.show()

    def show_window(self):
        if self.tray:
            self.tray.show_window()
        else:
            self.window.show()
            self.window.raise_()
            self.window.activateWindow()


class SingleInstance(QObject):
    """QLocalServer named after the app id. A second launch sends 'show' and
    exits; the first instance raises its window."""

    def __init__(self, name, on_show):
        super().__init__()
        self.name = name
        self.on_show = on_show
        self.server = None

    @staticmethod
    def signal_existing(name, timeout_ms=500):
        socket = QLocalSocket()
        socket.connectToServer(name)
        if not socket.waitForConnected(timeout_ms):
            return False
        socket.write(b"show")
        socket.waitForBytesWritten(timeout_ms)
        socket.disconnectFromServer()
        return True

    def listen(self):
        QLocalServer.removeServer(self.name)
        self.server = QLocalServer(self)
        self.server.newConnection.connect(self._on_connection)
        return self.server.listen(self.name)

    def _on_connection(self):
        while self.server.hasPendingConnections():
            conn = self.server.nextPendingConnection()
            conn.readyRead.connect(lambda c=conn: self._on_ready(c))

    def _on_ready(self, conn):
        data = bytes(conn.readAll())
        conn.disconnectFromServer()
        if b"show" in data:
            self.on_show()


def parse_args(argv):
    parser = argparse.ArgumentParser(prog=LAUNCHER_NAME, description=f"{APP_DISPLAY_NAME} v{VERSION}")
    parser.add_argument("--hidden", action="store_true", help="start minimized to the tray")
    parser.add_argument("--version", action="version", version=f"{APP_DISPLAY_NAME} {VERSION}")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(sys.argv[1:] if argv is None else argv)
    paths = linux_paths()
    configure_logging(paths.log_dir)

    app = QApplication(sys.argv)
    app.setApplicationName(APP_DISPLAY_NAME)
    app.setApplicationVersion(VERSION)
    app.setDesktopFileName(APP_ID)
    app.setQuitOnLastWindowClosed(False)
    app.setWindowIcon(QIcon.fromTheme(APP_ID, fallback_icon()))

    if SingleInstance.signal_existing(APP_ID):
        print("Application is already running.")
        return 0

    platform = LinuxPlatform(paths, launch_command=launch_command(bool(os.environ.get("FLATPAK_ID"))))
    service = QtWallpaperService(platform, paths)
    controller = BingQtApp(app, service, platform, paths)
    instance = SingleInstance(APP_ID, controller.show_window)
    if not instance.listen():
        log_msg("Single-instance server could not listen; continuing without it", "error")

    controller.start(hidden=args.hidden)
    try:
        return app.exec()
    finally:
        service.stop()


if __name__ == "__main__":
    sys.exit(main())
