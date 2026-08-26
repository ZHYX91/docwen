"""GUI resource helpers — path resolution and icon management.

Provides access to bundle assets (icons, images) and SVG/PNG icon
loading with theme-aware color adaptation.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Any

from PySide6.QtCore import QByteArray, QRectF, QSize, Qt
from PySide6.QtGui import QIcon, QIconEngine, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer

from docwen_runtime.resources import ResourceRegistry

logger = logging.getLogger(__name__)

# ── Path helpers ──────────────────────────────────────────────────────────


def app_root() -> Path:
    return ResourceRegistry.default().root


def asset_path(rel_path: str) -> Path:
    return ResourceRegistry.default().assets_dir() / rel_path


def i18n_locales_dir() -> Path:
    return ResourceRegistry.default().locales_dir()


# ── Icon resolution ───────────────────────────────────────────────────────


def get_icon_path() -> str | None:
    """Return the path to the application icon, preferring .ico on Windows."""
    if sys.platform == "win32":
        ico = asset_path("icon.ico")
        if ico.exists():
            return str(ico)
    png = asset_path("icon.png")
    if png.exists():
        return str(png)
    return None


# ── SVG icon loading ──────────────────────────────────────────────────────

_icon_cache: dict[str, Any] = {}


class _SvgIconEngine(QIconEngine):
    """Render SVG bytes at the device-pixel ratio requested by Qt."""

    def __init__(self, svg_bytes: bytes) -> None:
        super().__init__()
        self._svg_bytes = bytes(svg_bytes)

    def clone(self) -> _SvgIconEngine:
        return _SvgIconEngine(self._svg_bytes)

    def key(self) -> str:
        return "DocWenSvgIconEngine"

    def isNull(self) -> bool:
        return not QSvgRenderer(QByteArray(self._svg_bytes)).isValid()

    def actualSize(self, size: QSize, mode: QIcon.Mode, state: QIcon.State) -> QSize:
        del mode, state
        return QSize(size)

    def paint(self, painter: QPainter, rect, mode: QIcon.Mode, state: QIcon.State) -> None:
        del mode, state
        renderer = QSvgRenderer(QByteArray(self._svg_bytes))
        if renderer.isValid():
            renderer.render(painter, QRectF(rect))

    def pixmap(self, size: QSize, mode: QIcon.Mode, state: QIcon.State) -> QPixmap:
        del mode, state
        return self._render_pixmap(size, 1.0)

    def scaledPixmap(self, size: QSize, mode: QIcon.Mode, state: QIcon.State, scale: float) -> QPixmap:
        del mode, state
        return self._render_pixmap(size, scale)

    def _render_pixmap(self, size: QSize, scale: float) -> QPixmap:
        effective_scale = max(float(scale), 1.0)
        physical_size = QSize(
            max(1, round(size.width() * effective_scale)),
            max(1, round(size.height() * effective_scale)),
        )
        pixmap = QPixmap(physical_size)
        pixmap.fill(Qt.GlobalColor.transparent)
        pixmap.setDevicePixelRatio(effective_scale)
        painter = QPainter(pixmap)
        try:
            renderer = QSvgRenderer(QByteArray(self._svg_bytes))
            if renderer.isValid():
                renderer.render(painter, QRectF(0, 0, size.width(), size.height()))
        finally:
            painter.end()
        return pixmap


def _resolve_theme_icon_color() -> str:
    """Determine the current theme's text color for SVG fill adaptation."""
    try:
        from PySide6.QtGui import QPalette
        from PySide6.QtWidgets import QApplication

        app = QApplication.instance()
        if isinstance(app, QApplication):
            color = app.palette().color(QPalette.ColorRole.WindowText)
            return str(color.name())
    except Exception:
        pass
    return "#000000"


def load_svg_asset_icon(
    relative_path: str,
    *,
    color: str | None = None,
) -> Any | None:
    """Load an SVG asset through a resolution-independent Qt icon engine."""
    resolved_color = color or _resolve_theme_icon_color()
    cache_key = f"svg:{relative_path}:{resolved_color}"
    if cache_key in _icon_cache:
        return _icon_cache[cache_key]

    icon_path = asset_path(relative_path)
    if not icon_path.exists():
        logger.debug("SVG icon not found: %s", icon_path)
        return None

    try:
        svg_text = icon_path.read_text(encoding="utf-8", errors="ignore")

        # Adapt the monochrome application icon vocabulary to the active theme.
        for attribute in ("fill", "stroke"):
            for old_color in ("#000", "#000000", "black"):
                svg_text = svg_text.replace(
                    f'{attribute}="{old_color}"',
                    f'{attribute}="{resolved_color}"',
                )

        svg_bytes = svg_text.encode("utf-8")
        renderer = QSvgRenderer(QByteArray(svg_bytes))
        if not renderer.isValid():
            logger.debug("Invalid SVG icon: %s", icon_path)
            return None

        icon = QIcon(_SvgIconEngine(svg_bytes))
        _icon_cache[cache_key] = icon
        return icon
    except Exception as exc:
        logger.error("Failed to load SVG icon %s: %s", relative_path, exc)
        return None


def load_svg_icon(
    icon_name: str,
    *,
    color: str | None = None,
) -> Any | None:
    """Load a theme-aware vector icon from ``assets/icons``.

    Args:
        icon_name: Filename under assets/icons/ (e.g. ``"settings.svg"``).
        color: Override fill color; if None, auto-detects from theme.

    Returns:
        A QIcon, or None on failure.
    """
    return load_svg_asset_icon(f"icons/{icon_name}", color=color)


def load_image_icon(
    icon_name: str,
    size: tuple[int, int] | None = None,
) -> Any | None:
    """Load a raster icon (PNG/ICO) from assets/, optionally scaled.

    Args:
        icon_name: Filename under assets/ (e.g. ``"complete_icon.png"``).
        size: Optional target pixel size ``(width, height)`` for scaling.

    Returns:
        A QIcon or QPixmap, or None on failure.
    """
    cache_key = f"{icon_name}_{size}" if size else icon_name
    if cache_key in _icon_cache:
        return _icon_cache[cache_key]

    icon_path = asset_path(icon_name)
    if not icon_path.exists():
        logger.debug("Raster icon not found: %s", icon_path)
        return None

    try:
        from PySide6.QtCore import QSize
        from PySide6.QtGui import QIcon, QPixmap

        if size:
            pixmap = QPixmap(str(icon_path))
            if pixmap.isNull():
                return None
            scaled = pixmap.scaled(QSize(size[0], size[1]))
            _icon_cache[cache_key] = scaled
            return scaled

        icon = QIcon(str(icon_path))
        _icon_cache[cache_key] = icon
        return icon
    except Exception as exc:
        logger.error("Failed to load icon %s: %s", icon_name, exc)
        return None


# ── Window icon singleton ─────────────────────────────────────────────────


class _IconManager:
    """Manages the application window icon.

    Initialises once per process lifetime, setting the icon on both
    the QApplication and the root window.
    """

    _icon_path: str | None = None
    _initialized: bool = False

    @classmethod
    def initialize(cls, root_window: Any) -> bool:
        """Set the application window icon from the bundle assets.

        Args:
            root_window: The main window widget (must have ``setWindowIcon``).

        Returns:
            True if the icon was set successfully.
        """
        if cls._initialized:
            return True

        icon_path = get_icon_path()
        if not icon_path:
            logger.debug("No application icon found in assets")
            return False

        try:
            from PySide6.QtGui import QIcon
            from PySide6.QtWidgets import QApplication

            icon = QIcon(icon_path)
            if icon.isNull():
                logger.debug("Icon file could not be loaded: %s", icon_path)
                return False

            app = QApplication.instance()
            if isinstance(app, QApplication):
                app.setWindowIcon(icon)
            if hasattr(root_window, "setWindowIcon"):
                root_window.setWindowIcon(icon)

            cls._icon_path = icon_path
            cls._initialized = True
            logger.debug("Window icon set: %s", icon_path)
            return True
        except Exception as exc:
            logger.error("Failed to set application icon: %s", exc)
            return False

    @classmethod
    def get_icon_path(cls) -> str | None:
        """Return the resolved icon path, or None."""
        if cls._icon_path is None:
            cls._icon_path = get_icon_path()
        return cls._icon_path


def initialize_app_icon(root_window: Any) -> bool:
    """Convenience: initialise the application window icon.

    Wraps ``_IconManager.initialize()``.
    """
    return _IconManager.initialize(root_window)


def get_app_icon_path() -> str | None:
    """Return the resolved application icon path."""
    return _IconManager.get_icon_path()
