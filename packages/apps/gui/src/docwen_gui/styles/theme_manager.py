"""ThemeManager singleton for DocWen GUI.

Applies theme changes at runtime: QPalette construction, qfluentwidgets
sync, and global QSS stylesheet application.
"""

from __future__ import annotations

import logging
from typing import Literal
from typing import cast as _cast

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication, QWidget

from docwen_gui.font_utils import apply_application_font, normalize_font_size_preset

from .global_aggregate import build_global_stylesheet
from .theme_semantics import DEFAULT_THEME

logger = logging.getLogger(__name__)

ThemeName = Literal["light", "dark", "system"]


class ThemeManager:
    """Singleton that manages application-wide theme state."""

    _instance: ThemeManager | None = None

    def __init__(self) -> None:
        self._current_theme: str = DEFAULT_THEME
        self._font_size_preset: str = "default"
        self._app: QApplication | None = None

    @classmethod
    def get_instance(cls) -> ThemeManager:
        """Return the singleton instance, creating it if necessary."""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    @classmethod
    def reset_instance(cls) -> None:
        """Reset singleton instance (for testing)."""
        cls._instance = None

    def initialize(self, app: QApplication, theme_name: str) -> None:
        """One-time initialization with QApplication instance.

        Args:
            app: The QApplication instance.
            theme_name: Initial theme name (``"light"``, ``"dark"``, or ``"system"``).
        """
        self._app = app
        self.apply_theme(theme_name)

    def apply_theme(self, theme_name: str) -> None:
        """Apply theme at runtime.  Safe to call multiple times.

        Args:
            theme_name: One of ``"light"``, ``"dark"``, ``"system"``.
        """
        resolved = self._resolve_system_theme(theme_name)
        self._current_theme = resolved

        if self._app is None:
            logger.warning("ThemeManager.apply_theme called before initialize()")
            return

        # 1. Build and apply QPalette
        palette = self._build_application_palette(resolved)
        self._app.setPalette(palette)

        # 2. Sync qfluentwidgets theme + accent color (best-effort).
        # Fluent widgets default to teal #009FAA; pin the accent to the
        # palette Highlight so every widget family shares one accent color.
        try:
            import qfluentwidgets
            from qfluentwidgets import Theme as FluentTheme

            qfluentwidgets.setTheme(FluentTheme.DARK if resolved == "dark" else FluentTheme.LIGHT)
            qfluentwidgets.setThemeColor(palette.color(QPalette.ColorRole.Highlight))
        except ImportError:
            pass

        # 3. Apply global QSS stylesheet
        stylesheet = build_global_stylesheet(resolved, self._font_size_preset)
        self._app.setStyleSheet(stylesheet)

        # 4. Force repaint for visible top-level widgets. Hidden dialogs pick
        # up the global palette/stylesheet when shown and do not need an eager
        # polish pass.
        for widget in self._app.topLevelWidgets():
            if not widget.isVisible():
                continue
            widget.style().unpolish(widget)
            widget.style().polish(widget)
            widget.update()
            for child in widget.findChildren(QWidget):
                child.update()

    def get_current_theme(self) -> str:
        """Return the currently active (resolved) theme name."""
        return self._current_theme

    def apply_font_size_preset(self, preset: str | None) -> str:
        """Apply one semantic typography preset to existing and future UI."""
        normalized = normalize_font_size_preset(preset)
        self._font_size_preset = normalized

        if self._app is None:
            app = QApplication.instance()
            if isinstance(app, QApplication):
                self._app = app
        if self._app is None:
            logger.warning("ThemeManager.apply_font_size_preset called without QApplication")
            return normalized

        apply_application_font(self._app, font_size_preset=normalized)
        self.apply_theme(self._current_theme)
        return normalized

    def get_font_size_preset(self) -> str:
        return self._font_size_preset

    @staticmethod
    def get_available_themes() -> list[str]:
        """Return the list of supported theme names."""
        return ["light", "dark", "system"]

    @staticmethod
    def _resolve_system_theme(theme_name: str) -> str:
        """Resolve ``"system"`` to the actual OS color scheme.

        Falls back to ``"light"`` when detection is unavailable.
        """
        if theme_name == "system":
            try:
                app = QApplication.instance()
                if app is not None:
                    is_dark = _cast(QApplication, app).styleHints().colorScheme() == Qt.ColorScheme.Dark
                    return "dark" if is_dark else "light"
            except Exception:
                logger.debug("Failed to read system color scheme; falling back to light")
            return "light"
        if theme_name in ("light", "dark"):
            return theme_name
        logger.warning("Unknown theme name %r; falling back to %r", theme_name, DEFAULT_THEME)
        return DEFAULT_THEME

    @staticmethod
    def _build_application_palette(theme: str) -> QPalette:
        """Build a full QPalette for the given theme.

        Colors are chosen from the Slate colour scale to match the
        project's design tokens.
        """
        palette = QPalette()
        is_dark = theme == "dark"

        if is_dark:
            window = QColor("#0F172A")
            window_text = QColor("#F1F5F9")
            base = QColor("#1E293B")
            alternate_base = QColor("#334155")
            text = QColor("#F1F5F9")
            button = QColor("#1E293B")
            button_text = QColor("#F1F5F9")
            highlight = QColor("#3B82F6")
            highlighted_text = QColor("#FFFFFF")
            placeholder = QColor("#64748B")
            mid = QColor("#64748B")
            midlight = QColor("#334155")
        else:
            window = QColor("#F8FAFC")
            window_text = QColor("#0F172A")
            base = QColor("#FFFFFF")
            alternate_base = QColor("#F1F5F9")
            text = QColor("#0F172A")
            button = QColor("#FFFFFF")
            button_text = QColor("#0F172A")
            highlight = QColor("#3B82F6")
            highlighted_text = QColor("#FFFFFF")
            placeholder = QColor("#94A3B8")
            mid = QColor("#94A3B8")
            midlight = QColor("#CBD5E1")

        # ── Active & Inactive ──────────────────────────────────────
        role_map: dict[QPalette.ColorRole, QColor] = {
            QPalette.ColorRole.Window: window,
            QPalette.ColorRole.WindowText: window_text,
            QPalette.ColorRole.Base: base,
            QPalette.ColorRole.AlternateBase: alternate_base,
            QPalette.ColorRole.ToolTipBase: base,
            QPalette.ColorRole.ToolTipText: text,
            QPalette.ColorRole.Text: text,
            QPalette.ColorRole.Button: button,
            QPalette.ColorRole.ButtonText: button_text,
            QPalette.ColorRole.BrightText: QColor("#FFFFFF"),
            QPalette.ColorRole.Highlight: highlight,
            QPalette.ColorRole.HighlightedText: highlighted_text,
            QPalette.ColorRole.PlaceholderText: placeholder,
            QPalette.ColorRole.Mid: mid,
            QPalette.ColorRole.Midlight: midlight,
        }
        for role, color in role_map.items():
            palette.setColor(QPalette.ColorGroup.Active, role, color)
            palette.setColor(QPalette.ColorGroup.Inactive, role, color)

        # ── Disabled ───────────────────────────────────────────────
        disabled_text = QColor(text)
        disabled_text.setAlpha(96 if is_dark else 110)
        disabled_button_text = QColor(text)
        disabled_button_text.setAlpha(92 if is_dark else 118)
        disabled_highlight = QColor(highlight)
        disabled_highlight.setAlpha(84 if is_dark else 72)

        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, disabled_text)
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, disabled_text)
        palette.setColor(
            QPalette.ColorGroup.Disabled,
            QPalette.ColorRole.ButtonText,
            disabled_button_text,
        )
        palette.setColor(
            QPalette.ColorGroup.Disabled,
            QPalette.ColorRole.Highlight,
            disabled_highlight,
        )
        palette.setColor(
            QPalette.ColorGroup.Disabled,
            QPalette.ColorRole.HighlightedText,
            highlighted_text,
        )
        palette.setColor(
            QPalette.ColorGroup.Disabled,
            QPalette.ColorRole.PlaceholderText,
            placeholder,
        )

        # Non-text roles retain their normal colour when disabled
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Window, window)
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Base, base)
        palette.setColor(
            QPalette.ColorGroup.Disabled,
            QPalette.ColorRole.AlternateBase,
            alternate_base,
        )
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Button, button)
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Mid, mid)
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Midlight, midlight)

        return palette
