"""Application font selection for the GUI."""

from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)

DEFAULT_FONT_FAMILIES = [
    "Microsoft YaHei",
    "PingFang SC",
    "Noto Sans CJK SC",
    "Source Han Sans SC",
    "Malgun Gothic",
    "Segoe UI",
    "Arial",
    "Helvetica",
    "sans-serif",
]
DEFAULT_FONT_SIZE = 12
FONT_SIZE_PRESETS: dict[str, int] = {
    "small": 11,
    "default": DEFAULT_FONT_SIZE,
    "large": 13,
    "xlarge": 15,
}
FONT_SIZE_PRESET_DELTAS: dict[str, int] = {name: size - DEFAULT_FONT_SIZE for name, size in FONT_SIZE_PRESETS.items()}

_FONT_ALIASES = {
    "Microsoft YaHei": ["Microsoft YaHei", "微软雅黑", "YaHei"],
    "PingFang SC": ["PingFang SC", "苹方", "PingFang"],
    "Noto Sans CJK SC": ["Noto Sans CJK SC", "思源黑体", "Source Han Sans", "Noto Sans"],
    "Source Han Sans SC": ["Source Han Sans SC", "Source Han Sans", "思源黑体", "Noto Sans CJK SC"],
    "Malgun Gothic": ["Malgun Gothic", "MalgunGothic", "맑은 고딕"],
    "Segoe UI": ["Segoe UI", "SegoeUI", "Segoe UI Variable"],
    "Arial": ["Arial", "Arial MT"],
    "Helvetica": ["Helvetica", "Helvetica Neue"],
    "sans-serif": ["sans-serif", "Sans Serif"],
}

_system_fonts_cache: list[str] | None = None
_loaded_cjk_fonts = False


def get_system_fonts() -> list[str]:
    global _system_fonts_cache
    if _system_fonts_cache is not None:
        return _system_fonts_cache

    try:
        from PySide6.QtGui import QFontDatabase
        from PySide6.QtWidgets import QApplication

        fonts = list(QFontDatabase.families())
        if not fonts or get_available_font_from_list(fonts) is None:
            _load_known_cjk_font_files()
            fonts = list(QFontDatabase.families())
        if not fonts:
            app = QApplication.instance()
            if isinstance(app, QApplication):
                family = app.font().family()
                if family:
                    fonts = [family]
    except Exception as exc:  # pragma: no cover - defensive platform fallback
        logger.debug("Failed to query GUI fonts: %s", exc)
        fonts = []

    _system_fonts_cache = fonts
    return fonts


def get_available_font_from_list(system_fonts: list[str], font_families: list[str] | None = None) -> str | None:
    preferred = font_families or DEFAULT_FONT_FAMILIES
    for font in preferred:
        if font in system_fonts:
            return font
        for alias in _FONT_ALIASES.get(font, []):
            if alias in system_fonts:
                return alias
        font_lower = font.lower()
        for system_font in system_fonts:
            if font_lower in system_font.lower():
                return system_font
    return None


def get_available_font(font_families: list[str] | None = None) -> str | None:
    preferred = font_families or DEFAULT_FONT_FAMILIES
    system_fonts = get_system_fonts()
    if not system_fonts:
        return None
    return get_available_font_from_list(system_fonts, preferred)


def _load_known_cjk_font_files() -> None:
    global _loaded_cjk_fonts
    if _loaded_cjk_fonts:
        return
    _loaded_cjk_fonts = True
    try:
        from PySide6.QtGui import QFontDatabase
    except Exception:  # pragma: no cover - PySide import guard
        return

    for path in _known_cjk_font_files():
        if path.is_file():
            QFontDatabase.addApplicationFont(str(path))


def _known_cjk_font_files() -> list[Path]:
    candidates: list[Path] = []
    windows_fonts = Path("C:/Windows/Fonts")
    candidates.extend(
        [
            windows_fonts / "msyh.ttc",
            windows_fonts / "msyhbd.ttc",
            windows_fonts / "NotoSansSC-VF.ttf",
            windows_fonts / "simsun.ttc",
            windows_fonts / "simhei.ttf",
            windows_fonts / "malgun.ttf",
            windows_fonts / "malgunbd.ttf",
        ]
    )
    for root in (
        Path("/usr/share/fonts"),
        Path("/usr/local/share/fonts"),
        Path("/System/Library/Fonts"),
        Path("/Library/Fonts"),
    ):
        if root.is_dir():
            for pattern in (
                "*NotoSansCJK*.ttc",
                "*NotoSansSC*.ttf",
                "*SourceHanSans*.otf",
                "*PingFang*.ttc",
            ):
                candidates.extend(root.rglob(pattern))
    return candidates


def normalize_font_size_preset(preset: str | None) -> str:
    normalized = str(preset or "").strip().lower()
    if normalized in {"extra_large", "extra-large"}:
        normalized = "xlarge"
    return normalized if normalized in FONT_SIZE_PRESETS else "default"


def resolve_font_size_preset(preset: str | None) -> int:
    return FONT_SIZE_PRESETS[normalize_font_size_preset(preset)]


def resolve_typography_size(default_size: int, preset: str | None) -> int:
    """Apply a font preset as a delta while preserving semantic hierarchy."""
    normalized = normalize_font_size_preset(preset)
    return max(9, int(default_size) + FONT_SIZE_PRESET_DELTAS[normalized])


def apply_application_font(app: object, *, font_size_preset: str | None = None) -> None:
    try:
        from PySide6.QtGui import QFont
        from PySide6.QtWidgets import QApplication
    except Exception:  # pragma: no cover - PySide import guard
        return

    if not isinstance(app, QApplication):
        return

    family = get_available_font()
    current_font = app.font()
    font = QFont(current_font)
    if family:
        font.setFamily(family)
    font.setPointSize(resolve_font_size_preset(font_size_preset))
    app.setFont(font)
