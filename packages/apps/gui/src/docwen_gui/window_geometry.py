"""Pure schema-v2 main-window geometry and recovery policy."""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from math import isfinite
from typing import Literal, Protocol

GEOMETRY_SCHEMA_VERSION = 2

DEFAULT_CENTER_PANEL_SCREEN_X = 420
DEFAULT_WINDOW_Y = 0
DEFAULT_WINDOW_WIDTH = 476
DEFAULT_WINDOW_HEIGHT = 860
DEFAULT_MIN_WIDTH = 420
DEFAULT_MIN_HEIGHT = 720

GeometrySource = Literal["canonical", "startup-default"]


class GeometryConfigReader(Protocol):
    """The narrow injected config surface required by geometry projection."""

    def get(self, key: str, default: object = None) -> object: ...


@dataclass(frozen=True, slots=True)
class WindowRect:
    """A normal top-level window rectangle in runtime coordinates."""

    x: int
    y: int
    width: int
    height: int


@dataclass(frozen=True, slots=True)
class ScreenRect:
    """A screen work area in the same coordinate space as ``WindowRect``."""

    x: int
    y: int
    width: int
    height: int

    @property
    def right(self) -> int:
        return self.x + self.width

    @property
    def bottom(self) -> int:
        return self.y + self.height


@dataclass(frozen=True, slots=True)
class WindowGeometryPolicy:
    """Validated startup geometry projected from the current schema."""

    rect: WindowRect
    min_width: int
    min_height: int
    source: GeometrySource
    schema_version: int | None
    schema_supported: bool


@dataclass(frozen=True, slots=True)
class RecoveredWindowGeometry:
    """A window rectangle fitted to one currently available screen."""

    rect: WindowRect
    screen: ScreenRect | None
    effective_min_width: int
    effective_min_height: int
    disconnected: bool


_MISSING = object()


def _config_get(
    config_port: GeometryConfigReader | None,
    key: str,
    default: object,
) -> object:
    if config_port is None:
        return default
    try:
        return config_port.get(key, default)
    except Exception:
        return default


def _finite_int(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        numeric = int(value)
        return numeric if -2_147_483_648 <= numeric <= 2_147_483_647 else None
    if isinstance(value, float):
        if not isfinite(value):
            return None
        if value < -2_147_483_648 or value > 2_147_483_647:
            return None
        return round(value)
    return None


def _positive_int(value: object) -> int | None:
    numeric = _finite_int(value)
    return numeric if numeric is not None and numeric > 0 else None


def _read_int(config_port: GeometryConfigReader | None, key: str, fallback: int) -> int:
    value = _finite_int(_config_get(config_port, key, fallback))
    return fallback if value is None else value


def _read_positive_int(
    config_port: GeometryConfigReader | None,
    key: str,
    fallback: int,
) -> int:
    value = _positive_int(_config_get(config_port, key, fallback))
    return fallback if value is None else value


def _scale_int(scale_value: Callable[[int], int], value: int) -> int:
    try:
        scaled = _finite_int(scale_value(value))
    except Exception:
        scaled = None
    return value if scaled is None else scaled


def _unscale_int(unscale_value: Callable[[int], int], value: int) -> int:
    try:
        logical = _finite_int(unscale_value(value))
    except Exception:
        logical = None
    return value if logical is None else logical


def _read_schema_version(
    config_port: GeometryConfigReader | None,
) -> tuple[int | None, bool]:
    if config_port is None:
        return GEOMETRY_SCHEMA_VERSION, True
    try:
        raw_version = config_port.get(
            "gui.window.geometry_schema_version",
            _MISSING,
        )
    except Exception:
        return None, False
    if raw_version is _MISSING:
        return None, False
    if isinstance(raw_version, bool) or not isinstance(raw_version, int) or raw_version <= 0:
        return None, False
    return raw_version, raw_version == GEOMETRY_SCHEMA_VERSION


def load_window_geometry_policy(
    config_port: GeometryConfigReader | None,
    *,
    center_offset: int,
    scale_value: Callable[[int], int],
) -> WindowGeometryPolicy:
    """Load schema-v2 geometry and fail safely for unsupported versions."""

    schema_version, version_supported = _read_schema_version(config_port)
    schema_supported = version_supported

    logical_min_width = (
        _read_positive_int(config_port, "gui.window.min_width", DEFAULT_MIN_WIDTH)
        if schema_supported
        else DEFAULT_MIN_WIDTH
    )
    logical_min_height = (
        _read_positive_int(config_port, "gui.window.min_height", DEFAULT_MIN_HEIGHT)
        if schema_supported
        else DEFAULT_MIN_HEIGHT
    )
    min_width = _scale_int(scale_value, logical_min_width)
    min_height = _scale_int(scale_value, logical_min_height)
    min_width = max(1, min_width)
    min_height = max(1, min_height)

    if schema_supported:
        logical_anchor_x = _read_int(
            config_port,
            "gui.window.center_panel_screen_x",
            DEFAULT_CENTER_PANEL_SCREEN_X,
        )
        logical_y = _read_int(config_port, "gui.window.window_y", DEFAULT_WINDOW_Y)
        logical_width = _read_positive_int(
            config_port,
            "gui.window.default_width",
            DEFAULT_WINDOW_WIDTH,
        )
        logical_height = _read_positive_int(
            config_port,
            "gui.window.default_height",
            DEFAULT_WINDOW_HEIGHT,
        )
    else:
        logical_anchor_x = DEFAULT_CENTER_PANEL_SCREEN_X
        logical_y = DEFAULT_WINDOW_Y
        logical_width = DEFAULT_WINDOW_WIDTH
        logical_height = DEFAULT_WINDOW_HEIGHT
    anchor_x = _scale_int(scale_value, logical_anchor_x)
    y = _scale_int(scale_value, logical_y)
    width = max(1, _scale_int(scale_value, logical_width))
    height = max(1, _scale_int(scale_value, logical_height))
    return WindowGeometryPolicy(
        rect=WindowRect(
            x=anchor_x - int(center_offset),
            y=y,
            width=width,
            height=height,
        ),
        min_width=min_width,
        min_height=min_height,
        source="canonical",
        schema_version=schema_version,
        schema_supported=schema_supported,
    )


def build_canonical_geometry_values(
    rect: WindowRect,
    *,
    center_offset: int,
    unscale_value: Callable[[int], int],
) -> dict[str, object]:
    """Build the single sparse write used to persist normal window geometry."""

    return {
        "gui.window.geometry_schema_version": GEOMETRY_SCHEMA_VERSION,
        "gui.window.center_panel_screen_x": _unscale_int(
            unscale_value,
            rect.x + int(center_offset),
        ),
        "gui.window.window_y": _unscale_int(unscale_value, rect.y),
        "gui.window.default_width": max(1, _unscale_int(unscale_value, rect.width)),
        "gui.window.default_height": max(1, _unscale_int(unscale_value, rect.height)),
    }


def _intersection_area(left: WindowRect, right: ScreenRect) -> int:
    width = max(0, min(left.x + left.width, right.right) - max(left.x, right.x))
    height = max(0, min(left.y + left.height, right.bottom) - max(left.y, right.y))
    return width * height


def _distance_squared_to_screen(rect: WindowRect, screen: ScreenRect) -> int:
    center_x_twice = rect.x * 2 + rect.width
    center_y_twice = rect.y * 2 + rect.height
    nearest_x_twice = min(max(center_x_twice, screen.x * 2), screen.right * 2)
    nearest_y_twice = min(max(center_y_twice, screen.y * 2), screen.bottom * 2)
    delta_x = center_x_twice - nearest_x_twice
    delta_y = center_y_twice - nearest_y_twice
    return delta_x * delta_x + delta_y * delta_y


def _valid_screens(screens: Sequence[ScreenRect]) -> tuple[ScreenRect, ...]:
    return tuple(screen for screen in screens if screen.width > 0 and screen.height > 0)


def recover_window_geometry(
    candidate: WindowRect,
    screens: Sequence[ScreenRect],
    *,
    min_width: int,
    min_height: int,
) -> RecoveredWindowGeometry:
    """Fit a startup rectangle into the best current screen work area.

    A candidate that intersects a screen keeps its screen and is clamped to the
    work area. A candidate from a disconnected display has zero intersection;
    it is centered on the nearest remaining work area instead of being left at
    an unreachable edge. Negative-coordinate displays remain valid.
    """

    normalized_min_width = max(1, int(min_width))
    normalized_min_height = max(1, int(min_height))
    normalized_candidate = WindowRect(
        x=int(candidate.x),
        y=int(candidate.y),
        width=max(normalized_min_width, int(candidate.width)),
        height=max(normalized_min_height, int(candidate.height)),
    )
    available = _valid_screens(screens)
    if not available:
        return RecoveredWindowGeometry(
            rect=normalized_candidate,
            screen=None,
            effective_min_width=normalized_min_width,
            effective_min_height=normalized_min_height,
            disconnected=False,
        )

    intersections = tuple(_intersection_area(normalized_candidate, screen) for screen in available)
    largest_area = max(intersections)
    disconnected = largest_area == 0
    if disconnected:
        selected = min(
            available,
            key=lambda screen: _distance_squared_to_screen(normalized_candidate, screen),
        )
    else:
        selected = available[intersections.index(largest_area)]

    effective_min_width = min(normalized_min_width, selected.width)
    effective_min_height = min(normalized_min_height, selected.height)
    width = min(max(normalized_candidate.width, effective_min_width), selected.width)
    height = min(max(normalized_candidate.height, effective_min_height), selected.height)

    if disconnected:
        x = selected.x + (selected.width - width) // 2
        y = selected.y + (selected.height - height) // 2
    else:
        x = min(max(normalized_candidate.x, selected.x), selected.right - width)
        y = min(max(normalized_candidate.y, selected.y), selected.bottom - height)

    return RecoveredWindowGeometry(
        rect=WindowRect(x=x, y=y, width=width, height=height),
        screen=selected,
        effective_min_width=effective_min_width,
        effective_min_height=effective_min_height,
        disconnected=disconnected,
    )


def center_window_geometry(
    *,
    width: int,
    height: int,
    screen: ScreenRect,
    min_width: int,
    min_height: int,
) -> RecoveredWindowGeometry:
    """Center a size on one work area while applying the same size caps."""

    effective_min_width = min(max(1, int(min_width)), screen.width)
    effective_min_height = min(max(1, int(min_height)), screen.height)
    fitted_width = min(max(int(width), effective_min_width), screen.width)
    fitted_height = min(max(int(height), effective_min_height), screen.height)
    rect = WindowRect(
        x=screen.x + (screen.width - fitted_width) // 2,
        y=screen.y + (screen.height - fitted_height) // 2,
        width=fitted_width,
        height=fitted_height,
    )
    return RecoveredWindowGeometry(
        rect=rect,
        screen=screen,
        effective_min_width=effective_min_width,
        effective_min_height=effective_min_height,
        disconnected=False,
    )


def normalize_ui_scale(value: object) -> float | None:
    """Normalize a factor or percentage to a bounded UI scale."""

    if isinstance(value, bool):
        return None
    percentage = False
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        if text.endswith("%"):
            percentage = True
            text = text[:-1].strip()
        try:
            numeric = float(text)
        except ValueError:
            return None
    elif isinstance(value, int | float):
        numeric = float(value)
    else:
        return None
    if not isfinite(numeric) or numeric == 0:
        return None
    if percentage or numeric > 4:
        numeric /= 100
    return numeric if 0.5 <= numeric <= 4 else None


def load_window_scale_factor(
    config_port: GeometryConfigReader | None,
    *,
    detected_factor: float,
) -> float:
    """Resolve the declared window scale with strict enable-flag handling."""

    enabled = _config_get(config_port, "gui.dpi.enable_dpi_scaling", True)
    if enabled is False:
        return 1.0
    forced = normalize_ui_scale(_config_get(config_port, "gui.dpi.ui_scale", 0))
    if forced is not None:
        return forced
    detected = normalize_ui_scale(detected_factor)
    return 1.0 if detected is None else detected


__all__ = [
    "DEFAULT_CENTER_PANEL_SCREEN_X",
    "DEFAULT_MIN_HEIGHT",
    "DEFAULT_MIN_WIDTH",
    "DEFAULT_WINDOW_HEIGHT",
    "DEFAULT_WINDOW_WIDTH",
    "GEOMETRY_SCHEMA_VERSION",
    "GeometryConfigReader",
    "RecoveredWindowGeometry",
    "ScreenRect",
    "WindowGeometryPolicy",
    "WindowRect",
    "build_canonical_geometry_values",
    "center_window_geometry",
    "load_window_geometry_policy",
    "load_window_scale_factor",
    "normalize_ui_scale",
    "recover_window_geometry",
]
