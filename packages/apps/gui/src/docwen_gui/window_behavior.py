"""GUI-owned projection of the persisted main-window behavior policy."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol


class _ConfigGetter(Protocol):
    def get(self, key: str, default: Any = None) -> Any: ...


@dataclass(frozen=True, slots=True)
class WindowBehaviorPolicy:
    """Typed window behavior consumed by the GUI composition root."""

    remember_gui_state: bool = True
    auto_center: bool = False
    expand_side_panels: bool = False
    center_panel_width: int = 460
    left_panel_width: int = 400
    right_panel_width: int = 300


DEFAULT_WINDOW_BEHAVIOR = WindowBehaviorPolicy()


@dataclass(frozen=True, slots=True)
class PanelWidthPlan:
    """Runtime minimum widths derived from logical panel-width intent.

    The configured widths are desired proportions, not pixels captured from a
    particular maximized, snapped, or DPI-scaled window.  A narrow viewport
    therefore compresses all visible columns proportionally instead of
    preserving stale side-panel pixels and starving the center workflow.
    """

    left: int
    center: int
    right: int


def plan_panel_widths(
    policy: WindowBehaviorPolicy,
    *,
    container_width: int,
    horizontal_margins: int,
    spacing: int,
    left_visible: bool,
    right_visible: bool,
    scale_factor: float = 1.0,
) -> PanelWidthPlan:
    """Project semantic panel widths into one concrete viewport.

    Hidden columns always receive a zero minimum.  When the desired widths fit,
    their configured values remain intact.  Otherwise every visible column is
    reduced by the same ratio and integer rounding is distributed
    deterministically.  This makes the result independent of the order in
    which panels became visible.
    """

    scale = max(0.01, float(scale_factor))
    desired = {
        "left": max(1, round(policy.left_panel_width * scale)) if left_visible else 0,
        "center": max(1, round(policy.center_panel_width * scale)),
        "right": max(1, round(policy.right_panel_width * scale)) if right_visible else 0,
    }
    visible_names = tuple(name for name in ("left", "center", "right") if desired[name] > 0)
    gap_count = max(0, len(visible_names) - 1)
    usable = max(
        0,
        int(container_width) - max(0, int(horizontal_margins)) - gap_count * max(0, int(spacing)),
    )
    desired_total = sum(desired.values())
    if desired_total <= usable:
        return PanelWidthPlan(**desired)
    if usable <= 0:
        return PanelWidthPlan(left=0, center=0, right=0)

    # The real GUI minimum is far wider than the three visible columns, but
    # keeping this branch total makes the policy safe for synthetic/narrow
    # layouts too.
    minimum_each = 1 if usable >= len(visible_names) else 0
    allocated = {name: (desired[name] * usable) // desired_total for name in visible_names}
    if minimum_each:
        allocated = {name: max(minimum_each, value) for name, value in allocated.items()}

    overflow = sum(allocated.values()) - usable
    if overflow > 0:
        # Remove rounding overflow from the largest columns first while
        # retaining a non-zero visible minimum whenever possible.
        for name in sorted(visible_names, key=lambda item: (-allocated[item], item)):
            removable = max(0, allocated[name] - minimum_each)
            reduction = min(removable, overflow)
            allocated[name] -= reduction
            overflow -= reduction
            if overflow == 0:
                break

    remainder = usable - sum(allocated.values())
    if remainder > 0:
        ranked = sorted(
            visible_names,
            key=lambda name: (
                -((desired[name] * usable) % desired_total),
                0 if name == "center" else 1,
                name,
            ),
        )
        for index in range(remainder):
            allocated[ranked[index % len(ranked)]] += 1

    return PanelWidthPlan(
        left=allocated.get("left", 0),
        center=allocated.get("center", 0),
        right=allocated.get("right", 0),
    )


def _read_bool(config_port: _ConfigGetter | None, key: str, default: bool) -> bool:
    if config_port is None:
        return default
    try:
        value = config_port.get(key, default)
    except Exception:
        return default
    return value if isinstance(value, bool) else default


def _read_bounded_int(
    config_port: _ConfigGetter | None,
    key: str,
    default: int,
    *,
    minimum: int,
    maximum: int,
) -> int:
    if config_port is None:
        return default
    try:
        value = config_port.get(key, default)
    except Exception:
        return default
    if isinstance(value, bool) or not isinstance(value, int):
        return default
    return value if minimum <= value <= maximum else default


def load_window_behavior_policy(config_port: _ConfigGetter | None) -> WindowBehaviorPolicy:
    """Read bounded window flags and logical panel widths from configuration."""

    defaults = DEFAULT_WINDOW_BEHAVIOR
    return WindowBehaviorPolicy(
        remember_gui_state=_read_bool(
            config_port,
            "gui.window.remember_gui_state",
            defaults.remember_gui_state,
        ),
        auto_center=_read_bool(
            config_port,
            "gui.window.auto_center",
            defaults.auto_center,
        ),
        expand_side_panels=_read_bool(
            config_port,
            "gui.window.expand_side_panels",
            defaults.expand_side_panels,
        ),
        center_panel_width=_read_bounded_int(
            config_port,
            "gui.window.center_panel_width",
            defaults.center_panel_width,
            minimum=360,
            maximum=960,
        ),
        left_panel_width=_read_bounded_int(
            config_port,
            "gui.window.left_panel_width",
            defaults.left_panel_width,
            minimum=300,
            maximum=720,
        ),
        right_panel_width=_read_bounded_int(
            config_port,
            "gui.window.right_panel_width",
            defaults.right_panel_width,
            minimum=280,
            maximum=720,
        ),
    )
