"""Application content scale in Qt logical pixels, independent of screen DPI.

Design values are always stored at 100%. Declarative geometry bindings are
reapplied from those values, never from the last rounded widget dimensions.
Font-dependent geometry stays with the widget's normal layout/sizeHint code.
"""

from __future__ import annotations

import re
import weakref

from PySide6.QtCore import QObject, QSize
from shiboken6 import isValid

SCALE_PRESETS = (90, 100, 110, 125, 150)
_scale_percent = 100
_bindings: weakref.WeakSet[_MetricBinding] = weakref.WeakSet()


def normalize_scale_percent(value: object) -> int:
    return value if isinstance(value, int) and not isinstance(value, bool) and value in SCALE_PRESETS else 100


def scale_percent() -> int:
    return _scale_percent


def dp(value: float) -> int:
    """Convert one design distance to Qt logical pixels (not physical pixels)."""
    return round(value * _scale_percent / 100)


def scaled_stylesheet(stylesheet: str) -> str:
    """Scale the assembled base stylesheet once; preserve fractional font sizes."""
    return re.sub(
        r"(-?\d+(?:\.\d+)?)(px|pt)\b",
        lambda match: f"{float(match[1]) * _scale_percent / 100:g}{match[2]}",
        stylesheet,
    )


class _MetricBinding(QObject):
    def __init__(self, target: QObject) -> None:
        super().__init__(target)
        self.target = weakref.ref(target)
        self.metrics: dict[str, tuple[int | QSize, ...]] = {}
        _bindings.add(self)

    def apply(self, method: str) -> None:
        target = self.target()
        if target is None or not isValid(target):
            return
        values = self.metrics[method]
        scaled = tuple(
            QSize(dp(value.width()), dp(value.height()))
            if isinstance(value, QSize)
            else value
            if value in (-1, 16777215)
            else dp(value)
            for value in values
        )
        getattr(target, method)(*scaled)


def set_metric(target: QObject, method: str, *values: int | QSize) -> None:
    """Set and bind a declarative size, margin, spacing or icon size."""
    binding = getattr(target, "_docwen_ui_metrics", None)
    if not isinstance(binding, _MetricBinding):
        binding = _MetricBinding(target)
        target.__dict__["_docwen_ui_metrics"] = binding
    binding.metrics[method] = values
    binding.apply(method)


def apply_scale_percent(value: object) -> int:
    global _scale_percent
    _scale_percent = normalize_scale_percent(value)
    for binding in tuple(_bindings):
        if isValid(binding):
            for method in tuple(binding.metrics):
                binding.apply(method)
    return _scale_percent
