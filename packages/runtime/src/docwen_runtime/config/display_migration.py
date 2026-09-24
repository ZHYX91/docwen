"""One-time removal of the retired partial-window DPI settings.

Old window-only scale is deliberately not reinterpreted as content scale.
Only the retired display keys are changed; unrelated preferences survive.
The migration is idempotent and leaves no runtime compatibility aliases.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any


def migrate_display_preferences(data: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(data)
    dpi = result.get("dpi")
    if isinstance(dpi, dict):
        dpi.pop("enable_dpi_scaling", None)
        dpi.pop("ui_scale", None)
        if not dpi:
            result.pop("dpi")
    font = result.get("font")
    if isinstance(font, dict) and font.get("size_preset") in ("xlarge", "extra_large", "extra-large"):
        font["size_preset"] = "large"
    return result
