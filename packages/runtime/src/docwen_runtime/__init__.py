"""DocWen Runtime package.

Package import stays deliberately lightweight so the bundle composition root
can install the dependency egress guard before loading runtime services.
Submodules remain importable through normal ``from docwen_runtime import x``
semantics; the version is resolved lazily on first access.
"""

from __future__ import annotations

import os
import sys

# Deliberately duplicated in Core and Runtime: Core cannot depend on Runtime,
# and early Runtime/frozen startup must not import Core's product modules.
# ORT 1.31.1's POSIX provider latches this opt-out during initialization;
# Python socket audit hooks do not protect against native telemetry.
if sys.platform == "linux":
    if os.environ.get("ORT_DISABLE_TELEMETRY") != "1" and any(
        name == "onnxruntime" or name.startswith("onnxruntime.") for name in sys.modules
    ):
        raise RuntimeError("onnxruntime_telemetry_bootstrap_started_too_late")
    os.environ["ORT_DISABLE_TELEMETRY"] = "1"

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_core.version import __version__ as __version__

__all__ = ["__version__"]


def __getattr__(name: str) -> Any:
    if name == "__version__":
        from docwen_core.version import __version__

        return __version__
    raise AttributeError(name)
