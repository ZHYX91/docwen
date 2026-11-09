"""Pre-import opt-out for ONNX Runtime's native Linux telemetry provider.

The provider latches its setting at initialization. We cannot reconstruct an
embedding host's prior environment; a preloaded runtime is accepted only when
its current explicit opt-out is intact, never as proof of historical isolation.
"""

from __future__ import annotations

import os
import sys


def ensure_onnxruntime_telemetry_disabled() -> None:
    """Force the opt-out, rejecting an observably unsafe late bootstrap."""
    if sys.platform != "linux":
        return
    if os.environ.get("ORT_DISABLE_TELEMETRY") != "1" and any(
        name == "onnxruntime" or name.startswith("onnxruntime.") for name in sys.modules
    ):
        raise RuntimeError("onnxruntime_telemetry_bootstrap_started_too_late")
    os.environ["ORT_DISABLE_TELEMETRY"] = "1"


ensure_onnxruntime_telemetry_disabled()
