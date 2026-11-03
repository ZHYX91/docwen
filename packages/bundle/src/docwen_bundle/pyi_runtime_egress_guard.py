"""PyInstaller runtime hook that activates DocWen's egress guard first.

PyInstaller executes user-supplied runtime hooks before its built-in runtime
hooks.  Keep this module stdlib-only until the import-order assertion passes;
then import only the Runtime-owned guard implementation.
"""

from __future__ import annotations

import os
import sys

# Native dependencies must never precede this hook, even in frozen helpers.
_early_native_imports = sorted(name for name in sys.modules if name.startswith(("onnxruntime", "rapidocr_onnxruntime")))
if _early_native_imports:
    raise RuntimeError(f"dependency_egress_guard_started_too_late:{_early_native_imports}")

# Run before every other hook, including in frozen multiprocessing helpers.
# This native opt-out is independent of the Python dependency egress guard.
if sys.platform == "linux":
    os.environ["ORT_DISABLE_TELEMETRY"] = "1"

_FORBIDDEN_EARLY_IMPORT_PREFIXES = (
    "onnxruntime",
    "rapidocr_onnxruntime",
    "PySide6",
    "qfluentwidgets",
    "docwen_application",
    "docwen_cli",
    "docwen_core",
    "docwen_gui",
    "docwen_plugin_",
)


def _is_pyinstaller_multiprocessing_child() -> bool:
    """Recognize frozen multiprocessing workers before PyInstaller patches them."""

    if len(sys.argv) >= 2 and sys.argv[1] == "--multiprocessing-fork":
        return True
    return (
        len(sys.argv) >= 2
        and sys.argv[-2] == "-c"
        and sys.argv[-1].startswith(
            (
                "from multiprocessing.resource_tracker import main",
                "from multiprocessing.forkserver import main",
            )
        )
    )


def _assert_pre_product_import_boundary() -> None:
    early_imports = sorted(name for name in sys.modules if name.startswith(_FORBIDDEN_EARLY_IMPORT_PREFIXES))
    if early_imports:
        raise RuntimeError(f"dependency_egress_guard_started_too_late:{early_imports}")


if not _is_pyinstaller_multiprocessing_child():
    _assert_pre_product_import_boundary()

    from docwen_runtime.security.network import (
        activate_process_lifetime_dependency_egress_guard,
    )

    _status = activate_process_lifetime_dependency_egress_guard()
    if not _status.active or _status.bootstrap != "pyinstaller_runtime_hook":
        raise RuntimeError("dependency_egress_guard_runtime_hook_not_enforced")

    del _status
