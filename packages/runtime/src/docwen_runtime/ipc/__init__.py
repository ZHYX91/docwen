"""Single-instance ownership primitives.

GUI command delivery lives in :mod:`docwen_runtime.control`; this package only
owns the process lock and its test-mode toggle.
"""

from __future__ import annotations

from .single_instance import SingleInstance

__all__ = [
    "SingleInstance",
    "disable_ipc",
    "enable_ipc",
    "is_ipc_disabled",
]

# ── Test-mode IPC toggle ────────────────────────────────────────────────
# When True, single-instance locking becomes a no-op for isolated tests.
_ipc_disabled: bool = False


def disable_ipc() -> None:
    """Disable IPC globally (for tests)."""
    global _ipc_disabled
    _ipc_disabled = True


def enable_ipc() -> None:
    """Re-enable IPC globally (undo disable_ipc)."""
    global _ipc_disabled
    _ipc_disabled = False


def is_ipc_disabled() -> bool:
    """Return True if IPC has been globally disabled."""
    return _ipc_disabled
