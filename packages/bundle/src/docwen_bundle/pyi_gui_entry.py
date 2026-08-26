"""PyInstaller entry point for DocWen GUI.

This thin wrapper exists solely as a filesystem entry point for PyInstaller.
It delegates to ``docwen_bundle.gui_entry.main`` — the same entry used by the
``docwen-gui`` console_script and ``python -m docwen_gui``.

Usage (by PyInstaller only)::

    pyinstaller pyi_gui_entry.py --name=DocWen ...

Do NOT use this file as a user-facing entry point. Use ``docwen-gui``
(console_script) or ``python -m docwen_gui`` instead.
"""

from __future__ import annotations

import multiprocessing
import sys

import docwen_bundle.gui_entry as gui_entry


def _prepare_multiprocessing() -> None:
    """Let frozen child processes enter their multiprocessing bootstrap."""
    multiprocessing.freeze_support()


def _delegate() -> int:
    """Run the bundle GUI entry and return its exit code.

    Exposed as a function so tests can verify delegation without driving
    ``__main__`` execution. Resolves ``gui_entry.main`` at call time so
    monkeypatching ``gui_entry.main`` in tests takes effect.
    """
    return gui_entry.main()


if __name__ == "__main__":
    _prepare_multiprocessing()
    sys.exit(_delegate())
