"""Entry point for ``python -m docwen_gui``.

Delegates to ``docwen_bundle.gui_entry.main`` so that IPC bootstrap,
single-instance locking, file forwarding, controller wiring, and the
task-event bridge are all owned by the bundle composition root. The
gui package does not perform IPC bootstrap itself.
"""

from __future__ import annotations

import sys

import docwen_bundle.gui_entry as gui_entry


def main() -> int:
    """Run the GUI via the bundle entry point with full IPC bootstrap.

    Resolves ``gui_entry.main`` at call time so the gui package never
    holds a bound reference that bypasses the bundle composition root.
    """
    return gui_entry.main()


if __name__ == "__main__":
    sys.exit(main())
