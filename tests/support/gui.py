from __future__ import annotations

import contextlib
import time
from typing import Any


def shutdown_main_window(window: Any) -> None:
    from PySide6.QtWidgets import QApplication

    with contextlib.suppress(Exception):
        window.close()

    app = QApplication.instance()
    if app is not None:
        with contextlib.suppress(Exception):
            app.processEvents()

    deadline = time.monotonic() + 6.0
    while app is not None and getattr(window, "_execution_close_pending", False) and time.monotonic() < deadline:
        with contextlib.suppress(Exception):
            app.processEvents()
        time.sleep(0.005)

    if app is not None:
        with contextlib.suppress(Exception):
            app.processEvents()
