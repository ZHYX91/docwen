from __future__ import annotations

import contextlib
import time
from typing import Any


def wait_for_template_catalog(window: Any) -> None:
    from PySide6.QtTest import QTest

    deadline = time.monotonic() + 10.0
    while window._template_vm.busy and time.monotonic() < deadline:
        QTest.qWait(5)
    assert not window._template_vm.busy, "Template catalog did not finish loading"


def shutdown_main_window(window: Any) -> None:
    from PySide6.QtWidgets import QApplication

    with contextlib.suppress(Exception):
        window.close()

    app = QApplication.instance()
    if app is not None:
        with contextlib.suppress(Exception):
            app.processEvents()

    deadline = time.monotonic() + 6.0
    while (
        app is not None
        and time.monotonic() < deadline
        and (
            getattr(window, "_execution_close_pending", False)
            or getattr(getattr(window, "_template_vm", None), "busy", False)
        )
    ):
        with contextlib.suppress(Exception):
            app.processEvents()
        time.sleep(0.005)

    if app is not None:
        with contextlib.suppress(Exception):
            app.processEvents()
