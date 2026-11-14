"""Opt-in GUI startup timings for source and packaged acceptance.

The external launcher owns the process-start clock. These stages describe
application work, including the first paint and a subsequent event-loop turn;
they do not claim that the operating system has displayed a frame on a monitor.
"""

from __future__ import annotations

import logging
import os
import time
from pathlib import Path
from typing import Any


class StartupTrace:
    def __init__(self, report: str | None) -> None:
        self._report = Path(report) if report else None
        self._stages: dict[str, int] = {}
        self._written = False
        self.mark("python_entry")

    def mark(self, name: str) -> None:
        if self._report is not None:
            self._stages.setdefault(name, time.perf_counter_ns())

    def observe_window(self, window: Any) -> None:
        """Observe paint without changing normal window/event-loop scheduling."""
        if self._report is None:
            return
        from PySide6.QtCore import QEvent, QObject, QTimer

        trace = self

        class FirstPaintObserver(QObject):
            def eventFilter(self, watched: QObject, event: QEvent) -> bool:
                if watched is window and event.type() == QEvent.Type.Paint:
                    trace.mark("first_paint")
                    window.removeEventFilter(self)
                    QTimer.singleShot(0, self.after_paint)
                return False

            def after_paint(self) -> None:
                trace.mark("event_loop_ready")
                trace.write(window_visible=window.isVisible(), window_enabled=window.isEnabled())
                self.deleteLater()

        observer = FirstPaintObserver(window)
        window.installEventFilter(observer)
        # QObject parenting owns the native lifetime; retain the Python
        # subclass until its event callback has run as well.
        window._docwen_startup_observer = observer

    def write(self, *, window_visible: bool, window_enabled: bool) -> None:
        if self._report is None or self._written:
            return
        self._written = True
        import json

        start = self._stages["python_entry"]
        payload = {
            "schema": "docwen.gui_startup.v1",
            "pid": os.getpid(),
            "counter_ns": self._stages,
            "elapsed_ms": {name: (value - start) / 1_000_000 for name, value in self._stages.items()},
            "window_visible": window_visible,
            "window_enabled": window_enabled,
            "scope": "application stages; external launch and first conversion measured separately",
        }
        try:
            # Acceptance supplies an existing owned parent and a fresh path.
            # Never replace a previous run's evidence or create directories.
            with self._report.open("x", encoding="utf-8") as stream:
                json.dump(payload, stream, indent=2)
        except OSError:
            logging.getLogger(__name__).warning("Unable to write GUI startup diagnostics", exc_info=True)


gui_startup = StartupTrace(os.environ.get("DOCWEN_GUI_TEST_STARTUP_REPORT"))
