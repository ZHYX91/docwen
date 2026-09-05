"""Cancellable, identity-cached protection analysis outside the GUI thread."""

from __future__ import annotations

from collections import OrderedDict
from pathlib import Path
from threading import Event, Lock

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal, Slot

from docwen_gui.spreadsheet_protection import SpreadsheetProtectionInfo, inspect_xlsx_protection


def _identity(path: str) -> tuple[int, ...]:
    stat = Path(path).stat()
    return stat.st_dev, stat.st_ino, stat.st_size, stat.st_mtime_ns, stat.st_ctime_ns


class _Inspector:
    def __init__(self) -> None:
        self._cache: OrderedDict[str, tuple[tuple[int, ...], SpreadsheetProtectionInfo]] = OrderedDict()
        self._lock = Lock()

    def inspect(self, path: str) -> SpreadsheetProtectionInfo:
        try:
            identity = _identity(path)
            with self._lock:
                cached = self._cache.get(path)
                if cached is not None and cached[0] == identity:
                    self._cache.move_to_end(path)
                    return cached[1]
            result = inspect_xlsx_protection(path)
            if _identity(path) != identity:
                return SpreadsheetProtectionInfo(path, "unknown")
            with self._lock:
                self._cache[path] = (identity, result)
                self._cache.move_to_end(path)
                while len(self._cache) > 256:
                    self._cache.popitem(last=False)
            return result
        except Exception:
            # Background inspection is advisory; unreadable or changed inputs
            # fail closed here and still face authoritative admission at run time.
            return SpreadsheetProtectionInfo(path, "unknown")


class _ResultSignal(QObject):
    ready = Signal(int, object)


class _InspectionJob(QRunnable):
    def __init__(
        self, generation: int, paths: tuple[str, ...], inspector: _Inspector, cancelled: Event, closed: Event
    ) -> None:
        super().__init__()
        self.generation = generation
        self.paths = paths
        self.inspector = inspector
        self.cancelled = cancelled
        self.closed = closed
        self.signals = _ResultSignal()

    def run(self) -> None:
        results = []
        for path in self.paths:
            if self.cancelled.is_set() or self.closed.is_set():
                return
            results.append(self.inspector.inspect(path))
        if not self.cancelled.is_set() and not self.closed.is_set():
            self.signals.ready.emit(self.generation, tuple(results))


class SpreadsheetAnalysis(QObject):
    """Only the latest selection can publish results; identical renders are free."""

    changed = Signal()

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._inspector = _Inspector()
        self._generation = 0
        self._paths: tuple[str, ...] = ()
        self._cancelled = Event()
        self._closed = False
        self._lifetime = Event()
        self.destroyed.connect(lambda _object=None, token=self._lifetime: token.set())
        self.pending = False
        self.results: tuple[SpreadsheetProtectionInfo, ...] = ()

    def select(self, paths: tuple[str, ...], *, refresh: bool = False) -> None:
        paths = tuple(dict.fromkeys(paths))
        if self._closed or (paths == self._paths and not refresh):
            return
        self._cancelled.set()
        self._cancelled = Event()
        self._generation += 1
        self._paths = paths
        self.pending = bool(paths)
        self.results = ()
        if paths:
            job = _InspectionJob(self._generation, paths, self._inspector, self._cancelled, self._lifetime)
            job.signals.ready.connect(self._accept)
            # The job owns its signal object. Destruction of this receiver
            # disconnects it without touching widgets from a worker thread.
            QThreadPool.globalInstance().start(job)
        self.changed.emit()

    @Slot(int, object)
    def _accept(self, generation: int, results: tuple[SpreadsheetProtectionInfo, ...]) -> None:
        if self._closed or generation != self._generation:
            return
        self.results = results
        self.pending = False
        self.changed.emit()

    def close(self) -> None:
        self._closed = True
        self._lifetime.set()
        self._cancelled.set()
