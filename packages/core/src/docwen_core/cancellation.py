"""CancellationToken — cooperative cancellation for long-running tasks.

.. warning::

    ``CancellationToken`` uses ``threading.Lock`` and is therefore
    **thread-safe but NOT process-safe**.  It must never be pickled,
    serialised into a ``WorkerRequest`` payload, or shared across
    a process boundary.  For multi-process workers the runtime MUST
    deliver cancellation signals through a separate cross-process
    channel (e.g. ``multiprocessing.Event``, file sentinel, or pipe).

    ``_CancellationTokenViewImpl`` is the concrete read-only wrapper
    returned by ``CancellationToken.view()``.  The public API type that
    plugins and protocols should annotate against is the Protocol
    ``docwen_core.protocols.execution_context.CancellationTokenView``.
"""

from __future__ import annotations

import contextlib
import threading
from collections.abc import Callable


class CancellationToken:
    """A thread-safe cancellation token.

    The *source* (main process / task manager) calls ``cancel()``.
    Workers call ``check()`` or read ``is_cancelled`` periodically.

    **Scope**: single-process only (uses ``threading.Lock``).
    See module-level docstring for cross-process constraints.

    This class lives in core so plugins can depend on it without
    importing runtime internals.
    """

    def __init__(self) -> None:
        self._cancelled: bool = False
        self._cancel_reason: str = ""
        self._lock = threading.Lock()
        self._callbacks: list[Callable[[], None]] = []

    @property
    def is_cancelled(self) -> bool:
        """Return ``True`` if cancellation has been requested."""
        with self._lock:
            return self._cancelled

    def cancel(self, reason: str = "") -> None:
        """Request cancellation.

        Idempotent — calling multiple times has the same effect as calling once.

        Args:
            reason: Optional human-readable reason (``"timeout"``,
                    ``"user_cancelled"``, etc.).  Stored for inspection
                    but not forwarded to callbacks (callbacks are
                    ``Callable[[], None]``).
        """
        callbacks: list[Callable[[], None]] = []
        with self._lock:
            if not self._cancelled:
                self._cancelled = True
                self._cancel_reason = reason
                callbacks = list(self._callbacks)
        import contextlib

        for cb in callbacks:
            with contextlib.suppress(Exception):
                cb()  # callbacks must not break cancellation

    @property
    def reason(self) -> str:
        """The reason string passed to ``cancel()``, or ``""`` if not cancelled."""
        with self._lock:
            return self._cancel_reason

    def check(self) -> None:
        """Raise ``CancellationRequested`` if cancellation has been requested.

        Plugins should call this:
        - At the start of every loop iteration
        - Before calling an external tool or subprocess
        - Before writing a large file to staging
        """
        if self.is_cancelled:
            from docwen_core.errors import CancellationRequested

            raise CancellationRequested("Task has been cancelled")

    def add_callback(self, cb: Callable[[], None]) -> None:
        """Register a callback to be invoked on cancellation.

        If already cancelled the callback is invoked immediately.
        Callbacks should be fast and must not raise.
        """
        with self._lock:
            if self._cancelled:
                # Already cancelled — invoke immediately
                with contextlib.suppress(Exception):
                    cb()
                return
            self._callbacks.append(cb)

    def view(self) -> _CancellationTokenViewImpl:
        """Return a read-only view suitable for passing to plugins.

        The returned object satisfies the
        ``docwen_core.protocols.execution_context.CancellationTokenView`` Protocol.
        """
        return _CancellationTokenViewImpl(self)


class _CancellationTokenViewImpl:
    """Concrete read-only view of a ``CancellationToken`` (internal).

    This is the object returned by ``CancellationToken.view()``.
    Plugins should type-annotate against the Protocol
    ``docwen_core.protocols.execution_context.CancellationTokenView``,
    not against this concrete class.

    Like ``CancellationToken``, this object is **not process-safe**
    and must never be serialised across a process boundary.
    """

    def __init__(self, token: CancellationToken) -> None:
        self._token = token

    @property
    def is_cancelled(self) -> bool:
        return self._token.is_cancelled

    def check(self) -> None:
        self._token.check()
