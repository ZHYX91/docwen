"""Fake implementation of CancellationTokenView protocol."""

from __future__ import annotations


class FakeCancellationTokenView:
    """Fake cancellation token view — always allows progress.

    For cancellation tests (request → cancel → assert raises), use the
    real ``CancellationToken`` directly — this fake never raises.
    """

    def __init__(self, cancelled: bool = False) -> None:
        self._cancelled = cancelled

    @property
    def is_cancelled(self) -> bool:
        return self._cancelled

    def check(self) -> None:
        pass  # never raises — cancellation tests use real CancellationToken
