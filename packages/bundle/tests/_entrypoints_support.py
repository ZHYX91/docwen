from __future__ import annotations

from collections.abc import Callable
from types import SimpleNamespace
from typing import Any, cast

import pytest
from tests.support.subprocess_runner import run_subprocess

from docwen_core.models import TaskEvent

pytestmark = pytest.mark.unit


class _Signal:
    def __init__(self) -> None:
        self._callbacks: list[Callable[[], object]] = []

    def connect(self, callback: Callable[[], object]) -> None:
        self._callbacks.append(callback)

    def emit(self) -> None:
        for callback in list(self._callbacks):
            callback()


__all__ = (
    "Any",
    "Callable",
    "SimpleNamespace",
    "TaskEvent",
    "_Signal",
    "cast",
    "pytest",
    "pytestmark",
    "run_subprocess",
)
