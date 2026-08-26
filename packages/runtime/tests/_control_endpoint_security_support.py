"""Security and lifecycle contracts for Unix GUI control endpoints."""

from __future__ import annotations

import errno
import os
import socket
import stat
import sys
import threading
import time
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any
from uuid import uuid4

import pytest

pytestmark = pytest.mark.contract


def _private_directory(path: Path) -> Path:
    path.mkdir(mode=0o700, parents=True)
    path.chmod(0o700)
    return path


__all__ = (
    "Any",
    "Path",
    "TemporaryDirectory",
    "_private_directory",
    "errno",
    "os",
    "pytest",
    "pytestmark",
    "socket",
    "stat",
    "sys",
    "threading",
    "time",
    "uuid4",
)
