"""Shared per-user identity for local ownership and GUI control endpoints."""

from __future__ import annotations

import getpass
import hashlib
from pathlib import Path


def user_namespace() -> str:
    identity = f"{getpass.getuser()}|{Path.home()}".encode("utf-8", errors="replace")
    return hashlib.sha256(identity).hexdigest()[:16]
