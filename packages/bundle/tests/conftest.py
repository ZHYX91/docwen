"""Test path bootstrap for bundle package tests."""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[3]

LOCAL_SRC_PATHS = [
    PROJECT_ROOT,
    PROJECT_ROOT / "packages" / "core" / "src",
    PROJECT_ROOT / "packages" / "application" / "src",
    PROJECT_ROOT / "packages" / "runtime" / "src",
    PROJECT_ROOT / "packages" / "bundle" / "src",
    PROJECT_ROOT / "packages" / "apps" / "cli" / "src",
    PROJECT_ROOT / "packages" / "apps" / "gui" / "src",
]

for path in reversed(LOCAL_SRC_PATHS):
    text = str(path)
    if text not in sys.path:
        sys.path.insert(0, text)
