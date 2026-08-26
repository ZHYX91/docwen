"""Tests for GUI feedback helpers, AboutDialog, and icon resources.

Covers:
- feedback module: error, warn, info, confirm, notify, exception
- AboutDialog: creation, modal, required labels
- resources: IconManager, icon path resolution, SVG loading
- main_window: about button wiring, window opacity, icon initialization
"""

from __future__ import annotations

import pytest
from PySide6.QtCore import Qt

pytestmark = pytest.mark.gui


def _write_minimal_base_config_tree(base_dir) -> None:
    """Seeds a minimal base config tree so ConfigLoader does not crash.

    ConfigLoader expects every CONFIG_FILES entry to exist before the
    three-layer merge can proceed.  An empty tmp_path won't have them.
    """
    from pathlib import Path as _P

    from docwen_runtime.config.registry import CONFIG_FILES

    bd = _P(base_dir)
    for spec in CONFIG_FILES:
        path = bd / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


__all__ = (
    "Qt",
    "_write_minimal_base_config_tree",
    "pytest",
    "pytestmark",
)
