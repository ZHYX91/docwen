"""Contract tests for ConfigLoader: loading, three-layer merge, write/reset, helpers."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit


def _write_toml(filepath: Path, data: dict) -> None:
    """Write test data as TOML."""
    import tomlkit

    filepath.write_text(tomlkit.dumps(data), encoding="utf-8", newline="\n")


def write_minimal_base_config_tree(base_dir: Path) -> None:
    """Create an empty TOML file for every spec in the registry under *base_dir*."""
    from docwen_runtime.config.registry import CONFIG_FILES

    for spec in CONFIG_FILES:
        path = base_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


DIR = Path(__file__).resolve().parent

PROJECT_CONFIGS = DIR.parent.parent.parent / "configs"

__all__ = (
    "PROJECT_CONFIGS",
    "Any",
    "Path",
    "pytest",
    "pytestmark",
    "write_minimal_base_config_tree",
)
