"""Shared fixtures for gongwen optimizer tests."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_runtime.config.loader import ConfigLoader

PROJECT_ROOT = Path(__file__).resolve().parents[5]
PROJECT_CONFIGS = PROJECT_ROOT / "configs"


@pytest.fixture(autouse=True)
def _inject_numbering_clean_rules(tmp_path: Path) -> None:
    """Mirror app startup so heading-number cleanup uses the real rule source."""
    ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
