"""Repository gate: tests must never bind the default loader to a user profile."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_owned_config_loader_uses_disposable_session_root() -> None:
    from docwen_runtime.config import ConfigLoader

    isolated_root = Path(os.environ["DOCWEN_CONFIG_DIR"]).resolve()
    loader = ConfigLoader()

    assert loader.user_dir.resolve() == isolated_root
