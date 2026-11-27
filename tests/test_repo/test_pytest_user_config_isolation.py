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


def test_profile_and_tmp_path_share_one_lazy_directory(request: pytest.FixtureRequest) -> None:
    profile = Path(os.environ["DOCWEN_DATA_DIR"])
    assert not profile.parent.exists()

    # A profile consumer can write before a later fixture requests tmp_path.
    profile.mkdir(parents=True)
    marker = profile / "catalog.txt"
    marker.write_text("catalog", encoding="utf-8")
    assert request.getfixturevalue("tmp_path") == profile.parent
    assert marker.read_text(encoding="utf-8") == "catalog"
