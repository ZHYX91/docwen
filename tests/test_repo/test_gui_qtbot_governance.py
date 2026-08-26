from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.contract


def test_gui_tests_live_in_workspace_package() -> None:
    repo_root = Path(__file__).resolve().parents[2]
    gui_tests_dir = repo_root / "packages" / "apps" / "gui" / "tests"
    assert gui_tests_dir.is_dir()
    assert sorted(gui_tests_dir.glob("test_*.py"))
