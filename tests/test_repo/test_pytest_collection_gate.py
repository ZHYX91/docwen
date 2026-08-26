from __future__ import annotations

from pathlib import Path

import pytest
from tests._pytest_hooks import collection
from tests._pytest_hooks.dependencies import _NOT_COLLECTED_BY_REASON

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _preserve_not_collected_state() -> None:
    original = {reason: set(paths) for reason, paths in _NOT_COLLECTED_BY_REASON.items()}
    try:
        yield
    finally:
        _NOT_COLLECTED_BY_REASON.clear()
        for reason, paths in original.items():
            _NOT_COLLECTED_BY_REASON[reason].update(paths)


@pytest.mark.parametrize(
    "legacy_path",
    [
        Path("tests/test_gui"),
        Path("tests/test_repo/test_coexistence.py"),
        Path("packages/apps/cli/tests/test_main_validation.py"),
    ],
)
def test_legacy_monolith_paths_fail_collection_instead_of_disappearing(legacy_path: Path) -> None:
    with pytest.raises(pytest.UsageError, match="legacy monolith test path returned"):
        collection.pytest_ignore_collect(legacy_path, config=None)


def test_dependency_gates_remain_visible_not_collected_reasons(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(collection, "_PIL_OK", False)
    path = Path("packages/plugins/image/tests/test_image_core.py")

    assert collection.pytest_ignore_collect(path, config=None) is True
    assert collection._relative_collection_path(path) in _NOT_COLLECTED_BY_REASON["missing Pillow dependency"]


def test_ordinary_test_path_is_not_ignored() -> None:
    assert collection.pytest_ignore_collect(Path("packages/core/tests/test_models.py"), config=None) is False
