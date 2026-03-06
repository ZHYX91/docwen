"""services 单元测试。"""

from __future__ import annotations

import pytest

from docwen.services.strategies import registry

pytestmark = pytest.mark.unit


def test_registry_specs_are_resolvable() -> None:
    failures = registry.validate_registry_specs()
    assert failures == {}


def test_registry_runtime_is_consistent() -> None:
    failures = registry.validate_registry_runtime()
    assert failures == {}


def test_registry_specs_detect_duplicate_normalized_action_names(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        registry,
        "_ACTION_TO_MODULE",
        {
            "merge_pdfs": "docwen.services.strategies.layout.operations",
            " merge_pdfs ": "docwen.services.strategies.layout.operations",
        },
        raising=True,
    )
    failures = registry.validate_registry_specs()
    assert failures.get("action:merge_pdfs") is not None


def test_registry_specs_detect_empty_action_module_path(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        registry,
        "_ACTION_TO_MODULE",
        {"action_a": "   "},
        raising=True,
    )
    failures = registry.validate_registry_specs()
    assert failures.get("action:action_a") == "empty module path"
