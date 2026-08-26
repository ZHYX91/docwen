"""GUI resource helper tests."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


def test_gui_resource_helpers_delegate_to_runtime_registry(tmp_path, monkeypatch) -> None:
    from docwen_gui import resources as gui_resources
    from docwen_runtime.resources.registry import ResourceRegistry

    root = tmp_path / "bundle"
    (root / "assets" / "icons").mkdir(parents=True)
    (root / "i18n" / "locales").mkdir(parents=True)
    monkeypatch.setattr(ResourceRegistry, "default", classmethod(lambda cls: ResourceRegistry(root)))

    assert gui_resources.app_root() == root
    assert gui_resources.asset_path("icons/app.png") == root / "assets" / "icons" / "app.png"
    assert gui_resources.i18n_locales_dir() == root / "i18n" / "locales"
