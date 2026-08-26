"""Resource root discovery tests."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_pyinstaller_onedir_root_uses_deploy_root_and_internal_locales(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_runtime.resources.registry import ResourceRegistry, find_project_root

    deploy_root = tmp_path / "DocWen"
    internal_root = deploy_root / "_internal"
    (deploy_root / "templates").mkdir(parents=True)
    (deploy_root / "configs").mkdir()
    (deploy_root / "models").mkdir()
    locales_dir = internal_root / "docwen" / "i18n" / "locales"
    locales_dir.mkdir(parents=True)
    (locales_dir / "zh_CN.toml").write_text("[meta]\n", encoding="utf-8")

    import docwen_runtime.resources.registry as registry

    monkeypatch.setattr(registry.sys, "_MEIPASS", str(internal_root), raising=False)

    root = find_project_root()
    resources = ResourceRegistry(root)

    assert root == deploy_root
    assert resources.templates_dir() == deploy_root / "templates"
    assert resources.configs_dir() == deploy_root / "configs"
    assert resources.locales_dir() == locales_dir
