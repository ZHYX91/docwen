"""Resource-preflight contracts shared by packaged CLI and GUI verifiers."""

from __future__ import annotations

from pathlib import Path
from typing import Literal

import pytest
from tests.support.release_packaging import (
    use_compact_pymupdf_layout_manifest,
    write_packaged_common_resources,
    write_packaged_gui_assets,
)

pytestmark = pytest.mark.unit

_CommonMutation = Literal["all-locales", "one-locale", "template", "config", "models"]


@pytest.fixture(autouse=True)
def _use_compact_manifest(monkeypatch: pytest.MonkeyPatch) -> None:
    use_compact_pymupdf_layout_manifest(monkeypatch)


def _write_binary(binary_dir: Path, binary_name: str) -> None:
    binary_dir.mkdir()
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")


def _remove_common_component(binary_dir: Path, mutation: _CommonMutation, *, gui: bool) -> None:
    if mutation == "all-locales":
        locale_dir = binary_dir / "_internal" / "docwen" / "i18n" / "locales"
        for locale_path in locale_dir.glob("*"):
            locale_path.unlink()
    elif mutation == "one-locale":
        (binary_dir / "_internal" / "docwen" / "i18n" / "locales" / "en_US.toml").unlink()
    elif mutation == "template":
        (binary_dir / "templates" / "English General Template.docx").unlink()
    elif mutation == "config":
        relative_path = Path("proofread/engine.toml") if gui else Path("numbering/add.toml")
        (binary_dir / "configs" / relative_path).unlink()
    else:
        for model_path in (binary_dir / "models").rglob("*"):
            if model_path.is_file():
                model_path.unlink()


def test_packaged_cli_verifier_fails_when_resources_are_missing(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_cli

    binary_dir = tmp_path / "dist"
    binary_name = "DocWenCLI.exe"
    _write_binary(binary_dir, binary_name)

    with pytest.raises(RuntimeError, match="packaged_cli_resources_missing"):
        verify_packaged_cli.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name])


@pytest.mark.parametrize(
    ("mutation", "expected_error"),
    [
        ("all-locales", "packaged_cli_locales_missing"),
        ("one-locale", "packaged_cli_locales_missing"),
        ("template", "packaged_cli_templates_missing"),
        ("config", "packaged_cli_configs_missing"),
        ("models", "packaged_cli_models_missing"),
    ],
)
def test_packaged_cli_verifier_rejects_incomplete_common_resources(
    tmp_path: Path,
    mutation: _CommonMutation,
    expected_error: str,
) -> None:
    from scripts.release import verify_packaged_cli

    binary_dir = tmp_path / "dist"
    binary_name = "DocWenCLI.exe"
    _write_binary(binary_dir, binary_name)
    write_packaged_common_resources(binary_dir)
    _remove_common_component(binary_dir, mutation, gui=False)

    with pytest.raises(RuntimeError, match=expected_error):
        verify_packaged_cli.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name])


@pytest.mark.parametrize(
    ("mutation", "expected_error"),
    [
        ("all-locales", "packaged_gui_locales_missing"),
        ("one-locale", "packaged_gui_locales_missing"),
        ("template", "packaged_gui_templates_missing"),
        ("config", "packaged_gui_configs_missing"),
        ("models", "packaged_gui_models_missing"),
    ],
)
def test_packaged_gui_verifier_rejects_incomplete_common_resources(
    tmp_path: Path,
    mutation: _CommonMutation,
    expected_error: str,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "dist"
    binary_name = "DocWen.exe"
    _write_binary(binary_dir, binary_name)
    write_packaged_common_resources(binary_dir)
    write_packaged_gui_assets(binary_dir)
    _remove_common_component(binary_dir, mutation, gui=True)

    with pytest.raises(RuntimeError, match=expected_error):
        verify_packaged_gui.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name])


def test_packaged_gui_verifier_fails_when_assets_are_missing(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "dist"
    binary_name = "DocWen.exe"
    _write_binary(binary_dir, binary_name)
    write_packaged_common_resources(binary_dir)

    with pytest.raises(RuntimeError, match="packaged_gui_assets_missing"):
        verify_packaged_gui.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name])


def test_packaged_gui_asset_manifest_matches_current_runtime_assets() -> None:
    from scripts.release import verify_packaged_gui

    assets_dir = Path("assets")
    runtime_assets = {
        asset_path.relative_to(assets_dir).as_posix()
        for asset_path in assets_dir.rglob("*")
        if asset_path.is_file() and "screenshots" not in asset_path.relative_to(assets_dir).parts
    }

    assert set(verify_packaged_gui._REQUIRED_ASSET_FILES) == runtime_assets
