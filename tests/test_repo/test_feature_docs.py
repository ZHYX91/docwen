"""Current capability inventory guards."""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
CAPABILITIES = ROOT / "docs" / "capabilities.md"


def _row(feature_id: str) -> str:
    prefix = f"| {feature_id} |"
    return next(line for line in CAPABILITIES.read_text(encoding="utf-8").splitlines() if line.startswith(prefix))


def test_capability_inventory_is_unique_and_current_only() -> None:
    text = CAPABILITIES.read_text(encoding="utf-8")
    rows = [line for line in text.splitlines() if line.startswith("| FEAT-")]
    ids = [line.split("|", 2)[1].strip() for line in rows]

    assert len(rows) == 160
    assert len(set(ids)) == len(ids)
    assert not re.search(r"\bVIS-\d|\bF-\d", text)


def test_capability_inventory_tracks_current_owners_and_entrypoints() -> None:
    expected = {
        "FEAT-CLI-001": ("docwen_cli", "commands/execution_v3.py + commands/convert.py"),
        "FEAT-CLI-018": ("docwen_cli + docwen_runtime", "i18n.py + i18n/locales"),
        "FEAT-CLI-021": ("docwen_cli + docwen_bundle + docwen_gui", "commands/gui_control.py"),
        "FEAT-CONV-002": ("docwen_plugin_markdown", "to_docx/converter.py"),
        "FEAT-CONV-003": ("docwen_plugin_spreadsheet", "to_markdown/converter.py"),
        "FEAT-CONV-005": ("docwen_plugin_image", "format_conversion/converter.py"),
        "FEAT-CONV-022": ("docwen_plugin_markup", "publication/converter.py"),
        "FEAT-CONV-023": ("docwen_plugin_presentation", "pptx_md/converter.py"),
        "FEAT-CONV-027": ("docwen_plugin_print", "paged_output/converter.py"),
        "FEAT-RES-006": ("docwen_bundle", "docwen_bundle.cli_entry:main"),
        "FEAT-RES-007": ("docwen_bundle", "docwen_bundle.gui_entry:main"),
    }
    for feature_id, tokens in expected.items():
        row = _row(feature_id)
        for token in tokens:
            assert token in row


def test_workspace_and_import_boundaries_match_the_executable_configuration() -> None:
    pyproject = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))

    assert len(pyproject["tool"]["uv"]["workspace"]["members"]) == 17
    assert len(pyproject["tool"]["uv"]["sources"]) == 17
    assert len(pyproject["tool"]["importlinter"]["contracts"]) == 8
    assert "17 packages" in _row("FEAT-ARCH-001")
    assert "8 current contracts" in _row("FEAT-ARCH-009")


def test_capability_implementation_examples_resolve_to_current_files() -> None:
    paths = (
        "packages/plugins/markup/src/docwen_plugin_markup/publication/converter.py",
        "packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/converter.py",
        "packages/plugins/layout/src/docwen_plugin_layout/to_markdown/converter.py",
        "packages/plugins/print/src/docwen_plugin_print/paged_output/converter.py",
        "packages/apps/cli/src/docwen_cli/commands/convert.py",
        "packages/bundle/src/docwen_bundle/runtime_factory.py",
    )
    for relative_path in paths:
        assert (ROOT / relative_path).is_file()
