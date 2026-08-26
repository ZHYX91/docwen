"""Repository checks for current CLI documentation consistency."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _feature_row(feature_id: str) -> str:
    text = (PROJECT_ROOT / "docs" / "capabilities.md").read_text(encoding="utf-8")
    prefix = f"| {feature_id} "
    for line in text.splitlines():
        if line.startswith(prefix):
            return line
    raise AssertionError(f"{feature_id} row not found")


def test_cli_docs_describe_the_current_public_commands() -> None:
    cli = (PROJECT_ROOT / "docs" / "cli.md").read_text(encoding="utf-8")
    row = _feature_row("FEAT-CLI-001")

    for token in (
        "`info`",
        "`inspect FILE`",
        "`resources list|show TYPE`",
        "`schema [COMMAND_PATH...]`",
        "`convert FILE --to FORMAT --output PATH`",
        "`validate FILE [--report PATH]`",
        "`config reset GROUP --yes`",
        "`doctor`",
    ):
        assert token in cli
    assert "`convert FILE --to FORMAT --output PATH`" in row
    assert "`run` 命令" not in row


def test_cli_docs_link_the_machine_readable_schema() -> None:
    cli = (PROJECT_ROOT / "docs" / "cli.md").read_text(encoding="utf-8")
    schema_doc = (PROJECT_ROOT / "docs" / "specs" / "json-contracts.md").read_text(encoding="utf-8")

    assert "specs/json-contracts.schema.json" in cli
    assert "json-contracts.schema.json" in schema_doc
    assert (PROJECT_ROOT / "docs" / "specs" / "json-contracts.schema.json").is_file()


def test_route_docs_delegate_exact_inventory_to_executable_manifests() -> None:
    routes = (PROJECT_ROOT / "docs" / "specs" / "routes-and-actions.md").read_text(encoding="utf-8")

    assert "Plugin manifests are the executable source of truth" in routes
    assert "docwen inspect --json" in routes
    assert "docwen resources list formats --json" in routes
    assert "../capabilities.md" in routes


def test_root_readme_does_not_document_removed_cli_surfaces() -> None:
    """The public README must not revive protocol-2 command spellings."""
    readme = (PROJECT_ROOT / "README.md").read_text(encoding="utf-8")

    forbidden = (
        "DocWenCLI.exe run",
        "run --action",
        "DocWenCLI.exe list formats",
        "DocWenCLI.exe list actions",
        "`list formats [",
        "`list actions`",
    )
    for stale_surface in forbidden:
        assert stale_surface not in readme
