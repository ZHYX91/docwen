"""Golden tests for ROUTE-SHEET-001: XLSX/CSV → Markdown.

Covers GOLDEN-003: XLSX→MD with multi-sheet, merged cells, blocks.
"""

from __future__ import annotations

import json
import os
import tempfile
from collections.abc import Generator
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.golden

PROJECT_ROOT = Path(__file__).resolve().parents[4]

_XLSX_TO_MD_OLD_SYSTEM_FIXTURE = (
    PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_xlsx_to_markdown_semantics.json"
)


def _build_fake_context(
    input_path: str,
    staging_dir: str,
    options: dict[str, Any] | None = None,
    *,
    source_format: str | None = None,
) -> Any:
    """Build a fake PluginExecutionContext for direct plugin testing."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    token = CancellationToken()
    file_ref = FileRef(
        path=input_path,
        format=source_format or Path(input_path).suffix.lstrip("."),
        category="spreadsheet",
    )
    request = ConversionRequest(
        request_id="test-001",
        input_refs=[file_ref],
        target_format="md",
        options=options or {},
        output_policy=OutputPolicy(),
    )
    workspace = FakeWorkspaceHandle(input_path, staging_dir)
    progress = FakeProgressSink()
    config = FakeConfigView()
    logger = FakePluginLogger()

    return FakeExecutionContext(
        request=request,
        workspace=workspace,
        config=config,
        progress=progress,
        cancellation=token,
        logger=logger,
    )


def _build_runtime_pipeline() -> tuple[Any, Any]:
    """Build a full runtime pipeline for integration testing."""
    from docwen_plugin_spreadsheet.plugin import SpreadsheetPlugin

    plugin = SpreadsheetPlugin()
    return plugin, None  # PluginRegistry is created in individual tests


def _load_xlsx_to_md_old_system_fixture() -> dict[str, Any]:
    return json.loads(_XLSX_TO_MD_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _deliverable_artifacts(result: Any) -> list[Any]:
    manifests = [
        artifact for artifact in result.artifacts if artifact.media_type == "application/vnd.docwen.document-node+json"
    ]
    assert len(manifests) == 1
    return [artifact for artifact in result.artifacts if artifact not in manifests]


def _document_node_root(path: Path, output_dir: Path) -> Path:
    relative = path.relative_to(output_dir)
    assert len(relative.parts) >= 2
    root = output_dir / relative.parts[0]
    assert root.is_dir()
    return root


def _io_path(path: str | Path) -> Path:
    from docwen_runtime.path_io import filesystem_path

    return filesystem_path(str(path))


def _legacy_markdown_projection(content: str, artifacts: list[Any]) -> str:
    for artifact in artifacts:
        if artifact.media_type != "text/markdown" or artifact.is_primary or artifact.logical_path is None:
            continue
        legacy_name = artifact.metadata.get("source_suggested_name")
        if isinstance(legacy_name, str):
            content = content.replace(artifact.logical_path.split("/", 1)[1], legacy_name)
    return content.replace("![[../", "![[").replace("](../", "](")


__all__ = (
    "Any",
    "Generator",
    "Path",
    "_build_fake_context",
    "_deliverable_artifacts",
    "_document_node_root",
    "_io_path",
    "_legacy_markdown_projection",
    "_load_xlsx_to_md_old_system_fixture",
    "os",
    "pytest",
    "pytestmark",
    "tempfile",
)
