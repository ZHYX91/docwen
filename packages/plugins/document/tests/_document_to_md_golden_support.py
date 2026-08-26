"""Golden parity test for ROUTE-DOC-001 (DOCX → Markdown).

Uses both a generated sample DOCX (with known content) and a real
DOCX template to validate conversion output through the full runtime
pipeline (plugin → workspace → finalizer).

Golden case: GOLDEN-002
Route: ROUTE-DOC-001 (document → md)
Feature: FEAT-CONV-001 (DocxToMarkdownStrategy)

Comparison strategy (per Golden验收用例与比较策略.md):
- Paragraph text comparison
- Heading level verification
- Image count check
- Table structure comparison
- Acceptable differences: whitespace, image path differences
- Required: text content, heading levels, table data consistency
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy

pytestmark = [pytest.mark.golden, pytest.mark.contract]

PROJECT_ROOT = Path(__file__).resolve().parents[4]

_DOCX_TO_MD_OLD_SYSTEM_FIXTURE = (
    PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_docx_to_markdown_semantics.json"
)

_OLD_SYSTEM_COMPREHENSIVE_DOCX = (
    PROJECT_ROOT / "tests" / "fixtures" / "golden" / "md_to_docx_old" / "sample_golden.docx"
)


def _build_runtime_pipeline():
    """Build the complete runtime pipeline with the real DocumentPlugin."""
    from docwen_plugin_document import DocumentPlugin
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    plugin = DocumentPlugin()
    registry = PluginRegistry()
    registry.register(plugin)

    resolver = RouteResolver(registry)
    ws_root = tempfile.mkdtemp(prefix="docwen_golden_")
    ws_mgr = WorkspaceManager(root_dir=ws_root)
    finalizer = OutputFinalizer()
    task_mgr = TaskManager(registry, resolver, ws_mgr, finalizer)

    return plugin, task_mgr, ws_mgr, ws_root


def _run_conversion(
    task_mgr,
    input_path,
    output_dir,
    *,
    config_snapshot: dict[str, Any] | None = None,
    **options,
) -> Any:
    """Run a single conversion through the task manager."""
    opts = {"to_md_keep_images": True, "remove_numbering": True}
    opts.update(options)

    request = ConversionRequest(
        request_id="golden-test",
        input_refs=[
            FileRef(
                path=str(input_path),
                format="docx",
                category="document",
                size_bytes=Path(input_path).stat().st_size,
            )
        ],
        target_format="md",
        output_policy=OutputPolicy(output_dir=str(output_dir)),
        options=opts,
        config_snapshot=config_snapshot or {},
    )
    return task_mgr.execute_single(request)


def _load_docx_to_md_old_system_fixture() -> dict[str, Any]:
    return json.loads(_DOCX_TO_MD_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _heading_counts(markdown: str) -> dict[str, int]:
    counts: dict[str, int] = {}
    for line in markdown.splitlines():
        stripped = line.strip()
        if not stripped.startswith("#"):
            continue
        level = 0
        for ch in stripped:
            if ch == "#":
                level += 1
            else:
                break
        if 1 <= level <= 6 and level < len(stripped) and stripped[level] == " ":
            key = str(level)
            counts[key] = counts.get(key, 0) + 1
    return counts


def _document_node_root(path: Path, output_dir: Path) -> Path:
    relative = path.relative_to(output_dir)
    assert len(relative.parts) >= 2
    root = output_dir / relative.parts[0]
    assert root.is_dir()
    return root


__all__ = (
    "_OLD_SYSTEM_COMPREHENSIVE_DOCX",
    "ConversionRequest",
    "FileRef",
    "OutputPolicy",
    "Path",
    "_build_runtime_pipeline",
    "_document_node_root",
    "_heading_counts",
    "_load_docx_to_md_old_system_fixture",
    "_run_conversion",
    "os",
    "pytest",
    "pytestmark",
    "re",
)
