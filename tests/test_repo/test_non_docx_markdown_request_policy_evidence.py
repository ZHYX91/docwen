"""Fail-closed source inventory and evidence guards for VIS-167 / DEBT-03."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
OWNERS = (
    ROOT / "packages/plugins/image/src",
    ROOT / "packages/plugins/spreadsheet/src",
    ROOT / "packages/plugins/layout/src",
    ROOT / "packages/plugins/markup/src",
)


def _calls(path: Path) -> list[ast.Call]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    return [node for node in ast.walk(tree) if isinstance(node, ast.Call)]


def _call_name(call: ast.Call) -> str:
    function = call.func
    if isinstance(function, ast.Name):
        return function.id
    if isinstance(function, ast.Attribute):
        return function.attr
    return ""


def _keyword_names(call: ast.Call) -> set[str]:
    return {keyword.arg for keyword in call.keywords if keyword.arg is not None}


def test_admitted_non_docx_owners_have_no_direct_global_or_implicit_base64_reads() -> None:
    direct_global_reads: list[str] = []
    missing_policy_arguments: list[str] = []
    for owner in OWNERS:
        for path in owner.rglob("*.py"):
            for call in _calls(path):
                name = _call_name(call)
                location = f"{path.relative_to(ROOT)}:{call.lineno}"
                if name in {"get_markdown_asset_link_semantics", "get_ocr_blockquote_title"}:
                    direct_global_reads.append(location)
                if name == "get_markdown_export_modes" and "semantics" not in _keyword_names(call):
                    missing_policy_arguments.append(location)
                if name in {"build_base64_image_data_uri", "generate_image_markdown"} and (
                    "export_semantics" not in _keyword_names(call)
                ):
                    missing_policy_arguments.append(location)

    assert direct_global_reads == []
    assert missing_policy_arguments == []


def test_runtime_injects_one_generic_policy_and_concurrency_oracle_is_discoverable() -> None:
    task_manager = (ROOT / "packages/runtime/src/docwen_runtime/engine/task_manager.py").read_text(encoding="utf-8")
    context = (ROOT / "packages/runtime/src/docwen_runtime/_execution_context.py").read_text(encoding="utf-8")
    core = (ROOT / "packages/core/src/docwen_core/export_semantics/__init__.py").read_text(encoding="utf-8")
    tests = (ROOT / "packages/runtime/tests/test_request_scoped_markdown_export.py").read_text(encoding="utf-8")

    assert "MarkdownExportSemantics.from_config_snapshot(" in task_manager
    assert "runtime_policy_transaction" not in task_manager
    assert "markdown_export_semantics=MarkdownExportSemantics.from_config_snapshot(" in task_manager
    assert "def markdown_export_semantics(" in context
    assert "def resolve_markdown_request_policy(context: object)" in core
    assert "configure_export_semantics" not in core
    assert "get_markdown_export_semantics" not in core
    assert "test_concurrent_requests_keep_markdown_export_snapshots_isolated" in tests
