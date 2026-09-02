from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]


def _read(relative: str) -> str:
    return read_source_text(ROOT / relative)


def test_runtime_has_no_process_global_markdown_export_policy() -> None:
    export = _read("packages/core/src/docwen_core/export_semantics/__init__.py")
    loader = _read("packages/runtime/src/docwen_runtime/config/loader.py")
    task_manager = _read("packages/runtime/src/docwen_runtime/engine/task_manager.py")

    for removed_api in (
        "_current_semantics",
        "configure_export_semantics",
        "get_markdown_export_semantics",
    ):
        assert removed_api not in export
        assert removed_api not in loader
        assert removed_api not in task_manager
    unrelated_snapshot_services = task_manager.index("if request.config_snapshot:")
    context = task_manager.index("return RuntimeExecutionContext(", unrelated_snapshot_services)
    projection = task_manager.index("markdown_export_semantics=MarkdownExportSemantics.from_config_snapshot(")
    assert unrelated_snapshot_services < context < projection


def test_empty_and_nonempty_snapshots_share_the_pure_projection_path() -> None:
    core_tests = _read("packages/plugins/document/tests/test_request_scoped_docx_policy_*.py")
    runtime_tests = _read("packages/runtime/tests/test_request_scoped_markdown_export.py")

    assert "test_context_snapshot_owns_all_docx_policy" in core_tests
    assert "test_captured_request_snapshot_is_stable_until_execution" in core_tests
    assert "test_partial_nonempty_snapshot_uses_deterministic_defaults" in core_tests
    assert "test_empty_snapshot_uses_deterministic_defaults" in core_tests
    assert "test_empty_snapshot_projects_deterministic_request_defaults" in runtime_tests
