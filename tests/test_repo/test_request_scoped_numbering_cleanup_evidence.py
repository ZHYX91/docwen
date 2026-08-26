"""Fail-closed evidence guards for VIS-2026-07-19-144 request policy."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "request-scoped-numbering-cleanup-2026-07-19.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def test_runtime_builds_numbering_and_cleanup_from_the_same_request_snapshot() -> None:
    core = _read("packages/core/src/docwen_core/text/heading_numbering.py")
    builder = _read("packages/runtime/src/docwen_runtime/config/heading_cleanup.py")
    registry = _read("packages/runtime/src/docwen_runtime/numbering/registry.py")
    manager = _read("packages/runtime/src/docwen_runtime/engine/task_manager.py")
    context = _read("packages/runtime/src/docwen_runtime/_execution_context.py")
    loader = _read("packages/runtime/src/docwen_runtime/config/loader.py")

    assert "def compile_clean_rules_from_data(" in core
    assert "rules: Sequence[HeadingCleanupRule]," in core
    assert "detect_heading_prefix(text, rules=rules)" in core
    for removed_global_api in (
        "_INJECTED_RULES",
        "def _get_strip_rules(",
        "def reload_clean_rules(",
        "def set_clean_rules(",
        "def set_clean_rules_from_data(",
    ):
        assert removed_global_api not in core
    assert "def build_heading_cleanup_rules(" in builder
    assert "return compile_clean_rules_from_data(ordered)" in builder

    assert "def from_config_snapshot(" in registry
    assert "def with_config_snapshot(" in registry
    assert "snapshot = deepcopy(dict(config_snapshot))" in registry
    assert "get_config_loader" not in registry
    assert "ResourceRegistry" not in registry
    assert "def default(" not in registry
    assert "numbering_registry.with_config_snapshot(" in manager
    assert 'locale=request.options.get("locale")' in manager
    assert "build_heading_cleanup_rules(request.config_snapshot)" in manager
    assert "heading_cleanup_rules=heading_cleanup_rules" in manager
    assert "def heading_cleanup_rules(self)" in context

    assert "_inject_clean_rules" not in loader
    assert "if not clean_data:" not in loader


def test_all_conversion_time_heading_consumers_receive_explicit_request_rules() -> None:
    markdown_common = _read("packages/plugins/markdown/src/docwen_plugin_markdown/common_utils.py")
    markdown_numbering = _read("packages/plugins/markdown/src/docwen_plugin_markdown/numbering/converter.py")
    markdown_docx = _read("packages/plugins/markdown/src/docwen_plugin_markdown/to_docx/converter.py")
    document = _read("packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py")
    gongwen_plugin = _read("packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/plugin.py")
    gongwen_pipeline = _read("packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/pipeline.py")
    gongwen_reader = _read(
        "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/extraction/paragraph_reader.py"
    )
    gongwen_yaml = _read(
        "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/recognition/yaml_builder.py"
    )
    gongwen_utils = _read("packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/utils.py")

    assert "def remove_md_numbering(content: str, *, rules: Any = ())" in markdown_common
    assert "strip_heading_prefix(heading_text, rules=rules)" in markdown_common
    assert 'getattr(context, "heading_cleanup_rules", ()) or ()' in markdown_numbering
    assert "content = remove_md_numbering(content, rules=cleanup_rules)" in markdown_docx
    assert "heading_cleanup_rules=heading_cleanup_rules" in document
    assert "rules=heading_cleanup_rules" in document

    assert 'cleanup_rules=getattr(context, "heading_cleanup_rules", ()) or ()' in gongwen_plugin
    assert gongwen_pipeline.count("cleanup_rules=cleanup_rules") == 2
    assert "detect_heading_prefix(text, rules=cleanup_rules)" in gongwen_reader
    assert "cleanup_rules=cleanup_rules" in gongwen_yaml
    assert "strip_heading_prefix(t, rules=cleanup_rules)" in gongwen_utils


def test_request_isolation_regressions_remain_discoverable() -> None:
    runtime_tests = _read("packages/runtime/tests/test_request_scoped_numbering.py")
    core_tests = _read("packages/core/tests/text/test_heading_numbering.py")
    markdown_tests = _read("packages/plugins/markdown/tests/test_md_numbering.py")
    markdown_docx_tests = _read("packages/plugins/markdown/tests/test_md_to_docx_numbering_*.py")
    document_tests = _read("packages/plugins/document/tests/test_to_markdown_standard_parity_*.py")
    gongwen_tests = _read("packages/plugins/optimizers/gongwen/tests/test_gongwen_extraction.py")

    for test_name in (
        "test_task_manager_rebuilds_numbering_and_cleanup_from_each_request_snapshot",
        "test_concurrent_requests_keep_numbering_and_cleanup_snapshots_isolated",
        "test_two_config_loaders_cannot_rebind_a_frozen_request_snapshot",
        "test_config_loader_reload_does_not_publish_process_global_cleanup_rules",
        "test_task_manager_passes_explicit_empty_cleanup_without_fallback",
    ):
        assert f"def {test_name}(" in runtime_tests
    assert "Barrier(2)" in runtime_tests
    assert 'config_snapshot=_numbering_snapshot("shared", marker)' in runtime_tests

    assert "test_independent_rule_sets_do_not_share_state" in core_tests
    assert "test_request_cleanup_rules_are_the_only_rules_used" in markdown_tests
    assert "test_md_to_docx_text_and_word_native_use_request_cleanup_rules" in markdown_docx_tests
    assert "test_docx_to_md_full_convert_uses_request_cleanup_rules" in document_tests
    assert "test_request_cleanup_rules_are_the_only_rules_used" in gongwen_tests
    assert "test_runtime_gongwen_route_uses_request_cleanup_for_headings_and_attachment_yaml" in gongwen_tests
