"""Fail-closed evidence guards for VIS-2026-07-20-146 DOCX request policy."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "docx-markdown-request-policy-and-converter-isolation-2026-07-20.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def _function_source(relative_path: str, function_name: str) -> str:
    source = _read(relative_path)
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == function_name:
            return ast.get_source_segment(source, node) or ""
    raise AssertionError(f"missing {function_name} in {relative_path}")


def _method_source(relative_path: str, class_name: str, method_name: str) -> str:
    source = _read(relative_path)
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == class_name:
            for member in node.body:
                if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and member.name == method_name:
                    return ast.get_source_segment(source, member) or ""
    raise AssertionError(f"missing {class_name}.{method_name} in {relative_path}")


def test_public_document_route_owns_a_fresh_standard_converter() -> None:
    relative_path = "packages/plugins/document/src/docwen_plugin_document/plugin.py"
    source = _read(relative_path)
    constructor = _method_source(relative_path, "DocumentPlugin", "__init__")
    convert = _method_source(relative_path, "DocumentPlugin", "convert")

    assert "return DocxToMarkdownConverter().convert(context)" in convert
    assert "self._converter" not in source
    assert "DocxToMarkdownConverter()" not in constructor
    assert "self._smart_converter" in source


def test_request_snapshot_policy_is_authoritative_without_global_fallback() -> None:
    relative_path = "packages/plugins/document/src/docwen_plugin_document/to_markdown/request_policy.py"
    source = _read(relative_path)
    build = _function_source(relative_path, "build_docx_markdown_request_policy")
    for token in (
        "@dataclass(frozen=True, slots=True)",
        "class DocxMarkdownRequestPolicy:",
        'conversion = _section(context.config.get("conversion", {}))',
        'document = _section(context.config.get("document", {}))',
        'export_config = _section(context.config.get("export", {}))',
        "MarkdownExportSemantics.from_config(",
        "formatting = docx_markdown_formatting_config_from_conversion_config(conversion)",
        "syntax = docx_markdown_syntax_config_from_conversion_config(conversion)",
        "style_detector = style_detector_config_from_document_config(document)",
        "style_detector=style_detector",
    ):
        assert token in source

    for global_getter in (
        "get_markdown_export_semantics()",
        "get_docx_markdown_formatting_config()",
        "get_docx_markdown_syntax_config()",
        "get_docx_style_detector_config()",
    ):
        assert global_getter not in build

    assert "export_cfg=export_config" in build
    assert 'document.get("to_md_image_extraction_mode")' not in build
    assert 'document.get("to_md_ocr_placement_mode")' not in build
    assert 'document.get("to_md_table_merge_export_strategy")' in build
    assert 'export_config.get("to_md_table_merge_export_strategy")' not in build


def test_explicit_option_precedence_and_mode_compatibility_remain_fail_closed() -> None:
    relative_path = "packages/plugins/document/src/docwen_plugin_document/to_markdown/request_policy.py"
    source = _read(relative_path)
    build = _function_source(relative_path, "build_docx_markdown_request_policy")
    option_value = _function_source(relative_path, "_option_value")
    option_nonblank = _function_source(relative_path, "_option_nonblank")
    resolve_modes = _method_source(relative_path, "DocxMarkdownRequestPolicy", "resolve_export_modes")

    assert "if key not in options or options[key] is None:" in option_value
    assert "return options[key]" in option_value
    assert "isinstance(value, str) and not value.strip()" in option_nonblank
    assert '_option_nonblank(\n            options,\n            "image_mode"' in build
    assert '_option_nonblank(\n            options,\n            "ocr_placement"' in build
    assert 'if image_extraction_mode.strip().lower() == "base64":' in build
    assert 'ocr_placement_mode = "main_md"' in build
    assert "self.export.image_extraction_mode" in resolve_modes
    assert "self.export.ocr_placement_mode" in resolve_modes
    assert "options" not in resolve_modes.splitlines()[0]
    for option_name in (
        "preserve_formatting",
        "preserve_heading_formatting",
        "preserve_table_header_formatting",
        "page_break_separator",
        "section_break_separator",
        "horizontal_rule_separator",
    ):
        assert f'"{option_name}"' in source


def test_converter_applies_one_policy_before_parsing_and_propagates_it() -> None:
    relative_path = "packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py"
    source = _read(relative_path)
    convert = _method_source(relative_path, "DocxToMarkdownConverter", "convert")
    convert_once = _method_source(relative_path, "DocxToMarkdownConverter", "_convert_once")
    parse = _method_source(relative_path, "DocxToMarkdownConverter", "_parse_docx")
    images = _method_source(relative_path, "DocxToMarkdownConverter", "_process_images_for_output")
    ocr = _method_source(relative_path, "DocxToMarkdownConverter", "_ocr_per_image")

    assert "with self._conversion_lock:" in convert
    assert "return self._convert_once(context)" in convert
    assert "policy = build_docx_markdown_request_policy(context, options)" in convert_once
    assert "self._apply_request_policy(policy)" in convert_once
    assert convert_once.index("self._apply_request_policy(policy)") < convert_once.index("self._parse_docx(")
    assert "if not self._request_policy_resolved:" in parse
    assert "build_docx_markdown_request_policy(context, context.request.options)" in parse
    assert 'getattr(context, "heading_cleanup_rules", ()) or ()' in parse
    assert "style_detector_config = self._request_policy.style_detector" in parse
    assert "export_modes = self._request_policy.resolve_export_modes()" in parse

    for token in (
        "syntax_config=self._syntax_for_rendering()",
        "preserve_formatting=preserve_formatting",
        "self._request_policy.resolve_export_modes()",
        "self._request_policy.image_link_style",
        "export_semantics=self._request_policy.export",
        "self._request_policy.export.md_file_link_style",
        "self._request_policy.ocr_blockquote_title",
    ):
        assert token in source
    assert "export_semantics=self._request_policy.export" in images
    assert "get_base64_export_semantics" not in images
    assert "get_markdown_asset_link_semantics" not in images
    assert "get_ocr_blockquote_title" not in ocr
    assert "ocr_blockquote_title=" not in ocr
    assert "_format_main_ocr_blockquote(" in ocr
    assert "Merge exactly once after OCR" in parse


def test_shared_run_and_base64_helpers_accept_explicit_request_policy() -> None:
    runs = _read("packages/plugins/document/src/docwen_plugin_document/shared/markdown_runs.py")
    images = _read("packages/core/src/docwen_core/text/image_markdown.py")
    export = _read("packages/core/src/docwen_core/export_semantics/__init__.py")

    for token in (
        "def append_formatted_run_text(",
        "syntax_config: DocxMarkdownSyntaxConfig,",
        "append_formatted_run_text(",
        "syntax_config=syntax_config",
    ):
        assert token in runs
    for token in (
        "export_semantics: MarkdownExportSemantics,",
        "if not isinstance(export_semantics, MarkdownExportSemantics):",
        "compress_enabled = export_semantics.export_base64_compress_enabled",
        "threshold_kb = export_semantics.export_base64_compress_threshold_kb",
        "export_semantics=export_semantics",
    ):
        assert token in images
    assert "configure_export_semantics" not in export
    assert "get_markdown_export_semantics" not in export


def test_docx_policy_types_have_no_process_global_mutation_surface() -> None:
    features = _read("packages/core/src/docwen_core/docx_parsing/format_features.py")
    runs = _read("packages/plugins/document/src/docwen_plugin_document/shared/markdown_runs.py")

    for forbidden in (
        "_current_style_detector_config",
        "_current_docx_markdown_formatting_config",
        "_current_docx_markdown_syntax_config",
        "configure_docx_style_detector_config",
        "configure_docx_markdown_formatting_config",
        "configure_docx_markdown_syntax_config",
        "get_docx_style_detector_config",
        "get_docx_markdown_formatting_config",
        "get_docx_markdown_syntax_config",
        "runtime_policy_transaction",
    ):
        assert forbidden not in features
    assert "syntax_config: DocxMarkdownSyntaxConfig," in runs
    assert "syntax_config or" not in runs


def test_runtime_injects_locale_aware_ocr_title_from_the_request_snapshot() -> None:
    helper = _read("packages/runtime/src/docwen_runtime/config/ocr_output.py")
    task_manager = _read("packages/runtime/src/docwen_runtime/engine/task_manager.py")
    context = _read("packages/runtime/src/docwen_runtime/_execution_context.py")
    protocol = _read("packages/core/src/docwen_core/protocols/execution_context.py")
    policy = _read("packages/plugins/document/src/docwen_plugin_document/to_markdown/request_policy.py")
    runtime_tests = _read("packages/runtime/tests/test_i18n.py")

    for token in (
        "def build_ocr_blockquote_title(",
        'ocr_output.get("show_blockquote_title", True)',
        'ocr_output.get("blockquote_title_override_by_locale")',
        'locales_dir / f"{locale}.toml"',
        'ocr_table.get("blockquote_prefix")',
        "_LOCALE_CODE.fullmatch(candidate)",
    ):
        assert token in helper
    assert "ocr_blockquote_title=build_ocr_blockquote_title(" in task_manager
    assert 'requested_locale=request.options.get("locale")' in task_manager
    assert "def ocr_blockquote_title(self) -> str:" in context
    assert "def ocr_blockquote_title(self) -> str:" in protocol
    assert "ocr_blockquote_title = context.ocr_blockquote_title.strip()" in policy
    assert "export.ocr_blockquote_title_override_text" not in policy
    assert "def test_task_manager_injects_request_ocr_title_into_execution_context(" in runtime_tests


def test_request_policy_integration_cases_remain_discoverable() -> None:
    policy_tests = _read("packages/plugins/document/tests/test_request_scoped_docx_policy_*.py")
    helper_tests = _read("packages/plugins/document/tests/test_shared_markdown_runs.py")
    base64_tests = _read("packages/core/tests/test_image_markdown_base64.py")

    for test_name in (
        "test_context_snapshot_owns_all_docx_policy",
        "test_warmed_document_plugin_keeps_parallel_request_policies_isolated",
        "test_captured_request_snapshot_is_stable_until_execution",
        "test_export_owner_ignores_removed_document_mode_duplicates",
        "test_document_table_merge_owner_precedes_conversion_compatibility_value",
        "test_blank_mode_options_preserve_request_snapshot_modes",
        "test_request_policy_freezes_export_options_at_projection_time",
        "test_partial_nonempty_snapshot_uses_deterministic_defaults",
        "test_empty_snapshot_uses_deterministic_defaults",
        "test_explicit_options_preserve_false_ignore_and_omit_presentation",
        "test_formula_text_uses_request_syntax_instead_of_hardcoded_markers",
        "test_formula_text_respects_explicit_preserve_formatting_false",
        "test_request_syntax_reaches_sdt_and_nested_table_paths",
        "test_converter_base64_path_consumes_explicit_request_compression_policy",
        "test_converter_ocr_sidecar_uses_request_link_without_main_title",
        "test_converter_main_ocr_uses_request_title_policy",
        "test_converter_main_ocr_omits_disabled_title",
        "test_ocr_title_override_uses_request_locale",
        "test_converter_main_ocr_uses_runtime_injected_localized_fallback",
    ):
        assert f"def {test_name}(" in policy_tests
    assert "def test_append_formatted_run_text_honors_explicit_syntax_config(" in helper_tests
    assert "def test_render_paragraph_runs_coalesces_adjacent_runs_through_shared_helper(" in helper_tests
    assert "def test_base64_explicit_semantics_are_the_only_compression_source(" in base64_tests
    assert "def test_base64_rejects_missing_request_semantics(" in base64_tests
