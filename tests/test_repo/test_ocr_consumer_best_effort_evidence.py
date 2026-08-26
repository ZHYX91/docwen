"""Fail-closed guards for the typed OCR execution and consumer contract."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
OCR_PATH = PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "text" / "ocr.py"
CONTEXT_PATH = PROJECT_ROOT / "packages" / "runtime" / "src" / "docwen_runtime" / "_execution_context.py"
ADMISSION_PATH = PROJECT_ROOT / "packages" / "runtime" / "src" / "docwen_runtime" / "_request_admission.py"
TASK_MANAGER_PATH = PROJECT_ROOT / "packages" / "runtime" / "src" / "docwen_runtime" / "engine" / "task_manager.py"

EXPECTED_TYPED_CALLERS = {
    "packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py",
    "packages/plugins/image/src/docwen_plugin_image/to_markdown/converter.py",
    "packages/plugins/layout/src/docwen_plugin_layout/preprocess.py",
    "packages/plugins/layout/src/docwen_plugin_layout/to_markdown/converter.py",
    "packages/plugins/markup/src/docwen_plugin_markup/markdown_resources.py",
    "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/pipeline.py",
    "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn/invoice_cn/image_parser.py",
    "packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/converter.py",
    "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/to_markdown/converter.py",
}
EXPECTED_WARNING_OWNERS = {
    "packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py",
    "packages/plugins/image/src/docwen_plugin_image/to_markdown/converter.py",
    "packages/plugins/layout/src/docwen_plugin_layout/to_markdown/converter.py",
    "packages/plugins/markup/src/docwen_plugin_markup/markdown_resources.py",
    "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/pipeline.py",
    "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn/invoice_cn/converter.py",
    "packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/converter.py",
    "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/to_markdown/converter.py",
}


def _read(path: str | Path) -> str:
    resolved = path if isinstance(path, Path) else PROJECT_ROOT / path
    return read_source_text(resolved)


def _definition_source(source: str, name: str) -> str:
    tree = ast.parse(source)
    node = next(
        candidate
        for candidate in ast.walk(tree)
        if isinstance(candidate, (ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)) and candidate.name == name
    )
    segment = ast.get_source_segment(source, node)
    assert segment is not None
    return segment


def _production_plugin_sources() -> dict[str, str]:
    source_root = PROJECT_ROOT / "packages" / "plugins"
    return {
        path.relative_to(PROJECT_ROOT).as_posix(): path.read_text(encoding="utf-8")
        for path in source_root.glob("**/src/**/*.py")
    }


def test_core_has_one_typed_ocr_execution_entry_with_exact_statuses() -> None:
    source = _read(OCR_PATH)
    tree = ast.parse(source)
    status_node = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "OcrStatus")
    statuses = {
        node.targets[0].id: node.value.value
        for node in status_node.body
        if isinstance(node, ast.Assign)
        and len(node.targets) == 1
        and isinstance(node.targets[0], ast.Name)
        and isinstance(node.value, ast.Constant)
    }
    assert statuses == {
        "SUCCESS": "success",
        "NO_TEXT": "no_text",
        "INPUT_MISSING": "input_missing",
        "UNAVAILABLE": "unavailable",
        "MODEL_MISSING": "model_missing",
        "INITIALIZATION_FAILED": "initialization_failed",
        "RECOGNITION_FAILED": "recognition_failed",
    }

    outcome_class = _definition_source(source, "OcrOutcome")
    warning_formatter = _definition_source(source, "format_ocr_best_effort_warning")
    run_outcome = _definition_source(source, "run_ocr_outcome")
    assert "@dataclass(frozen=True, slots=True)" in source
    assert "status: OcrStatus" in outcome_class
    assert "def recognized_text(" in outcome_class
    assert "self.status is OcrStatus.SUCCESS" in outcome_class
    assert "_OCR_OPERATIONAL_FAILURE_DETAILS" in source
    assert "_OCR_RESULT_QUALITY_DETAILS" in source
    for token in (
        "OcrStatus.SUCCESS",
        "OcrStatus.NO_TEXT",
        "OcrStatus.INPUT_MISSING",
        "OcrStatus.UNAVAILABLE",
        "OcrStatus.MODEL_MISSING",
        "OcrStatus.INITIALIZATION_FAILED",
        "OcrStatus.RECOGNITION_FAILED",
        "machine-generated",
        "may have been missed",
        'suffix = f"; {context}" if context else ""',
        'return f"OCR best-effort result: status={normalized.value}; {quality_detail}{suffix}."',
        'return f"OCR best-effort fallback: status={normalized.value}; {detail}{suffix}."',
    ):
        assert token in warning_formatter or token in source
    for token in (
        "input_is_file = image_file.is_file()",
        "OcrStatus.INPUT_MISSING",
        "except _OcrModelFilesMissing",
        "except Exception",
        "OcrStatus.MODEL_MISSING",
        "OcrStatus.INITIALIZATION_FAILED",
        "OcrStatus.RECOGNITION_FAILED",
        "OcrStatus.NO_TEXT",
        "OcrStatus.SUCCESS",
        "with slot.invocation_lock:",
    ):
        assert token in run_outcome
    public_execution_names = {
        node.name
        for node in tree.body
        if isinstance(node, ast.FunctionDef)
        and node.name in {"run_ocr_outcome", "run_ocr", "extract_text_from_image", "extract_text_from_image_outcome"}
    }
    assert public_execution_names == {"run_ocr_outcome"}


def test_all_and_only_routed_ocr_consumers_use_typed_outcomes_and_warnings() -> None:
    sources = _production_plugin_sources()
    typed_callers = {path for path, source in sources.items() if "run_ocr_outcome" in source}
    warning_owners = {path for path, source in sources.items() if "OCR-BEST-EFFORT" in source}
    formatter_callers = {path for path, source in sources.items() if "format_ocr_best_effort_warning" in source}

    assert typed_callers == EXPECTED_TYPED_CALLERS
    assert warning_owners == EXPECTED_WARNING_OWNERS
    assert formatter_callers == EXPECTED_WARNING_OWNERS | {
        "packages/plugins/layout/src/docwen_plugin_layout/preprocess.py",
    }
    for path, source in sources.items():
        assert "_OCR_OPERATIONAL_FAILURE_DETAILS" not in source, path
        assert "_OCR_OPERATIONAL_FAILURES" not in source, path
        assert "outcome.text" not in source, path
    for path in EXPECTED_TYPED_CALLERS:
        assert "recognized_text" in sources[path], path

    prohibited_wrapper_callers: set[str] = set()
    for path, source in sources.items():
        tree = ast.parse(source, filename=path)
        prohibited_wrapper_names = {
            "run_ocr",
            "extract_text_from_image",
            "ocr_extract_text",
            "ocr_extract",
        }
        if any(
            isinstance(node, ast.Call)
            and (
                (isinstance(node.func, ast.Name) and node.func.id in prohibited_wrapper_names)
                or (isinstance(node.func, ast.Attribute) and node.func.attr in prohibited_wrapper_names)
            )
            for node in ast.walk(tree)
        ):
            prohibited_wrapper_callers.add(path)
        preflight_calls = [
            node
            for node in ast.walk(tree)
            if isinstance(node, ast.Call)
            and (
                (isinstance(node.func, ast.Name) and node.func.id == "ocr_available")
                or (isinstance(node.func, ast.Attribute) and node.func.attr == "ocr_available")
            )
        ]
        assert not preflight_calls, path
    assert prohibited_wrapper_callers == set()


def test_runtime_retains_streamed_diagnostics_and_deduplicates_exact_repeats() -> None:
    context_source = _read(CONTEXT_PATH)
    manager_source = _read(TASK_MANAGER_PATH)
    runtime_tests = _read("packages/runtime/tests/test_fake_closed_loop_*.py")
    report_diagnostic = _definition_source(context_source, "report_diagnostic")
    merge = _definition_source(manager_source, "_merge_diagnostics")

    assert "self.diagnostics.append(" in report_diagnostic
    assert "ConversionDiagnostic(" in report_diagnostic
    assert "def reported_diagnostics(" in context_source
    assert manager_source.count("state.runtime_context.reported_diagnostics") >= 4
    for token in (
        "diagnostic.level",
        "diagnostic.message",
        "diagnostic.code",
        "diagnostic.location",
        "if key in seen:",
        "seen.add(key)",
    ):
        assert token in merge
    for test_name in (
        "test_streamed_diagnostic_is_merged_and_deduplicated",
        "test_streamed_diagnostic_survives_plugin_exception",
    ):
        assert f"def {test_name}(" in runtime_tests
    assert 'pytest.param("cancel", "cancelled"' in runtime_tests
    assert 'pytest.param("runtime", "conversion_failed"' in runtime_tests


def test_public_task_manager_uses_pure_idempotent_snapshot_admission() -> None:
    admission = _read(ADMISSION_PATH)
    manager = _read(TASK_MANAGER_PATH)
    tests = _read("packages/runtime/tests/test_task_manager_ocr_option_projection.py")
    execute_single = _definition_source(manager, "execute_single")

    for token in (
        'if "ocr_language" not in options:',
        'options["ocr_language"] = ocr_language or "auto"',
        'if "locale" not in options:',
        'options["locale"] = locale or "zh_CN"',
        "return replace(",
    ):
        assert token in admission
    assert "ConfigLoader" not in admission
    assert "admit_markdown_ocr_options(request, request.config_snapshot)" in execute_single
    assert "config_loader" not in execute_single.casefold()
    for test_name in (
        "test_direct_single_projects_snapshot_values_without_mutating_caller",
        "test_direct_single_preserves_present_falsey_keys",
        "test_direct_single_keeps_projection_scope",
        "test_direct_batch_projects_each_derived_request",
    ):
        assert f"def {test_name}(" in tests


def test_typed_ocr_paths_and_best_effort_regressions_are_guarded() -> None:
    expected_tests = {
        "packages/core/tests/test_ocr_*.py": (
            "test_run_ocr_outcome_reports_missing_input_without_initializing_engine",
            "test_run_ocr_outcome_reports_missing_model",
            "test_run_ocr_outcome_classifies_model_directory_resolution_error_as_initialization_failure",
            "test_format_ocr_best_effort_warning_is_canonical_and_safe",
            "test_format_ocr_best_effort_warning_ignores_unknown_status",
            "test_ocr_outcome_exposes_text_only_for_success",
        ),
        "packages/plugins/image/tests/test_image_conversions_*.py": (
            "test_image_to_markdown_ocr_failure_is_best_effort",
            "test_image_to_markdown_typed_ocr_failures_are_safe_and_nonfatal",
            "test_image_to_markdown_no_text_warns_about_possible_missed_text",
            "test_image_to_markdown_ocr_success_warns_and_preserves_recognized_text",
            "test_tiff_to_markdown_emits_one_fragment_per_frame_and_continues_after_ocr_failure",
        ),
        "packages/plugins/markup/tests/test_markdown_resources.py": (
            "test_writer_reports_typed_ocr_failure_and_continues_with_later_images",
            "test_writer_warns_that_no_text_may_be_a_missed_best_effort_result",
        ),
        "packages/plugins/layout/tests/test_layout_conversions_*.py": (
            "test_direct_page_ocr_helper_preserves_typed_model_failure",
            "test_page_level_ocr_render_failure_preserves_base_conversion",
            "test_page_level_no_text_warns_about_possible_missed_text",
        ),
        "packages/plugins/layout/tests/test_markdown_preprocess_*.py": ("test_ocr_failure_preserves_local_image_link",),
        "packages/plugins/document/tests/test_request_scoped_docx_policy_*.py": (
            "test_document_all_ocr_outcomes_warn_and_continue_later_images",
        ),
        "packages/plugins/spreadsheet/tests/test_xlsx_to_md_golden_*.py": (
            "test_pipeline_xlsx_all_ocr_outcomes_warn_and_continue_later_images",
        ),
        "packages/plugins/presentation/tests/test_presentation_to_md_*.py": (
            "test_pptx_all_ocr_outcomes_warn_and_continue_later_images",
        ),
        "packages/plugins/optimizers/gongwen/tests/test_gongwen_golden.py": (
            "test_gongwen_all_ocr_outcomes_warn_and_continue_later_images",
        ),
        "packages/plugins/optimizers/invoice_cn/tests/test_invoice_conversions_*.py": (
            "test_direct_image_parser_preserves_typed_ocr_failure",
            "test_image_invoice_typed_ocr_failures_are_safe_and_nonfatal",
            "test_image_invoice_no_text_warns_without_reporting_false_success",
        ),
        "packages/plugins/optimizers/invoice_cn/tests/test_invoice_scan_pages.py": (
            "test_scan_pdf_ocr_failure_is_best_effort",
            "test_scan_pdf_typed_ocr_failure_falls_back_to_text_parser",
        ),
        "packages/apps/cli/tests/test_cli_runtime_export_defaults.py": (
            "test_cli_image_without_ocr_flag_disables_ocr_even_when_config_default_is_enabled",
            "test_cli_image_ocr_uses_export_ocr_placement_default_without_flag",
            "test_cli_image_ocr_uses_configured_ocr_language_default",
        ),
        "packages/apps/gui/tests/test_gui_e2e_conversion_*.py": (
            "test_mhtml_to_markdown_ocr_runs_through_gui_action_only_route",
        ),
    }
    for path, test_names in expected_tests.items():
        source = _read(path)
        for test_name in test_names:
            assert f"def {test_name}(" in source, (path, test_name)

    cli_source = _read("packages/apps/cli/tests/test_cli_runtime_export_defaults.py")
    assert "run_ocr_outcome" in cli_source
    assert "OcrOutcome" in cli_source and "OcrStatus.SUCCESS" in cli_source
    assert "ocr_extract_text" not in cli_source

    gui_source = _read("packages/apps/gui/tests/test_gui_e2e_conversion_*.py")
    assert '"docwen_plugin_markup.markdown_resources.run_ocr_outcome"' in gui_source
    assert "OcrOutcome" in gui_source and "OcrStatus.SUCCESS" in gui_source
