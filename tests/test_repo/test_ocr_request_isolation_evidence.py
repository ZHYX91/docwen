"""Fail-closed evidence guards for VIS-2026-07-20-147 OCR isolation."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "rapidocr-engine-invocation-isolation-and-request-availability-2026-07-20.md"
OCR_PATH = PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "text" / "ocr.py"
OCR_TEST_PATH = PROJECT_ROOT / "packages" / "core" / "tests" / "test_ocr_request_isolation.py"
ARCHITECTURE_TEST_PATH = PROJECT_ROOT / "tests" / "test_repo" / "test_ocr_architecture.py"


def _read(relative_path: str | Path) -> str:
    path = relative_path if isinstance(relative_path, Path) else PROJECT_ROOT / relative_path
    return path.read_text(encoding="utf-8")


def _definition_source(source: str, name: str) -> tuple[ast.FunctionDef | ast.AsyncFunctionDef, str]:
    tree = ast.parse(source)
    node = next(
        candidate
        for candidate in tree.body
        if isinstance(candidate, (ast.FunctionDef, ast.AsyncFunctionDef)) and candidate.name == name
    )
    segment = ast.get_source_segment(source, node)
    assert segment is not None
    return node, segment


def test_core_owns_one_invocation_lock_per_cached_engine_slot() -> None:
    source = _read(OCR_PATH)
    tree = ast.parse(source)
    slot = next(node for node in tree.body if isinstance(node, ast.ClassDef) and node.name == "_OcrEngineSlot")
    outcome_node, outcome_source = _definition_source(source, "run_ocr_outcome")

    assert any(
        isinstance(node, ast.Assign)
        and any(isinstance(target, ast.Attribute) and target.attr == "invocation_lock" for target in node.targets)
        and isinstance(node.value, ast.Call)
        and isinstance(node.value.func, ast.Attribute)
        and node.value.func.attr == "Lock"
        for node in ast.walk(slot)
    )
    assert "dict[tuple[str, str], _OcrEngineSlot]" in source
    assert "slot = _get_ocr_slot(" in outcome_source
    assert "with slot.invocation_lock:" in outcome_source
    assert "result, _elapsed = engine(str(image_file))" in outcome_source
    assert any(isinstance(node, ast.With) for node in ast.walk(outcome_node))
    assert "with _ocr_lock:\n" not in outcome_source
    function_names = {node.name for node in tree.body if isinstance(node, ast.FunctionDef)}
    assert {"run_ocr", "extract_text_from_image", "extract_text_from_image_outcome"}.isdisjoint(function_names)


def test_availability_and_reset_share_the_request_and_cache_ordering_contract() -> None:
    source = _read(OCR_PATH)
    available_node, available_source = _definition_source(source, "ocr_available")
    _reset_node, reset_source = _definition_source(source, "reset_ocr")

    positional = [argument.arg for argument in available_node.args.args]
    keyword_only = [argument.arg for argument in available_node.args.kwonlyargs]
    assert positional == ["ocr_language"]
    assert keyword_only == ["current_locale", "model_dir"]
    assert "_get_ocr(" in available_source
    assert "ocr_language," in available_source
    assert "current_locale=current_locale" in available_source
    assert "model_dir=model_dir" in available_source
    assert "except FileNotFoundError:" in available_source
    assert "with _ocr_lock:" in reset_source
    assert "_ocr_instances.clear()" in reset_source


def test_direct_regressions_cover_request_selection_parallelism_and_reset() -> None:
    tests = _read(OCR_TEST_PATH)
    required_tests = (
        "test_ocr_available_uses_requested_language_and_locale",
        "test_ocr_available_rejects_missing_requested_models_even_when_default_exists",
        "test_run_ocr_serializes_invocations_of_the_same_cached_engine",
        "test_run_ocr_allows_different_cached_engines_to_run_in_parallel",
        "test_reset_ocr_waits_for_in_flight_initialization_before_clearing_cache",
    )
    for test_name in required_tests:
        assert f"def {test_name}(" in tests

    for token in (
        "initialization_count == 1",
        "max_active_calls == 1",
        "max_active_calls == 2",
        "second_slot_obtained",
        "second_entered_before_release",
        "threading.Barrier(2)",
        "reset_completed_during_initialization",
        "initialization_count == 2",
        "japanese-only",
        "chinese-only",
        "japan_PP-OCRv4_rec_infer.onnx",
    ):
        assert token in tests


def test_repository_architecture_guard_requires_exact_request_locals() -> None:
    guard = _read(ARCHITECTURE_TEST_PATH)
    plugin_root = PROJECT_ROOT / "packages" / "plugins"
    callers: set[str] = set()
    offenders: list[str] = []
    for path in plugin_root.rglob("*.py"):
        if "src" not in path.relative_to(plugin_root).parts:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            func = node.func
            typed_names = {"run_ocr_outcome"}
            is_typed_call = (isinstance(func, ast.Name) and func.id in typed_names) or (
                isinstance(func, ast.Attribute) and func.attr in typed_names
            )
            if not is_typed_call:
                continue
            relative_path = path.relative_to(PROJECT_ROOT).as_posix()
            callers.add(relative_path)
            keyword_values = {keyword.arg: keyword.value for keyword in node.keywords if keyword.arg}
            for name in ("ocr_language", "current_locale"):
                value = keyword_values.get(name)
                if not isinstance(value, ast.Name) or value.id != name:
                    offenders.append(f"{relative_path}:{node.lineno}:{name}")

    assert callers == {
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
    assert offenders == []
    for token in (
        "test_plugins_do_not_preflight_typed_ocr_with_ocr_available",
        "test_plugin_typed_ocr_calls_pass_language_and_locale",
        'typed_call_names = {"run_ocr_outcome"}',
        'for name in ("ocr_language", "current_locale")',
        "keyword_values = {keyword.arg: keyword.value",
        "or keyword_values[name].id != name",
        "assert offenders == []",
    ):
        assert token in guard
