"""Fail-closed evidence guards for VIS-2026-07-20-148."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "runtime-port-ocr-language-locale-snapshot-fallback-2026-07-20.md"
ADAPTER_PATH = PROJECT_ROOT / "packages" / "runtime" / "src" / "docwen_runtime" / "adapters.py"
ADMISSION_PATH = PROJECT_ROOT / "packages" / "runtime" / "src" / "docwen_runtime" / "_request_admission.py"
TASK_MANAGER_PATH = PROJECT_ROOT / "packages" / "runtime" / "src" / "docwen_runtime" / "engine" / "task_manager.py"
DIRECT_TEST_PATH = PROJECT_ROOT / "packages" / "runtime" / "tests" / "test_runtime_ocr_option_projection.py"
DIRECT_MANAGER_TEST_PATH = (
    PROJECT_ROOT / "packages" / "runtime" / "tests" / "test_task_manager_ocr_option_projection.py"
)


def _read(relative_path: str | Path) -> str:
    path = relative_path if isinstance(relative_path, Path) else PROJECT_ROOT / relative_path
    return path.read_text(encoding="utf-8")


def _definition_source(source: str, name: str) -> str:
    tree = ast.parse(source)
    node = next(
        candidate
        for candidate in ast.walk(tree)
        if isinstance(candidate, (ast.FunctionDef, ast.AsyncFunctionDef)) and candidate.name == name
    )
    segment = ast.get_source_segment(source, node)
    assert segment is not None
    return segment


def test_shared_admission_projects_missing_ocr_options_from_one_snapshot() -> None:
    source = _read(ADMISSION_PATH)
    adapter_source = _read(ADAPTER_PATH)
    manager_source = _read(TASK_MANAGER_PATH)
    projection = _definition_source(source, "project_markdown_ocr_options")
    admit = _definition_source(source, "admit_markdown_ocr_options")
    execute = _definition_source(adapter_source, "execute")
    execute_single = _definition_source(manager_source, "execute_single")

    for token in (
        'request.target_format != "md"',
        'request.action_name == "process_md_numbering"',
        'if "ocr_language" not in options:',
        '_nested_value(config_snapshot, "image", "ocr_language")',
        'options["ocr_language"] = ocr_language or "auto"',
        'if "locale" not in options:',
        '_nested_value(config_snapshot, "gui", "language", "locale")',
        'options["locale"] = locale or "zh_CN"',
    ):
        assert token in source
    assert "request.options[" not in projection
    assert "ConfigLoader" not in source
    assert "replace(" in admit
    assert "config_snapshot = req.config_snapshot" in execute
    assert execute.count("self._config_loader.config.as_dict()") == 1
    assert "req = admit_markdown_ocr_options(req, config_snapshot)" in execute
    assert "admit_markdown_ocr_options(request, request.config_snapshot)" in execute_single
    assert "config_loader" not in execute_single.casefold()


def test_direct_regressions_cover_authority_precedence_and_scope() -> None:
    tests = _read(DIRECT_TEST_PATH)
    direct_manager_tests = _read(DIRECT_MANAGER_TEST_PATH)
    required_tests = (
        "test_empty_options_use_the_same_loader_snapshot_as_runtime_context",
        "test_real_config_loader_projects_reloaded_ocr_values_without_mutating_request",
        "test_request_snapshot_is_authoritative_over_the_live_loader",
        "test_partial_request_snapshot_uses_protocol_defaults_not_live_loader",
        "test_explicit_nonblank_option_wins_while_missing_peer_uses_snapshot",
        "test_present_falsey_options_are_preserved_by_projection",
        "test_explicit_nonblank_ocr_options_are_not_overwritten",
        "test_noncanonical_markdown_targets_do_not_receive_ocr_options",
        "test_markdown_numbering_action_does_not_receive_unconsumed_ocr_options",
        "test_request_without_any_snapshot_keeps_empty_options",
    )
    for test_name in required_tests:
        assert f"def {test_name}(" in tests
    for token in (
        '"image.ocr_language": "japanese"',
        '"gui.language.locale": "ja_JP"',
        "assert request.options == {}",
        "assert request.config_snapshot == {}",
        '"to_md_enable_ocr": False',
        '["docx", "markdown", "MD", " md "]',
        '{"ocr_language": "auto", "locale": "ja_JP"}',
        '{"ocr_language": "japanese", "locale": "zh_CN"}',
    ):
        assert token in tests

    for test_name in (
        "test_direct_single_projects_snapshot_values_without_mutating_caller",
        "test_direct_single_preserves_present_falsey_keys",
        "test_direct_single_keeps_projection_scope",
        "test_direct_batch_projects_each_derived_request",
    ):
        assert f"def {test_name}(" in direct_manager_tests
    for token in (
        '("partial", "md", {"image": {"ocr_language": "english"}}',
        '("empty", "md", {}, {})',
        '"non-markdown",',
        '"docx",',
        "assert request.options == {}",
    ):
        assert token in direct_manager_tests
