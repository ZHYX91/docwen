from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
CONTROLLER = ROOT / "packages/application/src/docwen_application/controller.py"
PRECONVERTER = ROOT / "packages/application/src/docwen_application/preconversion/pre_converter.py"
CONTROLLER_TESTS = ROOT / "packages/application/tests/test_controller_*.py"
PRECONVERTER_TESTS = ROOT / "packages/application/tests/test_pre_converter.py"
RUNTIME_TESTS = ROOT / "packages/runtime/tests/test_fake_closed_loop_*.py"
REPORT_NAME = "application-preconversion-protective-source-copy-2026-07-21.md"


def test_preconverter_isolates_the_external_bridge_behind_a_canonical_copy() -> None:
    preconverter = PRECONVERTER.read_text(encoding="utf-8")
    controller = CONTROLLER.read_text(encoding="utf-8")

    for token in (
        '"doc": ".doc"',
        '"wps": ".doc"',
        '"rtf": ".rtf"',
        '"odt": ".odt"',
        "source_suffix = _PROTECTIVE_SOURCE_SUFFIXES.get(source_format)",
        'protected_input = stage_path / f"input{source_suffix}"',
        "_prepare_protective_snapshot(source_path, protected_input, cancel)",
        "str(protected_input)",
        'diagnostic_code="PRECONVERSION_INPUT_COPY_FAILED"',
        "error_type=_bridge_failure_error_type(result)",
        "cleanup_failed=result.cleanup_failed",
    ):
        assert token in preconverter

    assert preconverter.count("_is_cancel_requested(cancel)") >= 3
    bridge_call = preconverter.split("result = convert_with_backend_priority(", 1)[1]
    assert bridge_call.lstrip().startswith("str(protected_input),")
    assert 'f"{stem}.{hub_format}"' in preconverter
    assert "error_type=pre_result.error_type" in controller
    assert "diagnostic_code=pre_result.diagnostic_code" in controller


def test_protective_copy_regressions_cover_mapping_cancel_failure_sniff_and_lifecycle() -> None:
    controller_tests = read_source_text(CONTROLLER_TESTS)
    preconverter_tests = read_source_text(PRECONVERTER_TESTS)
    runtime_tests = read_source_text(RUNTIME_TESTS)

    for token in (
        "test_pre_convert_bridges_a_content_typed_protective_copy_and_preserves_source_stem",
        "test_pre_convert_cancelled_before_copy_skips_copy_and_bridge",
        "test_pre_convert_cancelled_after_copy_skips_bridge",
        "test_pre_convert_copy_failure_is_structured_and_skips_bridge",
        "test_pre_convert_concurrent_cancel_wins_over_copy_failure",
        "test_pre_convert_rejects_hub_source_without_copy_or_bridge",
        "test_pre_convert_preserves_bridge_cancellation_outcome",
    ):
        assert token in preconverter_tests

    for token in (
        "test_preconversion_copy_failure_is_structured_and_cleans_staging",
        "test_cancel_after_protective_copy_skips_bridge_and_cleans_staging",
        "test_preconversion_copy_failure_stays_aligned_in_batch",
        "test_admitted_rtf_with_disguised_suffix_is_bridged_from_canonical_copy",
        'assert result.error.diagnostic_code == "PRECONVERSION_INPUT_COPY_FAILED"',
        '[("input.rtf", source_bytes)]',
    ):
        assert token in controller_tests

    assert "test_preconversion_same_stem_batch_finalizes_to_each_original_parent" in runtime_tests
    assert '["input.rtf", "input.rtf"]' in runtime_tests
    assert "all(not path.exists() for path, _content in protected_inputs)" in runtime_tests
