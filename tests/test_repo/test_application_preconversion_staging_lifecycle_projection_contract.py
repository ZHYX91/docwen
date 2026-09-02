from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
CONTROLLER = ROOT / "packages/application/src/docwen_application/controller.py"
PRECONVERTER = ROOT / "packages/application/src/docwen_application/preconversion/pre_converter.py"
CONTROLLER_TESTS = ROOT / "packages/application/tests/test_controller_*.py"
RUNTIME_TESTS = ROOT / "packages/runtime/tests/test_fake_closed_loop_*.py"
REPORT_NAME = "application-preconversion-staging-lifecycle-output-ownership-2026-07-21.md"


def test_preconversion_staging_has_owner_isolation_and_original_index_alignment() -> None:
    controller = CONTROLLER.read_text(encoding="utf-8")

    assert "tempfile.TemporaryDirectory(" in controller
    assert 'prefix="docwen_pre_"' in controller
    assert "staging_dir = Path(temp_owner.name) / str(idx)" in controller
    assert "tempfile.mkdtemp(" not in controller
    assert "input_indices: list[int]" in controller
    runnable_ref_appends = controller.count("new_refs.append(")
    assert runnable_ref_appends > 0
    assert controller.count("input_indices.append(idx)") == runnable_ref_appends
    assert "plan.input_indices," in controller
    assert "strict=True" in controller
    assert 'task_id = f"{plan.request.request_id}-{original_index}"' in controller
    assert "request_id=task_id" in controller
    assert controller.count("managed.cleanup()") == 2
    assert controller.index("if errors and not batch:") < controller.index(
        "converted_options = deepcopy(request.options)"
    )


def test_preconversion_preserves_source_policy_and_file_identity() -> None:
    controller = CONTROLLER.read_text(encoding="utf-8")
    preconverter = PRECONVERTER.read_text(encoding="utf-8")

    source_policy = controller.split("def _source_anchored_output_policy", 1)[1].split("def _configured_priority", 1)[0]
    assert "if output_policy.output_path or output_policy.output_dir:" in source_policy
    assert "return output_policy" in source_policy
    assert "return replace(output_policy, output_dir=str(Path(source_path).parent))" in source_policy
    for token in (
        "category=ref.category",
        "encoding=ref.encoding",
        "derived_metadata = deepcopy(ref.metadata)",
        "derived_metadata.pop(FILE_ADMISSION_ACCEPTANCE_METADATA_KEY, None)",
        '"warning_message": ref.warning_message',
        'warning_message=""',
        "metadata=derived_metadata",
        "size_bytes=Path(pre_result.pre_converted_path).stat().st_size",
    ):
        assert token in controller
    assert 'output_path = str(stage_path / f"{stem}.{hub_format}")' in preconverter
    assert 'f"{stem}_pre.{hub_format}"' not in preconverter


def test_preconversion_regressions_cover_collision_cleanup_policy_and_finalization() -> None:
    controller_tests = read_source_text(CONTROLLER_TESTS)
    runtime_tests = read_source_text(RUNTIME_TESTS)

    for token in (
        "test_preconversion_batch_isolates_same_stem_preserves_output_policy_and_cleans_staging",
        "test_preconversion_staging_is_cleaned_when_runtime_raises",
        "test_preconversion_staging_is_cleaned_when_backend_is_unavailable",
        "test_preconversion_staging_is_cleaned_when_backend_raises",
        "test_preconversion_staging_is_cleaned_when_request_rebuild_raises",
        'assert runtime_request.request_id == "r1-1"',
        '"source-none", "source-empty", "custom"',
        "failed preconversion must not rebuild options",
    ):
        assert token in controller_tests
    assert "test_preconversion_same_stem_batch_finalizes_to_each_original_parent" in runtime_tests
    assert "[path.parent.parent for path in final_paths] == [first_dir, second_dir]" in runtime_tests
    assert "path.parent.name == path.stem" in runtime_tests
    assert 'path.parent / "docwen-node.json"' in runtime_tests
    assert 'not list(tmp_path.glob("docwen_pre_*"))' in runtime_tests
