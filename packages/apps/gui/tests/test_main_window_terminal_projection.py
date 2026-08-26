"""Regressions for single-owner execution outcome projection."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

pytestmark = pytest.mark.gui


def _add_processing_entry(main_window, source: Path, operation_id: str) -> str:
    normalized = str(source).replace("\\", "/")
    added, rejected = main_window._batch_list_vm.add_files([str(source)])
    assert added == [normalized]
    assert rejected == []
    assert main_window._batch_list_vm.set_file_status(
        normalized,
        "processing",
        operation_id=operation_id,
    )
    return normalized


def _flush_runtime_event(main_window, event_type: str, payload: dict[str, object]) -> None:
    bridge = main_window.task_event_bridge
    assert bridge is not None
    bridge.enqueue(event_type, payload)
    assert bridge.flush() == 1


def _begin_single_telemetry(main_window, task_id: str) -> None:
    main_window.view_model.begin_execution_telemetry(task_id, (task_id,))


def _context(task_id: str, file_paths: list[str], *, batch: bool = False, aggregate: bool = False) -> dict:
    return {
        "request_id": task_id,
        "file_path": file_paths[0],
        "file_paths": file_paths,
        "display_name": "Test execution",
        "total_count": len(file_paths),
        "batch": batch,
        "aggregate": aggregate,
    }


@pytest.mark.parametrize("event_type", ["task_completed", "task_failed", "task_cancelled"])
def test_runtime_terminal_event_is_telemetry_not_operation_terminal(
    main_window,
    tmp_path: Path,
    event_type: str,
) -> None:
    source = tmp_path / f"{event_type}.md"
    source.write_text("# Input\n", encoding="utf-8")
    task_id = f"task-{event_type}"
    normalized = _add_processing_entry(main_window, source, task_id)
    _begin_single_telemetry(main_window, task_id)
    summaries: list[dict] = []
    main_window.view_model.task_summary_changed.connect(summaries.append)

    _flush_runtime_event(
        main_window,
        "task_started",
        {"task_id": task_id, "input_path": normalized, "message": "working"},
    )
    live_status = main_window.view_model.status_message
    _flush_runtime_event(main_window, event_type, {"task_id": task_id, "message": "terminal"})

    entry = main_window._batch_list_vm.get_file_entry(normalized)
    assert entry is not None
    assert entry.status == "processing"
    assert not entry.output_path
    assert main_window.view_model.status_message == live_status
    assert summaries == []


def test_execution_result_publishes_completed_with_output_path_atomically(
    main_window,
    tmp_path: Path,
) -> None:
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionResult

    source = tmp_path / "input.md"
    source.write_text("# Input\n", encoding="utf-8")
    output = tmp_path / "input.docx"
    output.write_bytes(b"result")
    task_id = "task-success"
    normalized = _add_processing_entry(main_window, source, task_id)
    _begin_single_telemetry(main_window, task_id)

    # Runtime completion may reach the GUI before the QThread result signal.
    _flush_runtime_event(main_window, "task_completed", {"task_id": task_id})
    before_result = main_window._batch_list_vm.get_file_entry(normalized)
    assert before_result is not None
    assert before_result.status == "processing"

    observed: list[tuple[str, str]] = []
    summaries: list[dict] = []

    def capture_terminal(file_path: str, status: str) -> None:
        entry = main_window._batch_list_vm.get_file_entry(file_path)
        assert entry is not None
        observed.append((status, entry.output_path))

    main_window._batch_list_vm.status_changed.connect(capture_terminal)
    main_window.view_model.task_summary_changed.connect(summaries.append)
    main_window._on_execution_finished(
        ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id="primary",
                    kind="primary",
                    staging_path=str(output),
                    suggested_name=output.name,
                    is_primary=True,
                )
            ],
        ),
        _context(task_id, [normalized]),
    )

    entry = main_window._batch_list_vm.get_file_entry(normalized)
    assert entry is not None
    assert entry.status == "completed"
    assert entry.output_path == str(output)
    assert observed == [("completed", str(output))]
    assert [summary["status"] for summary in summaries] == ["completed"]


@pytest.mark.parametrize(
    ("runtime_event", "error_type", "expected_status"),
    [
        ("task_failed", "conversion_failed", "failed"),
        ("task_cancelled", "cancelled", "cancelled"),
    ],
)
def test_execution_result_publishes_failure_payload_atomically(
    main_window,
    tmp_path: Path,
    runtime_event: str,
    error_type: str,
    expected_status: str,
) -> None:
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    source = tmp_path / f"{expected_status}.md"
    source.write_text("# Input\n", encoding="utf-8")
    retained = tmp_path / f"{expected_status}.docx"
    retained.write_bytes(b"retained")
    task_id = f"task-{expected_status}"
    normalized = _add_processing_entry(main_window, source, task_id)
    _begin_single_telemetry(main_window, task_id)
    main_window._batch_list_vm.set_file_status(normalized, "processing", output_path="stale-output")

    _flush_runtime_event(main_window, runtime_event, {"task_id": task_id, "message": "boom"})

    observed: list[tuple[str, str, str]] = []
    summaries: list[dict] = []

    def capture_terminal(file_path: str, status: str) -> None:
        entry = main_window._batch_list_vm.get_file_entry(file_path)
        assert entry is not None
        observed.append((status, entry.output_path, entry.error_message))

    main_window._batch_list_vm.status_changed.connect(capture_terminal)
    main_window.view_model.task_summary_changed.connect(summaries.append)
    main_window._on_execution_finished(
        ConversionResult(
            task_id=task_id,
            success=False,
            artifacts=[
                ArtifactManifest(
                    artifact_id="retained",
                    kind="intermediate",
                    staging_path=str(retained),
                    suggested_name=retained.name,
                )
            ],
            error=ConversionErrorInfo(error_type=error_type, message="boom"),
        ),
        _context(task_id, [normalized]),
    )

    expected_output = "" if expected_status == "cancelled" else str(retained)
    assert observed == [(expected_status, expected_output, "boom")]
    assert [summary["status"] for summary in summaries] == [expected_status]


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_failed_result_keeps_long_retained_output_reachable(main_window, tmp_path: Path) -> None:
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult
    from docwen_runtime.path_io import filesystem_path

    remaining = 260 - len(os.path.abspath(tmp_path)) - 1
    if remaining < 1 or remaining > 255:
        pytest.skip("pytest temp root cannot express an exact 260-character file")
    retained = tmp_path / ("r" * remaining)
    filesystem_path(retained).write_bytes(b"retained")
    result = ConversionResult(
        task_id="long-retained",
        success=False,
        artifacts=[
            ArtifactManifest(
                artifact_id="retained",
                kind="intermediate",
                staging_path=str(retained),
                suggested_name=retained.name,
            )
        ],
        error=ConversionErrorInfo(error_type="conversion_failed", message="boom"),
    )

    assert main_window._pick_existing_output_path(result) == str(retained)


@pytest.mark.skipif(os.name != "nt", reason="Win32 namespace policy")
def test_failed_result_skips_invalid_namespace_artifacts(main_window, tmp_path: Path) -> None:
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    retained = tmp_path / "retained.docx"
    retained.write_bytes(b"retained")

    def artifact(artifact_id: str, path: str, *, primary: bool) -> ArtifactManifest:
        return ArtifactManifest(
            artifact_id=artifact_id,
            kind="intermediate",
            staging_path=path,
            suggested_name=Path(path).name,
            is_primary=primary,
        )

    invalid = artifact(
        "invalid",
        r"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\report.docx",
        primary=True,
    )
    valid = artifact("retained", str(retained), primary=False)
    error = ConversionErrorInfo(error_type="conversion_failed", message="boom")

    with_fallback = ConversionResult(
        task_id="invalid-then-valid",
        success=False,
        artifacts=[invalid, valid],
        error=error,
    )
    invalid_only = ConversionResult(
        task_id="invalid-only",
        success=False,
        artifacts=[invalid],
        error=error,
    )

    assert main_window._pick_existing_output_path(with_fallback) == str(retained)
    assert main_window._pick_existing_output_path(invalid_only) == ""


@pytest.mark.parametrize("child_terminal", ["task_completed", "task_failed", "task_cancelled"])
def test_batch_child_terminal_does_not_publish_whole_operation_terminal(
    main_window,
    tmp_path: Path,
    child_terminal: str,
) -> None:
    first = tmp_path / "first.md"
    second = tmp_path / "second.md"
    first.write_text("# First\n", encoding="utf-8")
    second.write_text("# Second\n", encoding="utf-8")
    parent_id = "batch-parent"
    first_normalized = _add_processing_entry(main_window, first, parent_id)
    second_normalized = _add_processing_entry(main_window, second, parent_id)
    main_window.view_model.begin_execution_telemetry(
        parent_id,
        (f"{parent_id}-0", f"{parent_id}-1"),
    )
    summaries: list[dict] = []
    main_window.view_model.task_summary_changed.connect(summaries.append)

    _flush_runtime_event(
        main_window,
        "task_started",
        {"task_id": f"{parent_id}-0", "input_path": first_normalized, "message": "first"},
    )
    live_status = main_window.view_model.status_message
    _flush_runtime_event(
        main_window,
        child_terminal,
        {"task_id": f"{parent_id}-0", "message": "child terminal"},
    )

    assert main_window._batch_list_vm.get_file_entry(first_normalized).status == "processing"
    assert main_window._batch_list_vm.get_file_entry(second_normalized).status == "processing"
    assert main_window.view_model.status_message == live_status
    assert summaries == []


def test_batch_result_publishes_one_aligned_operation_summary(
    main_window,
    tmp_path: Path,
) -> None:
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    first = tmp_path / "batch-first.md"
    second = tmp_path / "batch-second.md"
    first.write_text("# First\n", encoding="utf-8")
    second.write_text("# Second\n", encoding="utf-8")
    output = tmp_path / "batch-first.docx"
    output.write_bytes(b"result")
    parent_id = "batch-result"
    paths = [
        _add_processing_entry(main_window, first, parent_id),
        _add_processing_entry(main_window, second, parent_id),
    ]
    main_window.view_model.begin_execution_telemetry(
        parent_id,
        (f"{parent_id}-0", f"{parent_id}-1"),
    )
    summaries: list[dict] = []
    main_window.view_model.task_summary_changed.connect(summaries.append)

    _flush_runtime_event(
        main_window,
        "task_completed",
        {"task_id": f"{parent_id}-0", "message": "child done"},
    )
    assert summaries == []

    main_window._on_execution_finished(
        [
            ConversionResult(
                task_id=f"{parent_id}-0",
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="first",
                        kind="primary",
                        staging_path=str(output),
                        suggested_name=output.name,
                        is_primary=True,
                    )
                ],
            ),
            ConversionResult(
                task_id=f"{parent_id}-1",
                success=False,
                error=ConversionErrorInfo(error_type="conversion_failed", message="second failed"),
            ),
        ],
        _context(parent_id, paths, batch=True),
    )

    first_entry = main_window._batch_list_vm.get_file_entry(paths[0])
    second_entry = main_window._batch_list_vm.get_file_entry(paths[1])
    assert first_entry is not None
    assert second_entry is not None
    assert (first_entry.status, first_entry.output_path) == ("completed", str(output))
    assert (second_entry.status, second_entry.error_message) == ("failed", "second failed")
    assert len(summaries) == 1
    assert summaries[0] == {
        "task_id": parent_id,
        "state": "partial",
        "completed_count": 2,
        "total_count": 2,
        "failed_count": 1,
        "skipped_count": 0,
        "cancelled_count": 0,
        "message": "second failed",
        "status": "partial",
    }


@pytest.mark.parametrize(
    ("runtime_event", "error_type", "expected_status"),
    [
        ("task_completed", "", "completed"),
        ("task_failed", "conversion_failed", "failed"),
        ("task_cancelled", "cancelled", "cancelled"),
    ],
)
def test_aggregate_result_is_single_owner_for_every_input_row(
    main_window,
    tmp_path: Path,
    runtime_event: str,
    error_type: str,
    expected_status: str,
) -> None:
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult

    first = tmp_path / "first.pdf"
    second = tmp_path / "second.pdf"
    first.write_bytes(b"%PDF-first")
    second.write_bytes(b"%PDF-second")
    output = tmp_path / "merged.pdf"
    output.write_bytes(b"%PDF-merged")
    task_id = f"aggregate-{expected_status}"
    paths = [
        _add_processing_entry(main_window, first, task_id),
        _add_processing_entry(main_window, second, task_id),
    ]
    _begin_single_telemetry(main_window, task_id)
    summaries: list[dict] = []
    main_window.view_model.task_summary_changed.connect(summaries.append)

    _flush_runtime_event(
        main_window,
        "task_started",
        {"task_id": task_id, "input_path": paths[0], "message": "merge"},
    )
    _flush_runtime_event(main_window, runtime_event, {"task_id": task_id, "message": "terminal"})
    assert [main_window._batch_list_vm.get_file_entry(path).status for path in paths] == [
        "processing",
        "processing",
    ]
    assert summaries == []

    success = expected_status == "completed"
    main_window._on_execution_finished(
        ConversionResult(
            task_id=task_id,
            success=success,
            artifacts=(
                [
                    ArtifactManifest(
                        artifact_id="merged",
                        kind="primary",
                        staging_path=str(output),
                        suggested_name=output.name,
                        is_primary=True,
                    )
                ]
                if success
                else []
            ),
            error=None if success else ConversionErrorInfo(error_type=error_type, message="terminal"),
        ),
        _context(task_id, paths, aggregate=True),
    )

    assert [main_window._batch_list_vm.get_file_entry(path).status for path in paths] == [
        expected_status,
        expected_status,
    ]
    expected_output = str(output) if success else ""
    assert [main_window._batch_list_vm.get_file_entry(path).output_path for path in paths] == [
        expected_output,
        expected_output,
    ]
    assert [summary["status"] for summary in summaries] == [expected_status]


def test_late_started_from_previous_operation_cannot_overwrite_new_operation(
    main_window,
    tmp_path: Path,
) -> None:
    source = tmp_path / "same.md"
    source.write_text("# Input\n", encoding="utf-8")
    normalized = _add_processing_entry(main_window, source, "operation-a")
    _begin_single_telemetry(main_window, "operation-a")
    main_window._batch_list_vm.set_file_status(
        normalized,
        "completed",
        output_path="operation-a.docx",
        operation_id="operation-a",
    )
    main_window._batch_list_vm.set_file_status(
        normalized,
        "processing",
        output_path="",
        operation_id="operation-b",
    )
    _begin_single_telemetry(main_window, "operation-b")
    _flush_runtime_event(
        main_window,
        "task_started",
        {"task_id": "operation-b", "input_path": normalized, "message": "current"},
    )

    _flush_runtime_event(
        main_window,
        "task_started",
        {"task_id": "operation-a", "input_path": normalized, "message": "stale"},
    )

    entry = main_window._batch_list_vm.get_file_entry(normalized)
    assert entry is not None
    assert entry.status == "processing"
    assert entry.operation_id == "operation-b"
    assert not entry.output_path
    assert main_window.view_model.current_task_id == "operation-b"
