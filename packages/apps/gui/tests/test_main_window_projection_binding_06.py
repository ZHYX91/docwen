"""Focused tests split from test_main_window_projection_binding.py."""

from __future__ import annotations

from ._main_window_projection_binding_support import (
    Path,
    SimpleNamespace,
    optimization_capability_projection,
    pytest,
)
from ._main_window_projection_binding_support import (
    left_frame as left_frame,
)

pytestmark = pytest.mark.gui
from ._main_window_projection_binding_support import (
    right_frame as right_frame,
)
from ._main_window_projection_binding_support import (
    right_stack as right_stack,
)
from ._main_window_projection_binding_support import (
    window as window,
)


class TestRuntimeRequestBinding:
    def test_batch_failed_result_with_retained_auxiliary_exposes_output_without_changing_counts(
        self,
        window,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        source = tmp_path / "legacy.doc"
        retained = tmp_path / "out" / "legacy_fromDoc.docx"
        source.write_text("legacy", encoding="utf-8")
        retained.parent.mkdir()
        retained.write_text("preserved", encoding="utf-8")
        normalized = str(source).replace("\\", "/")
        window._batch_list_vm.add_files([str(source)])
        opened: list[tuple[str, bool]] = []
        monkeypatch.setattr(
            window,
            "_open_path",
            lambda target, *, open_parent=False: opened.append((target, open_parent)) or True,
        )

        result = ConversionResult(
            task_id="batch-retained-0",
            success=False,
            artifacts=[
                ArtifactManifest(
                    artifact_id="retained-hub",
                    kind="auxiliary",
                    staging_path=str(retained),
                    suggested_name=retained.name,
                )
            ],
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message="downstream conversion failed",
            ),
        )
        window._on_execution_finished(
            [result],
            {
                "request_id": "batch-retained",
                "file_path": normalized,
                "file_paths": [normalized],
                "display_name": "Batch conversion (1 file)",
                "total_count": 1,
                "batch": True,
                "open_after_done": True,
            },
        )

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "failed"
        assert entry.error_message == "downstream conversion failed"
        assert entry.output_path == str(retained)
        summary = window._info_area_vm.task_summary
        assert summary.state == "failed"
        assert summary.completed_count == 1
        assert summary.failed_count == 1
        assert summary.navigate_path == normalized
        assert summary.navigation_kind == "failed"
        assert window._info_area_vm.guide_actions == [
            {"action_key": "open_output_dir", "target_path": str(retained.parent)},
            {"action_key": "view_failed_details", "target_path": normalized},
            {"action_key": "retry_failed", "target_path": ""},
            {"action_key": "add_more_files", "target_path": ""},
        ]
        history = window._info_area_vm.history_rows
        assert history[-1].message_type == "danger"
        assert history[-1].file_path == str(retained)
        assert history[-1].navigate_file_path == str(retained)
        assert opened == []

    def test_batch_multiple_retained_failures_keep_each_entry_output_reachable(
        self,
        window,
        tmp_path,
    ) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult
        from docwen_gui.widgets.batch_list import BatchEntryItemWidget

        sources = [tmp_path / "first.doc", tmp_path / "second.doc"]
        retained = [tmp_path / "out-a" / "first.docx", tmp_path / "out-b" / "second.docx"]
        for source in sources:
            source.write_text("legacy", encoding="utf-8")
        for output in retained:
            output.parent.mkdir()
            output.write_text("preserved", encoding="utf-8")
        normalized = [str(source).replace("\\", "/") for source in sources]
        window._batch_list_vm.add_files([str(source) for source in sources])

        results = [
            ConversionResult(
                task_id=f"multi-retained-{index}",
                success=False,
                artifacts=[
                    ArtifactManifest(
                        artifact_id=f"retained-{index}",
                        kind="auxiliary",
                        staging_path=str(output),
                        suggested_name=output.name,
                    )
                ],
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=f"failure-{index}",
                ),
            )
            for index, output in enumerate(retained)
        ]
        window._on_execution_finished(
            results,
            {
                "request_id": "multi-retained",
                "file_path": normalized[0],
                "file_paths": normalized,
                "display_name": "Batch conversion (2 files)",
                "total_count": 2,
                "batch": True,
            },
        )

        entries = [window._batch_list_vm.get_file_entry(path) for path in normalized]
        assert all(entry is not None and entry.status == "failed" for entry in entries)
        assert [entry.output_path for entry in entries if entry is not None] == [
            str(retained[0]),
            str(retained[1]),
        ]
        for entry in entries:
            assert entry is not None
            card = BatchEntryItemWidget(entry)
            assert card._primary_action_key == "open_output"
            card.close()
        assert [row.message for row in window._info_area_vm.history_rows if row.message.startswith("failure-")] == [
            "failure-0"
        ]

    @pytest.mark.parametrize(
        ("error_type", "expected_status"),
        [("cancelled", "cancelled"), ("skipped", "skipped")],
    )
    def test_batch_cancelled_or_skipped_result_clears_stale_output(
        self,
        window,
        tmp_path,
        error_type: str,
        expected_status: str,
    ) -> None:
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        source = tmp_path / f"{error_type}.md"
        source.write_text("# input", encoding="utf-8")
        normalized = str(source).replace("\\", "/")
        window._batch_list_vm.add_files([str(source)])
        window._batch_list_vm.set_file_status(
            normalized,
            "completed",
            output_path=str(tmp_path / f"stale-{error_type}.docx"),
        )

        window._on_execution_finished(
            [
                ConversionResult(
                    task_id=f"batch-{error_type}-0",
                    success=False,
                    error=ConversionErrorInfo(error_type=error_type, message=error_type),
                )
            ],
            {
                "request_id": f"batch-{error_type}",
                "file_path": normalized,
                "file_paths": [normalized],
                "display_name": "Batch conversion",
                "total_count": 1,
                "batch": True,
            },
        )

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == expected_status
        assert not entry.output_path

    def test_batch_all_success_with_warning_keeps_success_state_and_warning_tone(self, window, tmp_path) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionDiagnostic, ConversionResult

        first = tmp_path / "first.md"
        second = tmp_path / "second.md"
        first_output = tmp_path / "first.docx"
        second_output = tmp_path / "second.docx"
        for path in (first, second):
            path.write_text("# Title", encoding="utf-8")
        for path in (first_output, second_output):
            path.write_text("ok", encoding="utf-8")
        first_norm = str(first).replace("\\", "/")
        second_norm = str(second).replace("\\", "/")
        window._batch_list_vm.add_files([str(first), str(second)])

        def success_result(
            task_id: str,
            output_path: Path,
            diagnostics: list[ConversionDiagnostic],
        ) -> ConversionResult:
            return ConversionResult(
                task_id=task_id,
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id=f"artifact-{task_id}",
                        kind="primary",
                        staging_path=str(output_path),
                        suggested_name=output_path.name,
                        is_primary=True,
                    )
                ],
                diagnostics=diagnostics,
            )

        results = [
            success_result(
                "batch-warning-0",
                first_output,
                [
                    ConversionDiagnostic(
                        level="warning",
                        message="Numbering was approximated",
                        code="MD2DOCX-NUMBERING-APPROXIMATE",
                    )
                ],
            ),
            success_result(
                "batch-warning-1",
                second_output,
                [ConversionDiagnostic(level="info", message="done", code="MD2DOCX-OK")],
            ),
        ]

        window._on_execution_finished(
            results,
            {
                "request_id": "batch-warning",
                "file_path": first_norm,
                "file_paths": [first_norm, second_norm],
                "display_name": "Batch conversion (2 files)",
                "total_count": 2,
                "batch": True,
            },
        )

        first_entry = window._batch_list_vm.get_file_entry(first_norm)
        second_entry = window._batch_list_vm.get_file_entry(second_norm)
        assert first_entry is not None
        assert second_entry is not None
        assert first_entry.status == "completed"
        assert second_entry.status == "completed"
        summary = window._info_area_vm.task_summary
        assert summary.state == "success"
        assert summary.tone == "warning"
        assert summary.completed_count == 2
        assert summary.failed_count == 0
        history = window._info_area_vm.history_rows
        assert [row.message_type for row in history[-2:]] == ["warning", "warning"]
        assert history[-1].message == f"{first.name}: Numbering was approximated"
        assert history[-1].file_path == str(first_output)

    def test_batch_execution_missing_result_counts_as_processed_failure(self, window, tmp_path) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionResult

        first = tmp_path / "ok.md"
        second = tmp_path / "missing.md"
        output = tmp_path / "ok.docx"
        first.write_text("# OK", encoding="utf-8")
        second.write_text("# Missing", encoding="utf-8")
        output.write_text("ok", encoding="utf-8")
        first_norm = str(first).replace("\\", "/")
        second_norm = str(second).replace("\\", "/")
        window._batch_list_vm.add_files([str(first), str(second)])
        window._batch_list_vm.set_file_status(
            second_norm,
            "completed",
            output_path=str(tmp_path / "stale-missing.docx"),
        )

        window._on_execution_finished(
            [
                ConversionResult(
                    task_id="batch-1-0",
                    success=True,
                    artifacts=[
                        ArtifactManifest(
                            artifact_id="out-1",
                            kind="primary",
                            staging_path=str(output),
                            suggested_name=output.name,
                            is_primary=True,
                        )
                    ],
                )
            ],
            {
                "request_id": "batch-1",
                "file_path": first_norm,
                "file_paths": [first_norm, second_norm],
                "display_name": "Batch conversion (2 files)",
                "total_count": 2,
                "batch": True,
            },
        )

        second_entry = window._batch_list_vm.get_file_entry(second_norm)
        assert second_entry is not None
        assert second_entry.status == "failed"
        assert second_entry.error_message == "Invalid batch conversion result"
        assert not second_entry.output_path
        summary = window._info_area_vm._task_summary
        assert summary.state == "partial"
        assert summary.completed_count == 2
        assert summary.total_count == 2
        assert summary.failed_count == 1

    def test_cancel_active_batch_uses_controller_canonical_parent_once(self, window, tmp_path) -> None:
        from types import SimpleNamespace

        first = tmp_path / "a.md"
        second = tmp_path / "b.md"
        first.write_text("# A", encoding="utf-8")
        second.write_text("# B", encoding="utf-8")
        first_norm = str(first).replace("\\", "/")
        second_norm = str(second).replace("\\", "/")
        window._batch_list_vm.add_files([str(first), str(second)])
        window._batch_list_vm.set_file_status(first_norm, "processing", operation_id="batch-1-0")
        window._batch_list_vm.set_file_status(second_norm, "processing", operation_id="batch-1")
        window._active_threads["batch-1"] = object()  # type: ignore[assignment]
        window._view_model._current_task_id = "batch-1-0"

        cancelled: list[str] = []
        window._view_model._controller = SimpleNamespace(
            has_runtime=True,
            cancel=lambda task_id: cancelled.append(task_id),
            stop=lambda: None,
        )

        window._cancel_active_task()

        assert cancelled == ["batch-1"]

    def test_single_thread_setup_failure_releases_exact_cancellation_reservation(
        self,
        window,
        tmp_path,
        monkeypatch,
    ) -> None:
        import docwen_gui.main_window as main_window_module

        source = tmp_path / "setup-failure.md"
        source.write_text("# Setup failure", encoding="utf-8")
        normalized = str(source).replace("\\", "/")
        outcome = window._view_model.add_files([normalized])
        assert len(outcome.added) == 1
        request, context = window._build_request(
            file_path=normalized,
            target_format="docx",
            action_name="",
            options={},
        )
        context.update({"file_paths": [normalized], "total_count": 1})
        reservation = object()
        released: list[tuple[str, object]] = []
        controller = SimpleNamespace(
            has_runtime=True,
            describe_runtime_capabilities=optimization_capability_projection,
            prepare_execution_cancellation=lambda *_args, **_kwargs: reservation,
            release_execution_cancellation=lambda task_id, handle: released.append((task_id, handle)),
            stop=lambda: None,
        )
        window._view_model._controller = controller
        monkeypatch.setattr(window, "_build_request", lambda **_kwargs: (request, context))

        def fail_thread_setup(**_kwargs: object) -> object:
            raise RuntimeError("QThread setup failed")

        monkeypatch.setattr(main_window_module, "_ExecutionThread", fail_thread_setup)

        window._start_execution(
            file_path=normalized,
            target_format="docx",
            action_name="",
            options={},
        )

        assert released == [(request.request_id, reservation)]
        assert window._active_threads == {}
        assert window._action_area_vm.cancel_visible is False
        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "failed"
        assert entry.error_message == "QThread setup failed"

    def test_aggregate_thread_setup_failure_uses_shared_reservation_rollback(
        self,
        window,
        tmp_path,
        monkeypatch,
    ) -> None:
        import docwen_gui.main_window as main_window_module

        first = tmp_path / "first.pdf"
        second = tmp_path / "second.pdf"
        first.write_bytes(b"%PDF-1.4\n")
        second.write_bytes(b"%PDF-1.4\n")
        paths = [str(first).replace("\\", "/"), str(second).replace("\\", "/")]
        window._batch_list_vm.add_files(paths)
        request, context = window._build_aggregate_request(
            file_paths=paths,
            target_format="pdf",
            action_name="merge_pdfs",
            options={},
        )
        reservation = object()
        prepared: list[tuple[object, bool]] = []
        released: list[tuple[str, object]] = []

        def prepare(candidate: object, *, batch: bool = False) -> object:
            prepared.append((candidate, batch))
            return reservation

        controller = SimpleNamespace(
            has_runtime=True,
            describe_runtime_capabilities=optimization_capability_projection,
            prepare_execution_cancellation=prepare,
            release_execution_cancellation=lambda task_id, handle: released.append((task_id, handle)),
            stop=lambda: None,
        )
        window._view_model._controller = controller
        monkeypatch.setattr(window, "_build_aggregate_request", lambda **_kwargs: (request, context))
        monkeypatch.setattr(window, "_confirm_request_admission", lambda _request: True)

        def fail_thread_setup(**_kwargs: object) -> object:
            raise RuntimeError("Aggregate QThread setup failed")

        monkeypatch.setattr(main_window_module, "_ExecutionThread", fail_thread_setup)

        window._start_aggregate_execution(
            file_paths=paths,
            target_format="pdf",
            action_name="merge_pdfs",
            options={},
        )

        assert prepared == [(request, False)]
        assert released == [(request.request_id, reservation)]
        assert window._active_threads == {}
        assert window._action_area_vm.cancel_visible is False
        assert window._view_model._active_execution_id is None
        entries = [window._batch_list_vm.get_file_entry(path) for path in paths]
        assert all(entry is not None for entry in entries)
        assert [entry.status for entry in entries if entry is not None] == ["failed", "failed"]
        assert [entry.error_message for entry in entries if entry is not None] == [
            "Aggregate QThread setup failed",
            "Aggregate QThread setup failed",
        ]
