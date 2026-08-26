"""Focused tests split from test_main_window_projection_binding.py."""

from __future__ import annotations

from ._main_window_projection_binding_support import (
    _bind_admitted_ref,
    _write_format_fixture,
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
    @pytest.mark.parametrize("include_skipped", [False, True])
    def test_partially_completed_cancelled_batch_reports_real_counts(
        self,
        window,
        tmp_path,
        include_skipped: bool,
    ) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult
        from docwen_gui.i18n import t as _t
        from docwen_gui.main_window import _normalize_path

        sources = [tmp_path / f"input-{index}.docx" for index in range(4 if include_skipped else 3)]
        for source in sources:
            source.write_bytes(b"fixture")
        paths = [_normalize_path(str(source)) for source in sources]
        window._batch_list_vm.add_files([str(source) for source in sources])
        output = tmp_path / "input-0.md"
        output.write_text("# converted", encoding="utf-8")

        results = [
            ConversionResult(
                task_id="batch-cancel-0",
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
        ]
        if include_skipped:
            results.append(
                ConversionResult(
                    task_id="batch-skip-1",
                    success=False,
                    error=ConversionErrorInfo(error_type="skipped", message="Route was skipped"),
                )
            )
        results.extend(
            [
                ConversionResult(
                    task_id="batch-cancel-1",
                    success=False,
                    error=ConversionErrorInfo(error_type="cancelled", message="Task was cancelled"),
                ),
                ConversionResult(
                    task_id="batch-cancel-2",
                    success=False,
                    error=ConversionErrorInfo(error_type="cancelled", message="Task was cancelled"),
                ),
            ]
        )

        window._on_execution_finished(
            results,
            {
                "request_id": "batch-cancel",
                "file_paths": paths,
                "display_name": f"{len(sources)} files",
                "total_count": len(sources),
                "batch": True,
            },
        )

        summary = window._info_area_vm.task_summary
        assert summary.state == "cancelled"
        assert summary.completed_count == 1
        assert summary.total_count == len(sources)
        assert summary.failed_count == 0
        assert summary.skipped_count == int(include_skipped)
        assert summary.cancelled_count == 2
        assert window._info_area_vm.status_meta_text == _t("info_area.task_state_cancelled", "Cancelled")
        assert window._info_area_vm.status_summary_text == _t(
            "components.info_area.batch_completed",
            "Batch finished: {success} succeeded, {failed} failed, {skipped} skipped, {cancelled} cancelled",
            success=1,
            failed=0,
            skipped=int(include_skipped),
            cancelled=2,
        )
        window._info_area_vm._clear_transient_state()
        window._info_area_vm._refresh_status()
        assert _t("info_area.task_cancelled_count", "Cancelled: {cancelled}", cancelled=2) in (
            window._info_area_vm.status_summary_text
        )
        expected_statuses = ["completed"]
        if include_skipped:
            expected_statuses.append("skipped")
        expected_statuses.extend(["cancelled", "cancelled"])
        assert [window._batch_list_vm.get_file_entry(path).status for path in paths] == expected_statuses

    def test_named_validate_defers_markdown_target_to_runtime_catalog(self, window, tmp_path, monkeypatch) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}
        calls: list[dict[str, object]] = []

        def fake_start_execution(**kwargs: object) -> None:
            calls.append(kwargs)

        monkeypatch.setattr(window, "_start_execution", fake_start_execution)

        window._handle_named_action_requested(
            "validate",
            str(source),
            {"symbol_pairing": True},
        )

        assert calls == [
            {
                "file_path": str(source),
                "target_format": "",
                "action_name": "validate",
                "options": {"symbol_pairing": True},
            }
        ]

    def test_named_validate_routes_docx_context_to_canonical_docx_validate(self, window, tmp_path, monkeypatch) -> None:
        source = tmp_path / "sample.docx"
        _write_format_fixture(source, "docx")
        _bind_admitted_ref(window, source, "document", "docx")
        calls: list[dict[str, object]] = []

        def fake_start_execution(**kwargs: object) -> None:
            calls.append(kwargs)

        monkeypatch.setattr(window, "_start_execution", fake_start_execution)

        window._handle_named_action_requested(
            "validate",
            str(source),
            {"symbol_pairing": True},
        )

        assert calls == [
            {
                "file_path": str(source),
                "target_format": "",
                "action_name": "validate",
                "options": {"symbol_pairing": True},
            }
        ]

    def test_named_split_pdf_routes_to_pdf_action(self, window, tmp_path, monkeypatch) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "sample.pdf"
        source.write_bytes(b"%PDF-1.4\n")
        window._file_contexts = {_normalize_path(str(source)): ("pdf", "layout")}
        calls: list[dict[str, object]] = []

        def fake_start_execution(**kwargs: object) -> None:
            calls.append(kwargs)

        monkeypatch.setattr(window, "_start_execution", fake_start_execution)

        window._handle_named_action_requested(
            "split_pdf",
            str(source),
            {"split_mode": "custom", "pages": [1]},
        )

        assert calls == [
            {
                "file_path": str(source),
                "target_format": "",
                "action_name": "split_pdf",
                "options": {"split_mode": "custom", "pages": [1]},
            }
        ]

    def test_named_merge_pdfs_routes_to_aggregate_execution(self, window, tmp_path, monkeypatch) -> None:
        first = tmp_path / "b.pdf"
        second = tmp_path / "a.pdf"
        third = tmp_path / "ignored.ofd"
        for path in (first, second):
            path.write_bytes(b"%PDF-1.4\n")
        third.write_text("not a pdf", encoding="utf-8")

        window._batch_list_vm.add_files([str(first), str(second), str(third)])
        window._batch_list_vm.reorder_manual("layout", [str(second).replace("\\", "/"), str(first).replace("\\", "/")])
        calls: list[dict[str, object]] = []

        def fake_start_aggregate_execution(**kwargs: object) -> None:
            calls.append(kwargs)

        monkeypatch.setattr(window, "_start_aggregate_execution", fake_start_aggregate_execution)

        window._handle_named_action_requested("merge_pdfs", str(first), {})

        assert calls == [
            {
                "file_paths": [str(second).replace("\\", "/"), str(first).replace("\\", "/")],
                "target_format": "",
                "action_name": "merge_pdfs",
                "options": {},
            }
        ]

    def test_conversion_panel_merge_pdf_button_starts_aggregate_execution(self, window, tmp_path, monkeypatch) -> None:
        first = tmp_path / "b.pdf"
        second = tmp_path / "a.pdf"
        for path in (first, second):
            path.write_bytes(b"%PDF-1.4\n")

        window._batch_list_vm.add_files([str(first), str(second)])
        window._batch_list_vm.reorder_manual("layout", [str(second).replace("\\", "/"), str(first).replace("\\", "/")])
        window._conversion_panel_vm.set_file_info("layout", "pdf", file_path=str(first))
        calls: list[dict[str, object]] = []
        monkeypatch.setattr(window, "_start_aggregate_execution", lambda **kwargs: calls.append(kwargs))

        assert window._conversion_panel._merge_pdfs_button is not None
        window._conversion_panel._merge_pdfs_button.click()

        assert calls == [
            {
                "file_paths": [str(second).replace("\\", "/"), str(first).replace("\\", "/")],
                "target_format": "",
                "action_name": "merge_pdfs",
                "options": {},
            }
        ]

    def test_conversion_panel_merge_tables_button_starts_aggregate_execution(
        self, window, tmp_path, monkeypatch
    ) -> None:
        from openpyxl import Workbook

        first = tmp_path / "base.xlsx"
        second = tmp_path / "collect.xlsx"
        for path, name in ((first, "Alice"), (second, "Charlie")):
            workbook = Workbook()
            sheet = workbook.active
            assert sheet is not None
            sheet.append(["Name"])
            sheet.append([name])
            workbook.save(path)
            workbook.close()

        window._batch_list_vm.add_files([str(first), str(second)])
        window._conversion_panel_vm.merge_mode = 2
        window._conversion_panel_vm.set_file_info("spreadsheet", "xlsx", file_path=str(second))
        calls: list[dict[str, object]] = []
        monkeypatch.setattr(window, "_start_aggregate_execution", lambda **kwargs: calls.append(kwargs))

        assert window._conversion_panel._merge_tables_button is not None
        window._conversion_panel._merge_tables_button.click()

        assert calls == [
            {
                "file_paths": [str(second).replace("\\", "/"), str(first).replace("\\", "/")],
                "target_format": "",
                "action_name": "merge_tables",
                "options": {"merge_mode": "col"},
            }
        ]

    def test_conversion_panel_merge_tiff_button_starts_aggregate_execution(self, window, tmp_path, monkeypatch) -> None:
        from PIL import Image

        first = tmp_path / "red.png"
        second = tmp_path / "blue.jpg"
        Image.new("RGB", (8, 8), (255, 0, 0)).save(first)
        Image.new("RGB", (8, 8), (0, 0, 255)).save(second)

        window._batch_list_vm.add_files([str(first), str(second)])
        window._conversion_panel_vm.tiff_mode = "rgb"
        window._conversion_panel_vm.set_file_info("image", "png", file_path=str(first))
        calls: list[dict[str, object]] = []
        monkeypatch.setattr(window, "_start_aggregate_execution", lambda **kwargs: calls.append(kwargs))

        assert window._conversion_panel._merge_tiff_button is not None
        window._conversion_panel._merge_tiff_button.click()

        assert calls == [
            {
                "file_paths": [str(first).replace("\\", "/"), str(second).replace("\\", "/")],
                "target_format": "",
                "action_name": "merge_images_to_tiff",
                "options": {"mode": "RGB"},
            }
        ]

    def test_named_merge_pdfs_requires_two_pdf_targets(self, window, tmp_path, monkeypatch) -> None:
        source = tmp_path / "only.pdf"
        source.write_bytes(b"%PDF-1.4\n")
        window._batch_list_vm.add_files([str(source)])
        calls: list[dict[str, object]] = []
        monkeypatch.setattr(window, "_start_aggregate_execution", lambda **kwargs: calls.append(kwargs))

        window._handle_named_action_requested("merge_pdfs", str(source), {})

        assert calls == []

    def test_batch_mode_conversion_request_routes_to_batch_execution(self, window, tmp_path, monkeypatch) -> None:
        first = tmp_path / "b.md"
        second = tmp_path / "a.md"
        for path in (first, second):
            path.write_text("# Title", encoding="utf-8")
        document = tmp_path / "excluded.docx"
        _write_format_fixture(document, "docx")

        window._view_model.mode = "batch"
        window._batch_list_vm.add_files([str(first), str(second), str(document)])
        window._batch_list_vm.reorder_manual("text", [str(second).replace("\\", "/"), str(first).replace("\\", "/")])
        batch_calls: list[dict[str, object]] = []
        single_calls: list[dict[str, object]] = []
        monkeypatch.setattr(window, "_start_batch_execution", lambda **kwargs: batch_calls.append(kwargs))
        monkeypatch.setattr(window, "_start_execution", lambda **kwargs: single_calls.append(kwargs))

        window._handle_conversion_requested(
            "docx",
            str(first),
            {"remove_numbering": True},
            origin="conversion_panel",
        )

        assert single_calls == []
        assert batch_calls == [
            {
                "file_paths": [str(second).replace("\\", "/"), str(first).replace("\\", "/")],
                "target_format": "docx",
                "action_name": "",
                "options": {"remove_numbering": True},
            }
        ]

    def test_conversion_origins_do_not_leak_center_optimization_into_right_panel(
        self,
        window,
        tmp_path,
        monkeypatch,
    ) -> None:
        source = tmp_path / "sample.png"
        source.write_bytes(b"\x89PNG\r\n\x1a\n")
        from docwen_gui.view_models._optimization_filter import (
            OptimizationChoice,
            OptimizationChoicesResult,
        )

        window._action_area_vm._optimization_choices_result = OptimizationChoicesResult(
            status="ready",
            choices=(
                OptimizationChoice(
                    id="invoice_cn",
                    label="Invoice CN",
                    action_name="invoice_cn",
                    bindings=(),
                    route_options=("locale", "yaml_key_labels"),
                ),
            ),
        )
        window._action_area_vm.optimize_for_type = "invoice_cn"
        calls: list[dict[str, object]] = []
        monkeypatch.setattr(window, "_start_execution", lambda **kwargs: calls.append(kwargs))

        window._handle_conversion_panel_conversion_requested("jpg", str(source), {"quality": 90})
        window._handle_action_area_conversion_requested("md", str(source), {"to_md_enable_ocr": True})

        assert [call["action_name"] for call in calls] == ["", "invoice_cn"]

    def test_batch_request_contains_ordered_runtime_refs(self, window, tmp_path) -> None:
        first = tmp_path / "b.txt"
        second = tmp_path / "a.md"
        first.write_text("# B", encoding="utf-8")
        second.write_text("# A", encoding="utf-8")
        first_norm = str(first).replace("\\", "/")
        second_norm = str(second).replace("\\", "/")
        window._file_contexts = {
            first_norm: ("markdown", "markdown"),
            second_norm: ("markdown", "markdown"),
        }

        request, context = window._build_batch_request(
            file_paths=[second_norm, first_norm],
            target_format="docx",
            action_name="",
            options={"add_numbering": False},
        )

        assert request.target_format == "docx"
        assert request.action_name == ""
        assert request.options["add_numbering"] is False
        assert [ref.path.replace("\\", "/") for ref in request.input_refs] == [second_norm, first_norm]
        assert [ref.format for ref in request.input_refs] == ["markdown", "markdown"]
        assert [ref.category for ref in request.input_refs] == ["markdown", "markdown"]
        assert context["batch"] is True
        assert context["file_paths"] == [second_norm, first_norm]
        assert context["total_count"] == 2

    def test_batch_execution_partial_result_sets_rows_and_info_summary(self, window, tmp_path) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult
        from docwen_gui.i18n import t as _t

        first = tmp_path / "ok.md"
        second = tmp_path / "bad.md"
        output = tmp_path / "ok.docx"
        first.write_text("# OK", encoding="utf-8")
        second.write_text("# Bad", encoding="utf-8")
        output.write_text("ok", encoding="utf-8")
        first_norm = str(first).replace("\\", "/")
        second_norm = str(second).replace("\\", "/")
        window._batch_list_vm.add_files([str(first), str(second)])
        window._batch_list_vm.set_file_status(
            second_norm,
            "completed",
            output_path=str(tmp_path / "stale.docx"),
        )

        results = [
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
            ),
            ConversionResult(
                task_id="batch-1-1",
                success=False,
                error=ConversionErrorInfo(error_type="conversion_failed", message="boom"),
            ),
        ]

        window._on_execution_finished(
            results,
            {
                "request_id": "batch-1",
                "file_path": first_norm,
                "file_paths": [first_norm, second_norm],
                "display_name": "Batch conversion (2 files)",
                "target_format": "docx",
                "action_name": "",
                "options": {},
                "total_count": 2,
                "batch": True,
            },
        )

        first_entry = window._batch_list_vm.get_file_entry(first_norm)
        second_entry = window._batch_list_vm.get_file_entry(second_norm)
        assert first_entry is not None
        assert second_entry is not None
        assert first_entry.status == "completed"
        assert first_entry.output_path == str(output)
        assert second_entry.status == "failed"
        assert second_entry.error_message == "boom"
        assert not second_entry.output_path
        summary = window._info_area_vm._task_summary
        assert summary.state == "partial"
        assert summary.tone == "warning"
        assert summary.completed_count == 2
        assert summary.total_count == 2
        assert summary.failed_count == 1
        assert summary.navigate_path == second_norm
        assert window._info_area_vm.status_meta_text == _t("info_area.task_state_partial", "Partial failure")
        assert [action["action_key"] for action in window._info_area_vm.guide_actions] == [
            "open_output_dir",
            "view_failed_details",
            "retry_failed",
            "add_more_files",
        ]
        history = window._info_area_vm.history_rows
        assert [row.message_type for row in history[-2:]] == ["warning", "danger"]
        assert history[-2].message == _t(
            "components.info_area.batch_completed",
            "Batch finished: {success} succeeded, {failed} failed, {skipped} skipped, {cancelled} cancelled",
            success=1,
            failed=1,
            skipped=0,
            cancelled=0,
        )
        assert history[-1].message == "boom"
        assert history[-1].file_path == second_norm
        assert history[-1].navigate_file_path == second_norm
