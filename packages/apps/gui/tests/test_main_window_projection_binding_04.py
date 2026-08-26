"""Focused tests split from test_main_window_projection_binding.py."""

from __future__ import annotations

import pytest

from ._main_window_projection_binding_support import (
    _XLSX_TEMPLATE_ID,
    FakeController,
    FileRef,
    _load_request_templates,
    _make_window_with_config,
    _write_format_fixture,
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
    def test_xlsx_template_selection_is_added_to_spreadsheet_request(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}
        _load_request_templates(window)
        assert window._template_selector.activate_and_select("xlsx") is True

        request, context = window._build_request(
            file_path=str(source),
            target_format="xlsx",
            action_name="",
            options={},
        )

        assert request.options["template_name"] == _XLSX_TEMPLATE_ID
        assert context["options"]["template_name"] == _XLSX_TEMPLATE_ID

    def test_spreadsheet_template_transition_clears_stale_actions_and_injects_template(
        self,
        window,
        tmp_path,
    ) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}
        window._view_model.set_selected_file(FileRef(path=str(source), format="markdown", category="markdown"))
        _load_request_templates(window, docx=False)
        window._action_area_vm.optimize_for_type = "invoice_cn"
        window._action_area_vm.setup_for_aggregate("merge_pdfs", ["/a.pdf", "/b.pdf"])

        window._on_main_window_template_tab_changed("xlsx", "docx")
        assert window._template_selector.activate_and_select("xlsx") is True
        request, context = window._build_request(
            file_path=str(source),
            target_format="xlsx",
            action_name=window._action_area_vm.action_name,
            options=window._action_area_vm.collect_options(),
        )

        assert request.action_name == ""
        assert request.options["template_name"] == _XLSX_TEMPLATE_ID
        assert context["options"]["template_name"] == _XLSX_TEMPLATE_ID

    def test_xlsx_template_selection_is_added_to_csv_request(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}
        _load_request_templates(window)
        assert window._template_selector.activate_and_select("xlsx") is True

        request, context = window._build_request(
            file_path=str(source),
            target_format="csv",
            action_name="",
            options={},
        )

        assert request.options["template_name"] == _XLSX_TEMPLATE_ID
        assert context["options"]["template_name"] == _XLSX_TEMPLATE_ID

    def test_single_request_uses_output_preferences(self, qapp, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        output_dir = tmp_path / "exports"
        window = _make_window_with_config(
            qapp,
            {
                "output.directory.mode": "custom",
                "output.directory.custom_path": str(output_dir),
                "output.directory.create_date_subfolder": True,
                "output.directory.date_folder_format": "%Y%m%d",
                "output.behavior.auto_open_folder": True,
            },
        )
        try:
            source = tmp_path / "note.md"
            source.write_text("# Title", encoding="utf-8")
            window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}

            request, context = window._build_request(
                file_path=str(source),
                target_format="docx",
                action_name="",
                options={},
            )

            assert request.output_policy.output_dir == str(output_dir)
            assert request.output_policy.date_subfolder == "compact"
            assert request.output_policy.overwrite_mode == "rename"
            assert request.output_policy.open_after_done is True
            assert context["open_after_done"] is True
        finally:
            window.close()

    def test_relative_custom_output_is_frozen_against_request_working_directory(
        self, qapp, tmp_path, monkeypatch
    ) -> None:
        working_dir = tmp_path / "working"
        working_dir.mkdir()
        monkeypatch.chdir(working_dir)
        window = _make_window_with_config(
            qapp,
            {
                "output.directory.mode": "custom",
                "output.directory.custom_path": "relative-output",
            },
        )
        try:
            source = tmp_path / "note.md"
            source.write_text("# Title", encoding="utf-8")

            request, _context = window._build_request(
                file_path=str(source),
                target_format="docx",
                action_name="",
                options={},
            )

            assert request.output_policy.output_dir == str((working_dir / "relative-output").resolve())
        finally:
            window.close()

    def test_output_settings_failure_blocks_request_instead_of_falling_back(
        self,
        window,
        monkeypatch,
        tmp_path,
    ) -> None:
        from docwen_gui.i18n import t
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}

        controller = window._view_model.controller
        assert controller is not None

        def fail_output_settings(key: str, default: object = None) -> object:
            del default
            if key.startswith("output."):
                raise OSError("corrupt settings store")
            return None

        monkeypatch.setattr(controller.config_port, "get", fail_output_settings)
        window._start_execution(
            file_path=str(source),
            target_format="docx",
            action_name="",
            options={},
        )

        assert not window._active_threads
        assert window._info_area_vm.history_rows[-1].message == t("main_window.output_settings_unavailable")
        assert window._info_area_vm.history_rows[-1].message_type == "error"

    def test_template_registry_failure_is_not_rendered_as_successful_empty_catalog(
        self,
        qapp,
        monkeypatch,
    ) -> None:
        from docwen_gui.i18n import t
        from docwen_gui.main_window import MainWindow
        from docwen_gui.view_models.main_window_vm import MainWindowViewModel
        from docwen_runtime.templates import TemplateRegistry

        def fail_registry() -> object:
            raise OSError("template catalog unreadable")

        monkeypatch.setattr(TemplateRegistry, "default", fail_registry)
        candidate = MainWindow(view_model=MainWindowViewModel(controller=FakeController()))  # type: ignore[arg-type]
        candidate.setup_ui()
        try:
            assert candidate._template_selector is not None
            for template_type in ("docx", "xlsx"):
                selector = candidate._template_selector.get_selector(template_type)
                assert selector is not None
                assert selector._load_error == (
                    t("components.template_selector.unavailable"),
                    t("main_window.template_catalog_failed"),
                )
                assert selector._list.count() == 0
                assert selector._empty_action_button.isEnabled() is False
            assert candidate._info_area_vm.history_rows[-1].message == t("main_window.template_catalog_failed")
        finally:
            candidate.close()

    def test_aggregate_request_uses_output_preferences(self, qapp, tmp_path) -> None:
        output_dir = tmp_path / "exports"
        window = _make_window_with_config(
            qapp,
            {
                "output.directory.mode": "custom",
                "output.directory.custom_path": str(output_dir),
                "output.directory.create_date_subfolder": True,
                "output.directory.date_folder_format": "%Y-%m-%d",
                "output.behavior.auto_open_folder": False,
            },
        )
        try:
            first = tmp_path / "a.pdf"
            second = tmp_path / "b.pdf"
            first.write_bytes(b"%PDF-1.4\n")
            second.write_bytes(b"%PDF-1.4\n")
            window._batch_list_vm.add_files([str(first), str(second)])

            request, context = window._build_aggregate_request(
                file_paths=[str(first), str(second)],
                target_format="pdf",
                action_name="merge_pdfs",
                options={},
            )

            assert request.output_policy.output_dir == str(output_dir)
            assert request.output_policy.date_subfolder == "iso"
            assert request.output_policy.overwrite_mode == "rename"
            assert request.output_policy.open_after_done is False
            assert context["open_after_done"] is False
        finally:
            window.close()

    def test_success_callback_auto_opens_output_when_enabled(self, window, tmp_path, monkeypatch) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionResult
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        output = tmp_path / "out" / "note.docx"
        output.parent.mkdir()
        output.write_text("ok", encoding="utf-8")
        source.write_text("# Title", encoding="utf-8")
        window._batch_list_vm.add_files([str(source)])

        opened: list[tuple[str, bool]] = []
        monkeypatch.setattr(
            window, "_open_path", lambda target, *, open_parent=False: opened.append((target, open_parent)) or True
        )
        result = ConversionResult(
            task_id="task-1",
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id="artifact-1",
                    kind="primary",
                    staging_path=str(output),
                    suggested_name=output.name,
                    is_primary=True,
                )
            ],
        )

        window._on_execution_finished(
            result,
            {
                "request_id": "task-1",
                "file_path": _normalize_path(str(source)),
                "display_name": source.name,
                "total_count": 1,
                "open_after_done": True,
            },
        )

        assert opened == [(str(output), True)]

    def test_success_callback_projects_warning_diagnostics_to_info_area(self, window, tmp_path) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionDiagnostic, ConversionResult
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "rules.docx"
        output = tmp_path / "rules.md"
        _write_format_fixture(source, "docx")
        output.write_text("---\n---\n", encoding="utf-8")
        normalized = _normalize_path(str(source))
        # This test exercises result projection, not container admission. Use
        # an explicit resolver instead of asking the detector to inspect a
        # deliberately truncated four-byte ZIP fixture.
        window._batch_list_vm.add_files(
            [str(source)],
            file_resolver=lambda _path: {
                "detected_format": "docx",
                "workflow_category": "document",
            },
        )

        result = ConversionResult(
            task_id="gongwen-1",
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id="gongwen-output",
                    kind="primary",
                    staging_path=str(output),
                    suggested_name=output.name,
                    is_primary=True,
                )
            ],
            diagnostics=[
                ConversionDiagnostic(
                    level="warning",
                    message="缺少必需字段：成文日期",
                    code="GONGWEN-NEEDS-REVIEW",
                ),
                ConversionDiagnostic(level="info", message="done", code="GONGWEN-OK"),
            ],
        )

        window._on_execution_finished(
            result,
            {
                "request_id": "gongwen-1",
                "file_path": normalized,
                "display_name": source.name,
                "total_count": 1,
            },
        )

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed"
        assert entry.output_path == str(output)
        history = window._info_area_vm.history_rows
        assert [row.message_type for row in history[-2:]] == ["warning", "warning"]
        assert history[-1].message == "缺少必需字段：成文日期"
        assert history[-1].file_path == str(output)
        warning_row = window._info_area.get_history_row_widget(len(history) - 1)
        assert warning_row is not None
        assert warning_row.property("infoStatusTone") == "warning"
        assert warning_row.toolTip() == "缺少必需字段：成文日期"
        summary = window._info_area_vm.task_summary
        assert summary.state == "success"
        assert summary.tone == "warning"
        assert summary.completed_count == 1
        assert summary.failed_count == 0
        assert summary.navigate_path == str(output)

    def test_failed_result_with_retained_auxiliary_exposes_output_without_marking_success(
        self,
        window,
        tmp_path,
        monkeypatch,
    ) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "legacy.doc"
        retained = tmp_path / "out" / "legacy_fromDoc.docx"
        source.write_text("legacy", encoding="utf-8")
        retained.parent.mkdir()
        retained.write_text("preserved", encoding="utf-8")
        normalized = _normalize_path(str(source))
        window._batch_list_vm.add_files([str(source)])
        opened: list[tuple[str, bool]] = []
        monkeypatch.setattr(
            window,
            "_open_path",
            lambda target, *, open_parent=False: opened.append((target, open_parent)) or True,
        )

        result = ConversionResult(
            task_id="failed-with-output",
            success=False,
            artifacts=[
                ArtifactManifest(
                    artifact_id="missing-primary",
                    kind="primary",
                    staging_path=str(tmp_path / "missing-primary.docx"),
                    suggested_name="missing-primary.docx",
                    is_primary=True,
                ),
                ArtifactManifest(
                    artifact_id="retained-hub",
                    kind="auxiliary",
                    staging_path=str(retained),
                    suggested_name=retained.name,
                    is_primary=False,
                ),
            ],
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message="downstream conversion failed",
            ),
        )
        window._on_execution_finished(
            result,
            {
                "request_id": result.task_id,
                "file_path": normalized,
                "display_name": source.name,
                "total_count": 1,
                "open_after_done": True,
            },
        )

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "failed"
        assert entry.error_message == "downstream conversion failed"
        assert entry.output_path == str(retained)
        assert window._batch_list_vm.get_failed_file_count() == 1
        history = window._info_area_vm.history_rows
        assert history[-1].message_type == "danger"
        assert history[-1].file_path == str(retained)
        assert history[-1].navigate_file_path == str(retained)
        summary = window._info_area_vm.task_summary
        assert summary.state == "failed"
        assert summary.failed_count == 1
        assert summary.navigate_path == normalized
        assert summary.navigation_kind == "failed"
        assert window._info_area_vm.guide_actions == [
            {"action_key": "open_output_dir", "target_path": str(retained.parent)},
            {"action_key": "view_failed_details", "target_path": normalized},
            {"action_key": "retry_failed", "target_path": ""},
            {"action_key": "add_more_files", "target_path": ""},
        ]
        assert opened == []

    def test_cancelled_result_marks_batch_entry_cancelled_not_failed(self, window, tmp_path) -> None:
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult
        from docwen_gui.i18n import t as _t
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        stale_output = tmp_path / "stale.docx"
        source.write_text("# Title", encoding="utf-8")
        stale_output.write_text("stale", encoding="utf-8")
        normalized = _normalize_path(str(source))
        window._batch_list_vm.add_files([str(source)])

        result = ConversionResult(
            task_id="task-1",
            success=False,
            artifacts=[
                ArtifactManifest(
                    artifact_id="stale",
                    kind="auxiliary",
                    staging_path=str(stale_output),
                    suggested_name=stale_output.name,
                )
            ],
            error=ConversionErrorInfo(error_type="cancelled", message="Task was cancelled"),
        )
        window._on_execution_finished(
            result,
            {
                "request_id": "task-1",
                "file_path": normalized,
                "display_name": source.name,
                "total_count": 1,
            },
        )

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "cancelled"
        assert not entry.output_path
        assert window._batch_list_vm.get_failed_file_count() == 0
        assert window._info_area_vm.status_meta_text == _t("info_area.task_state_cancelled", "Cancelled")
        assert window._info_area_vm.status_tone == "warning"
        assert window._info_area_vm._task_summary.cancelled_count == 1
        assert window._info_area_vm.guide_actions == [{"action_key": "add_more_files", "target_path": ""}]
