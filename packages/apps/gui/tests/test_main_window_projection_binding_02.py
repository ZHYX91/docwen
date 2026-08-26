"""Focused tests split from test_main_window_projection_binding.py."""

from __future__ import annotations

from ._main_window_projection_binding_support import (
    _BATCH_SCAN_LIMIT,
    _PROJECTION_CONVERSION,
    _PROJECTION_HIDDEN,
    ConversionContext,
    FileRef,
    MainWindowUiProjection,
    Path,
    QApplication,
    QDropEvent,
    QMimeData,
    QPointF,
    Qt,
    QUrl,
    RightPanelSlot,
    _emit,
    _file_ref,
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


class TestMainWindowBatchSync:
    @pytest.mark.parametrize(
        ("file_name", "content", "expected_format"),
        [
            ("plain.txt", "plain UTF-8 text\n", "txt"),
            ("notes.md", "# Heading\n\nMarkdown body.\n", "markdown"),
        ],
    )
    def test_batch_entry_fallback_preserves_detected_text_format_and_canonical_workflow(
        self,
        window,
        tmp_path,
        file_name: str,
        content: str,
        expected_format: str,
    ) -> None:
        source = tmp_path / file_name
        source.write_text(content, encoding="utf-8")

        added, failed = window._batch_list_vm.add_files([str(source)])

        assert failed == []
        assert len(added) == 1
        assert window._view_model.files == []
        entry = window._batch_list_vm.get_file_entry(str(source))
        assert entry is not None
        assert entry.detected_format == expected_format
        assert entry.workflow_category == "markdown"
        assert window._batch_list_vm.get_file_display_category(str(source)) == "text"

        request_ref = window._request_file_ref(str(source))
        assert request_ref.format == expected_format
        assert request_ref.category == "markdown"

    def test_batch_sync_reuses_file_ref_warning_and_inspection_without_redetection(
        self, window, tmp_path, monkeypatch
    ) -> None:
        source = tmp_path / "renamed.txt"
        source.write_text("# heading", encoding="utf-8")
        ref = FileRef(
            path=str(source),
            format="markdown",
            category="markdown",
            warning_message="declared TXT, detected Markdown",
            metadata={"inspection": {"decision": "allow_with_warning"}},
        )
        monkeypatch.setattr(
            "docwen_core.detection.inspect_file",
            lambda _path: (_ for _ in ()).throw(AssertionError("batch must reuse FileRef")),
        )

        window._sync_files_from_main_vm([ref])

        entry = window._batch_list_vm.get_file_entry(str(source))
        assert entry is not None
        assert entry.workflow_category == "markdown"
        assert window._batch_list_vm.get_file_display_category(ref.path) == "text"
        assert entry.warning_message == ref.warning_message
        assert entry.metadata == ref.metadata
        assert window._input_area_vm.selection_tone == "warning"
        assert window._input_area_vm.selection_message == ref.warning_message

    def test_batch_tab_switch_updates_selection_and_clears_stale_panels(self, window, tmp_path, qapp) -> None:
        note = tmp_path / "note.md"
        document = tmp_path / "document.docx"
        note.write_text("# Note\n", encoding="utf-8")
        _write_format_fixture(document, "docx")
        window._view_model.mode = "batch"
        window._view_model.add_files([str(note), str(document)])
        qapp.processEvents()

        window._batch_list._activate_tab("document")
        qapp.processEvents()

        assert window._batch_list_vm.current_category == "document"
        assert window._view_model.selected_file is not None
        assert window._view_model.selected_file.path == str(document)
        assert window._conversion_panel_vm.file_list == [str(document).replace("\\", "/")]

        window._batch_list._activate_tab("spreadsheet")
        qapp.processEvents()

        assert window._batch_list_vm.current_category == "spreadsheet"
        assert window._view_model.selected_file is None
        assert window._right_panel_frame.isHidden()

    def test_batch_sync_activates_majority_category_and_selects_visible_file(self, window, tmp_path) -> None:
        first_doc_path = tmp_path / "a.docx"
        second_doc_path = tmp_path / "b.docx"
        sheet_path = tmp_path / "c.xlsx"
        _write_format_fixture(first_doc_path, "docx")
        _write_format_fixture(second_doc_path, "docx")
        _write_format_fixture(sheet_path, "xlsx")
        first_doc = str(first_doc_path)
        second_doc = str(second_doc_path)
        sheet = str(sheet_path)

        window._view_model.add_files([first_doc, second_doc, sheet])
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()

        assert window._batch_list_vm.current_category == "document"
        assert window._batch_list_vm.get_file_count("document") == 2
        selected = window._view_model.selected_file
        assert selected is not None
        assert selected.path == first_doc
        assert window._batch_list.get_current_file() == first_doc.replace("\\", "/")

    def test_input_area_folder_add_syncs_all_files_into_batch_list(self, window, tmp_path) -> None:
        folder = tmp_path / "large"
        folder.mkdir()
        total = _BATCH_SCAN_LIMIT + 5
        for index in range(total):
            (folder / f"doc-{index:03}.txt").write_text("content")
        window._view_model.set_mode("batch")
        window._input_area_vm.set_mode("batch")
        window._input_area_vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"

        window._input_area_vm.add_files([str(folder)])
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()

        assert len(window._view_model.files) == total
        assert window._batch_list_vm.entry_count == total
        assert window._batch_list_vm.current_category == "text"
        selected = window._view_model.selected_file
        assert selected is not None
        assert selected.path.endswith("doc-000.txt")

    def test_input_area_url_drop_large_folder_syncs_all_files_into_batch_list(self, window, tmp_path) -> None:
        folder = tmp_path / "large-url-drop"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        total = _BATCH_SCAN_LIMIT + 9
        supported_files = []
        for index in range(total):
            target_dir = nested if index % 5 == 0 else folder
            file_path = target_dir / f"doc-{index:03}.txt"
            file_path.write_text("content")
            supported_files.append(str(file_path))
        (nested / "ignore.bin").write_text("unsupported")
        window._view_model.set_mode("batch")
        window._input_area_vm.set_mode("batch")
        window._input_area_vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        main_files_changed: list[list[object]] = []
        batch_files_added: list[tuple[list[str], list[str]]] = []
        batch_entry_counts: list[int] = []
        window._view_model.files_changed.connect(lambda refs: main_files_changed.append(list(refs)))
        window._batch_list_vm.files_added.connect(
            lambda added, failed: batch_files_added.append((list(added), list(failed)))
        )
        window._batch_list_vm.entry_count_changed.connect(batch_entry_counts.append)

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(folder))])
        drop_event = QDropEvent(
            QPointF(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        window._input_area.dropEvent(drop_event)
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()

        assert drop_event.isAccepted()
        expected_order = sorted(supported_files, key=lambda path: path.casefold())
        assert [ref.path for ref in window._view_model.files] == expected_order
        assert window._batch_list_vm.get_files() == [path.replace("\\", "/") for path in expected_order]
        assert len(main_files_changed) == 1
        assert len(batch_files_added) == 1
        assert len(batch_entry_counts) == 1
        assert len(batch_files_added[0][0]) == total
        assert batch_files_added[0][1] == []
        assert batch_entry_counts == [total]
        assert len(window._view_model.files) == total
        assert window._batch_list_vm.entry_count == total
        assert window._batch_list_vm.current_category == "text"
        assert window._input_area_vm.selection_tone == "warning"
        assert "ignore.bin" not in "\n".join(window._batch_list_vm.get_files())

    def test_switching_to_single_restores_current_batch_file_selection(self, window, tmp_path) -> None:
        first_doc_path = tmp_path / "a.docx"
        second_doc_path = tmp_path / "b.docx"
        _write_format_fixture(first_doc_path, "docx")
        _write_format_fixture(second_doc_path, "docx")
        first_doc = str(first_doc_path)
        second_doc = str(second_doc_path)

        window._view_model.set_mode("batch")
        window._view_model.add_files([first_doc, second_doc])
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()

        window._view_model.clear_selected_file()
        assert window._view_model.selected_file is None

        window._view_model.set_mode("single")
        if app_instance:
            app_instance.processEvents()

        selected = window._view_model.selected_file
        assert selected is not None
        assert selected.path == first_doc
        assert window._view_model.ui_projection.right_panel_visible is True

    def test_conversion_panel_file_list_respects_batch_reorder(self, window, tmp_path) -> None:
        first_doc_path = tmp_path / "a.docx"
        second_doc_path = tmp_path / "b.docx"
        _write_format_fixture(first_doc_path, "docx")
        _write_format_fixture(second_doc_path, "docx")
        first_doc = str(first_doc_path)
        second_doc = str(second_doc_path)

        window._view_model.set_mode("batch")
        window._view_model.add_files([first_doc, second_doc])
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()

        list_widget = window._batch_list._tabs["document"]
        list_widget.setCurrentRow(0)
        assert window._batch_list._move_current_item_down() is True

        projection = MainWindowUiProjection(
            left_panel_visible=True,
            right_panel_visible=True,
            right_panel_slot=RightPanelSlot.CONVERSION,
            center_action_visible=True,
            info_area_visible=True,
            conversion_context=ConversionContext(
                category="document",
                current_format="docx",
                file_path=second_doc,
            ),
            template_context=None,
        )
        _emit(window, projection)

        assert window._conversion_panel_vm.file_list == [
            second_doc.replace("\\", "/"),
            first_doc.replace("\\", "/"),
        ]


class TestWidgetPersistence:
    def test_right_panel_widgets_survive_hide_show(self, window, right_stack) -> None:
        """Hiding the right panel must not destroy/recreate the stack widgets."""
        _emit(window, _PROJECTION_CONVERSION)
        cp_before = window.conversion_panel
        _emit(window, _PROJECTION_HIDDEN)
        _emit(window, _PROJECTION_CONVERSION)
        cp_after = window.conversion_panel
        assert cp_before is cp_after, "Conversion panel must not be re-created on visibility change"


class TestActionOnlyBinding:
    def test_other_file_hides_right_panel_but_sets_action_area(self, window, right_frame) -> None:
        """Old selector route ``other`` was action_only: no right panel, but action area remains usable."""
        window._view_model.set_selected_file(_file_ref("/tmp/sample.epub", "other", "epub"))

        assert right_frame.isHidden() or not right_frame.isVisible()
        assert window._action_area_vm.visible is True
        assert window._action_area_vm.file_type == "epub"
