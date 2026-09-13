"""Focused tests split from test_main_window_projection_binding.py."""

from __future__ import annotations

from ._main_window_projection_binding_support import (
    _PROJECTION_BATCH_HIDDEN_RIGHT,
    _PROJECTION_CONVERSION,
    _PROJECTION_HIDDEN,
    _PROJECTION_TEMPLATE,
    ConversionContext,
    FileRef,
    MainWindowUiProjection,
    Path,
    QApplication,
    RightPanelSlot,
    _emit,
    _file_ref,
    _make_window_with_config,
    _root_grid,
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


def test_ipc_file_received_does_not_add_an_activity_record(window) -> None:
    before = list(window._info_area_vm.history_rows)
    window._on_ipc_file_received("C:/tmp/example.md")
    assert window._info_area_vm.history_rows == before


class TestPanelVisibility:
    def test_initial_single_no_file_state_hides_left_panel(self, window, left_frame) -> None:
        assert window._view_model.mode == "single"
        assert window._view_model.selected_file is None
        assert left_frame.isHidden() or not left_frame.isVisible()
        assert _root_grid(window).columnStretch(0) == 0

    def test_no_file_hides_side_panel_frames(self, window, left_frame, right_frame) -> None:
        _emit(window, _PROJECTION_HIDDEN)
        # When hidden, side panel frames are explicitly setVisible(False).
        assert left_frame.isHidden() or not left_frame.isVisible()
        assert right_frame.isHidden() or not right_frame.isVisible()

    def test_batch_projection_makes_left_panel_visible(self, window, left_frame) -> None:
        _emit(window, _PROJECTION_BATCH_HIDDEN_RIGHT)
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()

        assert left_frame.isHidden() is False

    def test_template_projection_makes_right_panel_visible(self, window, right_frame) -> None:
        _emit(window, _PROJECTION_TEMPLATE)
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()
        # Right panel should NOT be hidden after template projection.
        assert right_frame.isHidden() is False

    def test_conversion_projection_makes_right_panel_visible(self, window, right_frame) -> None:
        _emit(window, _PROJECTION_CONVERSION)
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()
        assert right_frame.isHidden() is False

    def test_hidden_left_panel_does_not_reserve_grid_stretch(self, window, left_frame) -> None:
        _emit(window, _PROJECTION_HIDDEN)

        assert left_frame.isHidden() or not left_frame.isVisible()
        assert _root_grid(window).columnStretch(0) == 0

    def test_visible_left_panel_stays_fixed_when_expansion_disabled(self, window, left_frame) -> None:
        _emit(window, _PROJECTION_BATCH_HIDDEN_RIGHT)

        assert left_frame.isHidden() is False
        assert _root_grid(window).columnStretch(0) == 0

        _emit(window, _PROJECTION_HIDDEN)
        assert _root_grid(window).columnStretch(0) == 0

    def test_hidden_right_panel_does_not_reserve_grid_stretch(self, window, right_frame) -> None:
        _emit(window, _PROJECTION_HIDDEN)

        assert right_frame.isHidden() or not right_frame.isVisible()
        assert _root_grid(window).columnStretch(2) == 0

    def test_visible_right_panel_stays_fixed_when_expansion_disabled(self, window, right_frame) -> None:
        _emit(window, _PROJECTION_CONVERSION)

        assert right_frame.isHidden() is False
        assert _root_grid(window).columnStretch(2) == 0

        _emit(window, _PROJECTION_HIDDEN)
        assert _root_grid(window).columnStretch(2) == 0


class TestRightSlotSwitching:
    def test_template_slot_shows_template_selector(self, window, right_stack) -> None:
        _emit(window, _PROJECTION_TEMPLATE)
        from docwen_gui.widgets.template_selector_tabbed import TabbedTemplateSelector

        current = right_stack.currentWidget()
        assert isinstance(current, TabbedTemplateSelector), (
            f"Expected TabbedTemplateSelector, got {type(current).__name__}"
        )

    def test_conversion_slot_shows_conversion_panel(self, window, right_stack) -> None:
        _emit(window, _PROJECTION_CONVERSION)
        from docwen_gui.widgets.conversion_panel import ConversionPanel

        current = right_stack.currentWidget()
        assert isinstance(current, ConversionPanel), f"Expected ConversionPanel, got {type(current).__name__}"

    def test_switching_slots_updates_stack(self, window, right_stack) -> None:
        _emit(window, _PROJECTION_TEMPLATE)
        _emit(window, _PROJECTION_CONVERSION)
        _emit(window, _PROJECTION_TEMPLATE)
        from docwen_gui.widgets.template_selector_tabbed import TabbedTemplateSelector

        current = right_stack.currentWidget()
        assert isinstance(current, TabbedTemplateSelector)

    def test_switching_files_rederives_text_template_and_document_actions(self, window, right_stack) -> None:
        txt_ref = _file_ref("/tmp/plain.md", "markdown", "txt")
        docx_ref = _file_ref("/tmp/report.docx", "document", "docx")
        markdown_ref = _file_ref("/tmp/readme.txt", "markdown", "markdown")

        window._view_model.set_selected_file(txt_ref)
        assert window._view_model.ui_projection.right_panel_slot == RightPanelSlot.TEMPLATE
        assert window._action_area_vm.file_type == "docx"
        assert window._action_area_vm.file_path == txt_ref.path

        window._view_model.set_selected_file(docx_ref)
        assert window._view_model.ui_projection.right_panel_slot == RightPanelSlot.CONVERSION
        assert window._action_area_vm.file_type == "document"
        assert window._action_area_vm.file_path == docx_ref.path

        window._view_model.set_selected_file(markdown_ref)
        assert window._view_model.ui_projection.right_panel_slot == RightPanelSlot.TEMPLATE
        assert window._action_area_vm.file_type == "docx"
        assert window._action_area_vm.file_path == markdown_ref.path
        assert right_stack.currentWidget() is window._template_selector


class TestConversionPanelPdfInfoBinding:
    def test_pdf_page_reader_uses_content_for_unknown_suffix(self, tmp_path: Path) -> None:
        """The GUI page summary consumes the admitted PDF parser, not the name."""
        import fitz

        from docwen_gui.main_window import _read_pdf_total_pages

        source = tmp_path / "actual-pdf-content.unknown"
        document = fitz.open()
        try:
            document.new_page()
            document.new_page()
            document.save(source)
        finally:
            document.close()

        assert _read_pdf_total_pages(str(source)) == 2

    def test_layout_pdf_projection_syncs_page_count_and_file_name(self, window, monkeypatch, tmp_path) -> None:
        import docwen_gui.main_window as main_window

        source = tmp_path / "annual-report.pdf"
        source.write_bytes(b"%PDF-1.4\n")
        monkeypatch.setattr(main_window, "_read_pdf_total_pages", lambda file_path: 14)

        projection = MainWindowUiProjection(
            left_panel_visible=False,
            right_panel_visible=True,
            right_panel_slot=RightPanelSlot.CONVERSION,
            center_action_visible=True,
            info_area_visible=True,
            conversion_context=ConversionContext(
                category="layout",
                current_format="pdf",
                file_path=str(source),
            ),
            template_context=None,
        )
        _emit(window, projection)

        assert window._conversion_panel_vm.pdf_total_pages == 14
        assert window._conversion_panel_vm.pdf_file_name == "annual-report.pdf"

    def test_layout_pdf_projection_keeps_default_pdf_info_when_unreadable(self, window, monkeypatch, tmp_path) -> None:
        import docwen_gui.main_window as main_window

        source = tmp_path / "broken.pdf"
        source.write_bytes(b"not a pdf")
        monkeypatch.setattr(main_window, "_read_pdf_total_pages", lambda file_path: None)

        projection = MainWindowUiProjection(
            left_panel_visible=False,
            right_panel_visible=True,
            right_panel_slot=RightPanelSlot.CONVERSION,
            center_action_visible=True,
            info_area_visible=True,
            conversion_context=ConversionContext(
                category="layout",
                current_format="pdf",
                file_path=str(source),
            ),
            template_context=None,
        )
        _emit(window, projection)

        assert window._conversion_panel_vm.pdf_total_pages == 0
        assert window._conversion_panel_vm.pdf_file_name == ""

    def test_switching_to_unreadable_pdf_clears_previous_metadata_in_place(
        self,
        window,
        monkeypatch,
        tmp_path,
    ) -> None:
        import docwen_gui.main_window as main_window

        readable = tmp_path / "readable-a.pdf"
        unreadable = tmp_path / "unreadable-b.pdf"
        readable.write_bytes(b"%PDF-1.4\n")
        unreadable.write_bytes(b"not a pdf")
        monkeypatch.setattr(
            main_window,
            "_read_pdf_total_pages",
            lambda file_path: 12 if Path(file_path).name == readable.name else None,
        )

        def projection_for(source: Path) -> MainWindowUiProjection:
            return MainWindowUiProjection(
                left_panel_visible=False,
                right_panel_visible=True,
                right_panel_slot=RightPanelSlot.CONVERSION,
                center_action_visible=True,
                info_area_visible=True,
                conversion_context=ConversionContext(
                    category="layout",
                    current_format="pdf",
                    file_path=str(source),
                ),
                template_context=None,
            )

        _emit(window, projection_for(readable))
        page_edit = window._conversion_panel._page_input_edit
        info_label = window._conversion_panel._pdf_info_label
        split_button = window._conversion_panel._split_pdf_button
        assert page_edit is not None
        assert info_label is not None
        assert split_button is not None
        page_edit.setText("1-3")
        assert window._conversion_panel_vm.page_input == "1-3"
        assert window._conversion_panel_vm.pdf_total_pages == 12

        _emit(window, projection_for(unreadable))

        assert window._conversion_panel._page_input_edit is page_edit
        assert window._conversion_panel._pdf_info_label is info_label
        assert window._conversion_panel._split_pdf_button is split_button
        assert window._conversion_panel_vm.page_input == ""
        assert window._conversion_panel_vm.pdf_total_pages == 0
        assert window._conversion_panel_vm.pdf_file_name == ""
        assert page_edit.text() == ""
        assert info_label.isHidden()
        assert split_button.isEnabled() is False


class TestTemplateCatalogBinding:
    def test_template_refresh_does_not_project_inactive_default(self, window):
        window._view_model.set_selected_file(FileRef(path="C:/test/source.md", format="markdown", category="markdown"))
        selector = window._template_selector
        selector.restore_current_tab("docx")
        selector.get_selector("xlsx")._list.setCurrentRow(-1)
        window._show_template_target_mode("docx", "C:/test/source.md")
        window._action_area_vm.md_add_numbering = True
        window._render_template_catalog()
        assert selector.current_tab == "docx"
        assert window._action_area_vm.file_type == "docx"
        assert window._action_area_vm.md_add_numbering is True

    @pytest.mark.parametrize(
        ("template_type", "expected_mode"),
        [("docx", "docx"), ("xlsx", "md_to_spreadsheet")],
    )
    def test_template_selection_projects_matching_center_action(
        self,
        window,
        template_type: str,
        expected_mode: str,
    ) -> None:
        window._view_model.set_selected_file(FileRef(path="C:/test/source.md", format="markdown", category="markdown"))

        window._on_main_window_template_selected(template_type, "Template")

        assert window._action_area_vm.file_type == expected_mode
        assert window._action_area_vm.file_path == "C:/test/source.md"

    def test_template_tab_change_projects_spreadsheet_mode_before_selection(self, window) -> None:
        window._view_model.set_selected_file(FileRef(path="C:/test/source.md", format="markdown", category="markdown"))

        window._on_main_window_template_tab_changed("xlsx", "docx")

        assert window._action_area_vm.file_type == "md_to_spreadsheet"

    def test_template_selection_after_tab_change_does_not_reset_center_options(self, window, monkeypatch) -> None:
        window._view_model.set_selected_file(FileRef(path="C:/test/source.md", format="markdown", category="markdown"))
        setup_calls = 0
        original_setup = window._action_area_vm.setup_for_md_to_spreadsheet

        def counted_setup(file_path: str) -> None:
            nonlocal setup_calls
            setup_calls += 1
            original_setup(file_path)

        monkeypatch.setattr(window._action_area_vm, "setup_for_md_to_spreadsheet", counted_setup)

        # This is the real callback sequence for a tab with no previous choice:
        # the tab projects its mode, then automatic template selection reports
        # the chosen template.  The second callback must not rebuild defaults.
        window._on_main_window_template_tab_changed("xlsx", "docx")
        window._action_area_vm.target_format = "csv"
        window._on_main_window_template_selected("xlsx", "Budget")

        assert setup_calls == 1
        assert window._action_area_vm.target_format == "csv"

    def test_selecting_another_template_in_same_tab_preserves_user_options(self, window) -> None:
        window._view_model.set_selected_file(FileRef(path="C:/test/source.md", format="markdown", category="markdown"))
        window._action_area_vm.md_add_numbering = True
        window._action_area_vm.set_proofread_option("sensitive_word", True)

        window._on_main_window_template_selected("docx", "Another template")

        assert window._action_area_vm.md_add_numbering is True
        assert window._action_area_vm.proofread_options["sensitive_word"] is True

    def test_main_template_selector_restores_configured_default_type(self, qapp) -> None:
        window = _make_window_with_config(qapp, {"gui.template.md_default_template": "xlsx"})
        try:
            selector = window._template_selector
            assert selector is not None
            assert selector.current_tab == "xlsx"
        finally:
            window.close()

    @pytest.mark.parametrize("configured", ["", "unknown", None])
    def test_main_template_selector_invalid_default_falls_back_to_docx(self, qapp, configured) -> None:
        window = _make_window_with_config(qapp, {"gui.template.md_default_template": configured})
        try:
            selector = window._template_selector
            assert selector is not None
            assert selector.current_tab == "docx"
            assert window._main_template_default_type == "docx"
        finally:
            window.close()

    def test_runtime_settings_apply_restores_changed_template_default(self, qapp) -> None:
        values: dict[str, object] = {"gui.template.md_default_template": "docx"}
        window = _make_window_with_config(qapp, values)
        try:
            window._view_model.set_selected_file(
                FileRef(path="C:/test/source.md", format="markdown", category="markdown")
            )
            values["gui.template.md_default_template"] = "xlsx"

            window._apply_runtime_window_settings()

            selector = window._template_selector
            assert selector is not None
            assert selector.current_tab == "xlsx"
            assert window._action_area_vm.file_type == "md_to_spreadsheet"
        finally:
            window.close()

    def test_unrelated_settings_apply_does_not_override_manual_template_tab(self, qapp) -> None:
        values: dict[str, object] = {"gui.template.md_default_template": "docx"}
        window = _make_window_with_config(qapp, values)
        try:
            selector = window._template_selector
            assert selector is not None
            selector.restore_current_tab("xlsx")

            window._apply_runtime_window_settings()

            assert selector.current_tab == "xlsx"
        finally:
            window.close()

    def test_main_window_loads_template_lists_and_details_from_registry(self, window, tmp_path):
        from docx import Document

        from docwen_runtime.templates import TemplateRegistry

        templates_dir = tmp_path / "templates"
        templates_dir.mkdir()
        Document().save(str(templates_dir / "Report.docx"))
        registry = TemplateRegistry(templates_dir)
        window._template_vm.manager.registry = registry
        window._load_templates_into_main_selector()
        template = registry.list_templates()[0]
        selector = window._template_selector.get_selector("docx")
        assert selector._list.count() == 1
        selector.select_template(template.id, selection_source="user")
        assert "Report" in selector._details_label.text()
        assert selector.get_selected_resource_id() == template.id

    def test_template_management_button_focuses_current_template(self, window):
        selector = window._template_selector.get_selector("docx")
        selected = selector.get_selected_resource_id()
        selector._manage_button.click()
        dialog = window._settings_dialog
        assert dialog.current_section() == "templates"
        assert dialog._tabs["templates"]._selected_id() == selected
        dialog.close()

    def test_template_empty_state_button_opens_management(self, window):
        selector = window._template_selector.get_selector("xlsx")
        selector.clear_all()
        selector._empty_manage_button.click()
        dialog = window._settings_dialog
        assert dialog.current_section() == "templates"
        assert dialog._tabs["templates"]._target() == "xlsx"
        dialog.close()
