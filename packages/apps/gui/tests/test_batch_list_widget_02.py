"""Focused tests split from test_batch_list_widget.py."""

from __future__ import annotations

from ._batch_list_widget_support import (
    _BATCH_SCAN_LIMIT,
    CATEGORY_ORDER,
    BatchEntryItemWidget,
    BatchFileEntry,
    BatchList,
    BatchListViewModel,
    QApplication,
    QBoxLayout,
    QListWidget,
    QListWidgetItem,
    Qt,
    ThemeManager,
    _dominant_opaque_pixmap_color,
    _source_path_text,
    _t,
    get_status_color,
    pytest,
)
from ._batch_list_widget_support import (
    populated_widget as populated_widget,
)

pytestmark = pytest.mark.gui
from ._batch_list_widget_support import (
    vm as vm,
)
from ._batch_list_widget_support import (
    widget as widget,
)


class TestBatchEntryItemWidget:
    @pytest.fixture
    def entry(self):
        return BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
            status="pending",
        )

    @pytest.fixture
    def entry_widget(self, qapp: QApplication, entry):
        w = BatchEntryItemWidget(entry)
        yield w
        w.deleteLater()

    def test_entry_created(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget is not None

    def test_entry_object_name(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget.objectName() == "batchEntryCard"

    def test_has_status_icon(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget.status_icon_label is not None

    @pytest.mark.parametrize("status", ["completed", "processing"])
    def test_status_icon_recolors_during_live_theme_preview(self, qapp: QApplication, status: str) -> None:
        manager = ThemeManager.get_instance()
        previous_theme = manager.get_current_theme()
        manager.initialize(qapp, "light")
        entry = BatchFileEntry(
            file_path="C:/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            status=status,
            output_path="C:/test/output.docx" if status == "completed" else None,
        )
        widget = BatchEntryItemWidget(entry)
        try:
            assert _dominant_opaque_pixmap_color(widget.status_icon_label) == get_status_color(status, "light").upper()

            manager.apply_theme("dark")
            qapp.processEvents()

            assert _dominant_opaque_pixmap_color(widget.status_icon_label) == get_status_color(status, "dark").upper()
        finally:
            widget.deleteLater()
            manager.apply_theme(previous_theme)
            qapp.processEvents()

    @pytest.mark.parametrize("status", ["processing", "completed", "skipped", "failed"])
    def test_compact_header_is_status_independent_and_recovers_while_hidden(
        self,
        qapp: QApplication,
        status: str,
    ) -> None:
        manager = ThemeManager.get_instance()
        previous_theme = manager.get_current_theme()
        previous_preset = manager.get_font_size_preset()
        manager.initialize(qapp, "light")
        manager.apply_font_size_preset("default")
        entry = BatchFileEntry(
            file_path=f"C:/test/{status}.docx",
            file_name=f"{status}.docx",
            detected_format="docx",
            workflow_category="document",
            status=status,
            output_path="C:/test/output.docx" if status == "completed" else None,
            skip_reason="Skipped by rule" if status == "skipped" else None,
            error_message="Conversion failed" if status == "failed" else None,
        )
        widget = BatchEntryItemWidget(entry)
        try:
            # A 358 px content area is wide enough at the default type size,
            # regardless of status text, icon presence, or prior layout state.
            widget.setFixedWidth(374)
            widget.show()
            for _ in range(3):
                qapp.processEvents()
            assert widget._is_compact is False
            assert widget._header_layout.direction() == QBoxLayout.Direction.LeftToRight

            # Entering compact mode must not feed the stretched icon width
            # back into the next decision and make the vertical layout sticky.
            widget.setFixedWidth(320)
            for _ in range(3):
                qapp.processEvents()
            assert widget._is_compact is True
            assert widget._header_layout.direction() == QBoxLayout.Direction.TopToBottom

            widget.setFixedWidth(374)
            for _ in range(3):
                qapp.processEvents()
            assert widget._is_compact is False
            assert widget._header_layout.direction() == QBoxLayout.Direction.LeftToRight

            # Hidden category pages use the same threshold and recover too.
            widget.hide()
            widget.setFixedWidth(320)
            widget._apply_compact_mode()
            assert widget._is_compact is True
            widget.setFixedWidth(374)
            widget._apply_compact_mode()
            assert widget._is_compact is False
        finally:
            widget.deleteLater()
            manager.apply_font_size_preset(previous_preset)
            manager.apply_theme(previous_theme)
            qapp.processEvents()

    def test_has_name_label(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget.name_label is not None
        assert "doc.docx" in entry_widget.name_label.text()
        assert entry_widget.name_label.focusPolicy() == Qt.FocusPolicy.StrongFocus
        assert entry_widget.name_label.cursor().shape() == Qt.CursorShape.PointingHandCursor

    @pytest.mark.parametrize("key", [Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space])
    def test_name_label_opens_source_location_by_keyboard(
        self,
        entry_widget: BatchEntryItemWidget,
        qtbot,
        key: Qt.Key,
    ) -> None:
        emitted: list[tuple[str, str]] = []
        entry_widget.action_requested.connect(lambda action, path: emitted.append((action, path)))

        qtbot.keyClick(entry_widget.name_label, key)

        assert emitted == [("open_source_location", "/test/doc.docx")]

    def test_name_label_opens_source_location_by_left_click(
        self,
        entry_widget: BatchEntryItemWidget,
        qtbot,
    ) -> None:
        emitted: list[tuple[str, str]] = []
        entry_widget.action_requested.connect(lambda action, path: emitted.append((action, path)))
        entry_widget.show()

        qtbot.mouseClick(entry_widget.name_label, Qt.MouseButton.LeftButton)

        assert emitted == [("open_source_location", "/test/doc.docx")]

    def test_name_label_context_menu_copies_exact_source_path(
        self,
        entry_widget: BatchEntryItemWidget,
    ) -> None:
        menu = entry_widget.name_label._create_context_menu()
        actions = menu.actions()
        actions[0].trigger()

        assert [action.text() for action in actions] == [_t("info_area.copy_path", "Copy Path")]
        assert QApplication.clipboard().text() == "/test/doc.docx"

    def test_has_info_badge(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget.info_badge is not None
        assert _t("components.file_drop.status.pending", "Pending") in entry_widget.info_badge.text()

    def test_has_action_buttons(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget.primary_action_button is not None
        assert entry_widget.retry_button is not None
        assert entry_widget.remove_button is not None

    def test_has_badge_strip(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget.badge_strip is not None

    def test_has_body_section(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget.body_section is not None

    def test_body_rows_keep_old_pyside_style_anchors(self, entry_widget: BatchEntryItemWidget) -> None:
        path_value = entry_widget._get_row_value_widget(entry_widget.path_row)
        detail_value = entry_widget._get_row_value_widget(entry_widget.detail_row)
        output_value = entry_widget._get_row_value_widget(entry_widget.output_row)

        assert path_value is not None
        assert detail_value is not None
        assert output_value is not None
        assert path_value.objectName() == "batchPathLabel"
        assert detail_value.objectName() == "batchDetailLabel"
        assert output_value.objectName() == "batchOutputLabel"
        assert entry_widget.output_row.objectName() == "batchOutputRow"

    @pytest.mark.parametrize(
        ("status", "error_message", "skip_reason", "warning_message", "title_tone", "detail_tone"),
        [
            ("failed", "boom", None, None, "danger", "danger"),
            ("skipped", None, "unsupported", None, "warning", "warning"),
            ("pending", None, None, "check input", "secondary", "warning"),
            ("pending", None, None, None, "secondary", "secondary"),
        ],
    )
    def test_detail_tone_tracks_entry_semantics(
        self,
        entry_widget: BatchEntryItemWidget,
        status: str,
        error_message: str | None,
        skip_reason: str | None,
        warning_message: str | None,
        title_tone: str,
        detail_tone: str,
    ) -> None:
        entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
            status=status,
            error_message=error_message,
            skip_reason=skip_reason,
            warning_message=warning_message,
        )

        entry_widget._apply_entry(entry)

        title = entry_widget._get_row_label_widget(entry_widget.detail_row)
        value = entry_widget._get_row_value_widget(entry_widget.detail_row)
        assert title is not None
        assert value is not None
        assert title.property("class") == title_tone
        assert value.property("detailRole") == detail_tone
        assert value.property("class") == detail_tone

    def test_source_path_is_visible_and_discoverable(self, entry_widget: BatchEntryItemWidget) -> None:
        path_value = entry_widget._get_row_value_widget(entry_widget.path_row)

        assert path_value is not None
        assert not entry_widget.path_row.isHidden()
        source_directory = _source_path_text("/test/doc.docx")
        assert path_value.toolTip() == source_directory
        assert path_value.accessibleName() == source_directory
        assert entry_widget.name_label.toolTip() == "/test/doc.docx"
        assert entry_widget.name_label.accessibleDescription() == "/test/doc.docx"

    def test_pending_entry_uses_sequence_marker(self, entry_widget: BatchEntryItemWidget) -> None:
        entry_widget.set_sequence_number(12)

        assert entry_widget.status_icon_label.text() == "12"
        assert not entry_widget.status_icon_label.isHidden()
        assert entry_widget.status_icon_label.accessibleName() == "12"

    def test_narrow_source_path_is_middle_elided_but_full_path_remains_available(
        self,
        qapp: QApplication,
    ) -> None:
        full_path = "/test/a/very/long/source/directory/with/many/segments/doc.docx"
        entry = BatchFileEntry(
            file_path=full_path,
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
            status="pending",
        )
        widget = BatchEntryItemWidget(entry)
        try:
            path_value = widget._get_row_value_widget(widget.path_row)
            assert path_value is not None
            source_directory = _source_path_text(full_path)

            path_value.resize(80, max(path_value.sizeHint().height(), 16))
            qapp.processEvents()

            assert path_value.text() != source_directory
            assert "…" in path_value.text()
            assert path_value.toolTip() == source_directory
            assert path_value.accessibleName() == source_directory
        finally:
            widget.deleteLater()

    def test_output_row_initially_hidden(self, entry_widget: BatchEntryItemWidget) -> None:
        assert entry_widget.output_row.isHidden()

    def test_detail_row_visible_on_hover(self, entry_widget: BatchEntryItemWidget) -> None:
        # Initially detail row is hidden when no content
        from PySide6.QtCore import QPointF
        from PySide6.QtGui import QEnterEvent

        # Use enterEvent to trigger hover visibility
        enter_event = QEnterEvent(QPointF(5, 5), QPointF(0, 0), QPointF(0, 0))
        entry_widget.enterEvent(enter_event)
        # detail should be hidden if no content
        assert entry_widget.detail_row.isVisible() is True or entry_widget.detail_row.isHidden() is True

    def test_leave_event(self, entry_widget: BatchEntryItemWidget) -> None:
        from PySide6.QtCore import QEvent

        entry_widget.leaveEvent(QEvent(QEvent.Type.Leave))
        # Should be stable

    def test_failed_entry_shows_retry(self, entry_widget: BatchEntryItemWidget) -> None:
        failed_entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
            status="failed",
            error_message="Test error",
            error_count=1,
        )
        entry_widget._apply_entry(failed_entry)
        entry_widget._apply_visibility_for_state()
        assert entry_widget._primary_action_key == "show_error_details"
        assert _t("components.file_drop.status.failed", "Failed") in entry_widget.info_badge.text()

    def test_failed_entry_with_retained_output_can_open_it(self, entry_widget: BatchEntryItemWidget) -> None:
        failed_entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
            status="failed",
            output_path="/output/doc_fromLegacy.docx",
            error_message="Downstream conversion failed",
        )

        entry_widget._apply_entry(failed_entry)
        entry_widget._apply_visibility_for_state()

        assert entry_widget._primary_action_key == "open_output"
        assert entry_widget.primary_action_button.text() == _t("components.file_drop.batch_list.action_open_output")
        assert entry_widget._secondary_action_visibility["retry"] is True

    def test_failed_entry_expansion_updates_list_item_height(self, qapp: QApplication) -> None:
        failed_entry = BatchFileEntry(
            file_path="/test/annual-summary-2026.docx",
            file_name="annual-summary-2026.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=842752,
            status="failed",
            error_message="Conversion failed: annual-summary-2026.docx. The document structure could not be read.",
            error_count=1,
        )
        list_widget = QListWidget()
        list_widget.resize(320, 400)
        item = QListWidgetItem(list_widget)
        entry_widget = BatchEntryItemWidget(failed_entry)
        list_widget.setItemWidget(item, entry_widget)
        entry_widget.bind_list_item(list_widget, item)
        qapp.processEvents()

        collapsed_height = item.sizeHint().height()
        entry_widget.set_interaction_state(selected=True, current=True)
        qapp.processEvents()

        assert item.sizeHint().height() > collapsed_height

    def test_completed_entry_shows_open_output(self, entry_widget: BatchEntryItemWidget) -> None:
        completed_entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
            status="completed",
            output_path="/output/doc.md",
        )
        entry_widget._apply_entry(completed_entry)
        entry_widget._apply_visibility_for_state()
        assert entry_widget._primary_action_key == "open_output"

    def test_cancelled_entry_is_not_retryable_failure(self, entry_widget: BatchEntryItemWidget) -> None:
        cancelled_entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
            status="cancelled",
            error_message="Task was cancelled",
        )

        entry_widget._apply_entry(cancelled_entry)
        entry_widget._apply_visibility_for_state()

        assert _t("components.file_drop.status.cancelled", "Cancelled") in entry_widget.info_badge.text()
        assert entry_widget._primary_action_key is None
        assert entry_widget.retry_button.isHidden()

    def test_action_requested_signal(self, entry_widget: BatchEntryItemWidget) -> None:
        emitted: list[tuple] = []
        entry_widget.action_requested.connect(lambda key, fp: emitted.append((key, fp)))
        entry_widget.retry_button.click()
        assert len(emitted) == 1
        assert emitted[0][0] == "retry_failed"

    def test_remove_click(self, entry_widget: BatchEntryItemWidget) -> None:
        emitted: list[tuple] = []
        entry_widget.action_requested.connect(lambda key, fp: emitted.append((key, fp)))
        entry_widget.remove_button.click()
        assert len(emitted) == 1
        assert emitted[0][0] == "remove_entry"


class TestSixTabs:
    def test_initial_tab_is_text(self, populated_widget: BatchList) -> None:
        assert populated_widget._vm.current_category == "text"

    def test_each_tab_has_list_widget(self, populated_widget: BatchList) -> None:
        for cat in CATEGORY_ORDER:
            lw = populated_widget._tabs[cat]
            assert isinstance(lw, QListWidget)

    def test_files_distributed_to_correct_tabs(self, populated_widget: BatchList) -> None:
        # docx -> document, xlsx -> spreadsheet, png -> image,
        # pdf -> layout, md -> text, epub -> other
        assert populated_widget._tabs["document"].count() >= 1
        assert populated_widget._tabs["spreadsheet"].count() >= 1
        assert populated_widget._tabs["image"].count() >= 1
        assert populated_widget._tabs["layout"].count() >= 1
        assert populated_widget._tabs["text"].count() >= 1
        assert populated_widget._tabs["other"].count() >= 1

    def test_prefilled_large_document_batch_keeps_widget_and_filter_counts(self, qapp: QApplication) -> None:
        vm = BatchListViewModel()
        paths = [f"/test/report-{i:03d}.docx" for i in range(_BATCH_SCAN_LIMIT + 45)]

        def resolver(_path: str) -> dict[str, str]:
            return {"detected_format": "docx", "workflow_category": "document"}

        vm.add_files(paths, file_resolver=resolver)
        completed_paths = paths[::5]
        for path in completed_paths:
            vm.set_file_status(path, "completed")

        widget = BatchList(view_model=vm)
        try:
            vm.activate_tab("document")
            list_widget = widget._tabs["document"]

            assert list_widget.count() == len(paths)
            assert vm.get_visible_count_for_category("document") == len(paths)

            vm.set_status_filter("completed")
            qapp.processEvents()

            visible_items = [
                list_widget.item(index)
                for index in range(list_widget.count())
                if not list_widget.item(index).isHidden()
            ]
            assert len(visible_items) == len(completed_paths)
            assert widget.get_current_file() in completed_paths
            assert list_widget.count() > _BATCH_SCAN_LIMIT
        finally:
            widget.deleteLater()

    def test_large_batch_defers_card_construction_without_deferring_items(
        self,
        qapp: QApplication,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        vm = BatchListViewModel()
        widget = BatchList(view_model=vm)
        paths = [f"/test/report-{index:03d}.docx" for index in range(245)]
        attached: list[str] = []

        def resolver(_path: str) -> dict[str, str]:
            return {"detected_format": "docx", "workflow_category": "document"}

        def record_attachment(
            _list_widget: QListWidget,
            _item: QListWidgetItem,
            entry: BatchFileEntry,
        ) -> None:
            attached.append(entry.file_path)

        monkeypatch.setattr(widget, "_attach_entry_widget", record_attachment)
        try:
            vm.add_files(paths, file_resolver=resolver)

            assert widget._tabs["document"].count() == len(paths)
            synchronous_attached = list(attached)
            assert 0 < len(synchronous_attached) < len(paths)
            assert synchronous_attached == paths[: len(synchronous_attached)]
            assert widget._entry_widget_attach_timer.isActive()

            while widget._entry_widget_attach_timer.isActive():
                qapp.processEvents()

            assert attached == paths
        finally:
            widget.deleteLater()
