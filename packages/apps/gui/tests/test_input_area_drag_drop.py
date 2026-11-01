"""InputArea MIME drag/drop and complete file identity regressions."""

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, QPoint, QPointF, Qt, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent

from docwen_gui.view_models.input_area_vm import _BATCH_SCAN_LIMIT, InputAreaViewModel
from docwen_gui.widgets.input_area import (
    InputArea,
)

from ._input_area_widget_support import (
    input_vm as input_vm,
)
from ._input_area_widget_support import (
    main_vm as main_vm,
)
from ._input_area_widget_support import (
    widget as widget,
)

pytestmark = pytest.mark.gui


class TestDragDropMime:
    def test_accepts_url_mime(self, widget: InputArea) -> None:
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile("/tmp/test.docx")])
        assert mime.hasUrls()

    def test_accepts_text_mime(self) -> None:
        mime = QMimeData()
        mime.setText("/tmp/test.docx")
        assert mime.hasText()

    def test_drag_enter_previews_batch_folder_scan(self, widget: InputArea, tmp_path) -> None:
        folder = tmp_path / "input"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        supported = folder / "keep.txt"
        unsupported = nested / "ignore.bin"
        supported.write_text("1")
        unsupported.write_text("2")
        widget.view_model.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        widget.view_model.set_mode("batch")

        mime = QMimeData()
        mime.setText(str(folder))
        event = QDragEnterEvent(
            QPoint(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        widget.dragEnterEvent(event)

        assert event.isAccepted()
        assert "1" in widget._selection_label.text()
        assert "ignore.bin" in widget._selection_label.toolTip()
        assert widget._feedback_frame.property("feedbackTone") == "info"

    def test_drop_batch_folder_reports_partial_skips(self, widget: InputArea, tmp_path, qtbot) -> None:
        folder = tmp_path / "mixed"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        supported = folder / "keep.txt"
        unsupported = nested / "ignore.bin"
        supported.write_text("1")
        unsupported.write_text("2")
        widget.view_model.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        widget.view_model.set_mode("batch")
        added: list[list[str]] = []
        widget.view_model.files_added.connect(lambda paths: added.append(list(paths)))

        mime = QMimeData()
        mime.setText(str(folder))
        event = QDropEvent(
            QPointF(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        widget.dropEvent(event)
        qtbot.waitUntil(lambda: bool(added))

        assert event.isAccepted()
        assert added == [[str(supported)]]
        assert "1" in widget._selection_label.text()
        assert widget._feedback_frame.property("feedbackTone") == "warning"

    def test_url_drop_large_folder_summarizes_preview_but_adds_all_supported_files(
        self,
        widget: InputArea,
        tmp_path,
        qtbot,
    ) -> None:
        folder = tmp_path / "large-url-drop"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        total_supported = _BATCH_SCAN_LIMIT + 7
        supported_files = []
        for index in range(total_supported):
            target_dir = nested if index % 3 == 0 else folder
            file_path = target_dir / f"doc-{index:03}.txt"
            file_path.write_text("content")
            supported_files.append(str(file_path))
        (nested / "ignore.bin").write_text("unsupported")
        widget.view_model.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        widget.view_model.set_mode("batch")
        added: list[list[str]] = []
        widget.view_model.files_added.connect(lambda paths: added.append(list(paths)))

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(folder))])
        preview = widget.view_model.build_drag_preview([str(folder)])
        enter_event = QDragEnterEvent(
            QPoint(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        widget.dragEnterEvent(enter_event)

        assert enter_event.isAccepted()
        assert preview.has_degraded_preview is True
        assert widget._selection_label.text() == preview.message
        assert widget._feedback_frame.property("feedbackTone") == "info"

        drop_event = QDropEvent(
            QPointF(0, 0),
            Qt.DropAction.CopyAction,
            mime,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )

        widget.dropEvent(drop_event)
        qtbot.waitUntil(lambda: bool(added), timeout=10000)

        assert drop_event.isAccepted()
        assert len(added) == 1
        assert set(added[0]) == set(supported_files)
        assert len(added[0]) == total_supported
        assert "ignore.bin" not in "\n".join(added[0])
        assert widget._feedback_frame.property("feedbackTone") == "warning"


@pytest.mark.parametrize("mode", ["single", "batch"])
@pytest.mark.parametrize("mime_kind", ["bare", "braced", "quoted", "urls"])
def test_drop_keeps_complete_file_identity_with_an_existing_prefix(
    widget: InputArea, input_vm: InputAreaViewModel, tmp_path: Path, mode: str, mime_kind: str, qtbot
) -> None:
    prefix = tmp_path / "报告.md"
    prefix.write_text("# Prefix\n", encoding="utf-8")
    complete = tmp_path / "报告.md 2026.md"
    complete.write_text("# Complete\n", encoding="utf-8")
    input_vm.set_mode(mode)
    observed: list[list[str]] = []
    input_vm.files_added.connect(lambda paths: observed.append(list(paths)))
    mime = QMimeData()
    if mime_kind == "urls":
        mime.setUrls([QUrl.fromLocalFile(str(complete))])
        mime.setText(str(prefix))  # Native URL identity wins over text fallback.
    else:
        wrappers = {"bare": "{}", "braced": "{{{}}}", "quoted": '"{}"'}
        mime.setText(wrappers[mime_kind].format(complete))
    event = QDropEvent(
        QPointF(12, 12),
        Qt.DropAction.CopyAction,
        mime,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    widget.dropEvent(event)
    qtbot.waitUntil(lambda: bool(observed))
    assert observed == [[str(complete)]]
    assert event.isAccepted()


def test_text_drop_preserves_real_multiple_quoted_file_identities(widget, input_vm, tmp_path, qtbot) -> None:
    paths = [tmp_path / "first note.md", tmp_path / "second note.md"]
    for path in paths:
        path.write_text("# Input\n", encoding="utf-8")
    input_vm.set_mode("batch")
    observed = []
    input_vm.files_added.connect(lambda values: observed.append(list(values)))
    mime = QMimeData()
    mime.setText(" ".join(f'"{path}"' for path in paths))
    event = QDropEvent(
        QPointF(12, 12),
        Qt.DropAction.CopyAction,
        mime,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    widget.dropEvent(event)
    qtbot.waitUntil(lambda: bool(observed))
    assert observed == [[str(path) for path in paths]]
