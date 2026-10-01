"""GUI clipboard image ingress through real QMimeData and MainWindow paths."""

from __future__ import annotations

import base64
import hashlib
import io
from pathlib import Path

import pytest
from PIL import Image
from PySide6.QtCore import QMimeData, Qt
from PySide6.QtGui import QColor, QImage
from PySide6.QtWidgets import QApplication

from docwen_application.controller import ApplicationController
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_core.models.clipboard_document import (
    ClipboardHardBreak,
    ClipboardImageRef,
    ClipboardText,
    iter_clipboard_inlines,
    load_clipboard_document_bytes,
)
from docwen_gui.clipboard_image_bytes import freeze_qimage
from docwen_gui.main_window import MainWindow
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.gui


class _CountingMimeData(QMimeData):
    def __init__(self) -> None:
        super().__init__()
        self.calls = {"text": 0, "formats": 0, "data": 0, "has_image": 0, "image_data": 0}

    def reset_calls(self) -> None:
        for key in self.calls:
            self.calls[key] = 0

    def text(self) -> str:
        self.calls["text"] += 1
        return super().text()

    def formats(self) -> list[str]:
        self.calls["formats"] += 1
        return super().formats()

    def data(self, mime_type: str):
        self.calls["data"] += 1
        return super().data(mime_type)

    def hasImage(self) -> bool:
        self.calls["has_image"] += 1
        return super().hasImage()

    def imageData(self):
        self.calls["image_data"] += 1
        return super().imageData()


@pytest.fixture
def window(qapp: QApplication, qtbot, tmp_path: Path):
    controller = ApplicationController(runtime_port=create_runtime_port())
    controller.start()
    value = MainWindow(
        view_model=MainWindowViewModel(controller=controller),
        clipboard_input_root=tmp_path / "clipboard-inputs",
    )
    qtbot.addWidget(value)
    value.resize(720, 620)
    value.show()
    qapp.processEvents()
    yield value
    value.close()
    controller.stop()
    qapp.clipboard().clear()
    qapp.processEvents()


def _selected(window: MainWindow, qtbot):
    qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: window.view_model.selected_file is not None)
    selected = window.view_model.selected_file
    assert selected is not None
    return selected


def _png_data_uri(rgba: tuple[int, int, int, int]) -> tuple[bytes, str]:
    stream = io.BytesIO()
    Image.new("RGBA", (3, 2), rgba).save(stream, format="PNG")
    payload = stream.getvalue()
    return payload, "data:image/png;base64," + base64.b64encode(payload).decode("ascii")


def _authored_plain(document) -> str:
    parts: list[str] = []
    for inline in iter_clipboard_inlines(document):
        if isinstance(inline, ClipboardText):
            parts.append(inline.value)
        elif isinstance(inline, ClipboardHardBreak):
            parts.append("\n")
    return "".join(parts)


def test_busy_paste_does_not_touch_clipboard_formats(window: MainWindow, qapp: QApplication) -> None:
    mime = _CountingMimeData()
    mime.setText("must not be read")
    mime.setHtml("<p>must not be read</p>")
    qapp.clipboard().setMimeData(mime)
    mime.reset_calls()
    window._action_area_vm.show_cancel()

    window._on_paste_requested()

    assert mime.calls == {"text": 0, "formats": 0, "data": 0, "has_image": 0, "image_data": 0}


def test_plain_only_reads_text_once_without_touching_rich_or_bitmap(
    window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    raw = "  # exact plain\nhttps://example.test/a  "
    mime = _CountingMimeData()
    mime.setText(raw)
    mime.setHtml("<p>different HTML</p>")
    preview = QImage(4, 4, QImage.Format.Format_RGBA8888)
    preview.fill(QColor(1, 2, 3, 255))
    mime.setImageData(preview)
    qapp.clipboard().setMimeData(mime)
    mime.reset_calls()

    window._on_paste_requested(plain_text_only=True)

    selected = _selected(window, qtbot)
    assert Path(selected.path).read_text(encoding="utf-8") == raw
    assert mime.calls["text"] == 1
    assert mime.calls["formats"] == 0
    assert mime.calls["data"] == 0
    assert mime.calls["has_image"] == 0
    assert mime.calls["image_data"] == 0


def test_default_plain_path_reads_qmime_text_once(window: MainWindow, qapp: QApplication, qtbot) -> None:
    raw = "plain clipboard body"
    mime = _CountingMimeData()
    mime.setText(raw)
    qapp.clipboard().setMimeData(mime)
    mime.reset_calls()

    window._on_paste_requested()

    selected = _selected(window, qtbot)
    assert Path(selected.path).read_text(encoding="utf-8") == raw
    assert mime.calls["text"] == 1


def test_invalid_rich_with_bitmap_keeps_complete_frozen_plain(
    window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    raw = "  keep https://example.test/x **markdown**\nTAIL\u00a0 "
    mime = _CountingMimeData()
    mime.setText(raw)
    mime.setData("text/html", b"Version:1.0\r\nStartHTML:0000000010\r\n")
    preview = QImage(96, 64, QImage.Format.Format_RGBA8888)
    preview.fill(QColor(200, 10, 20, 255))
    mime.setImageData(preview)
    qapp.clipboard().setMimeData(mime)
    mime.reset_calls()

    window._on_paste_requested()

    selected = _selected(window, qtbot)
    assert selected.format == "markdown"
    assert Path(selected.path).read_text(encoding="utf-8") == raw
    assert mime.calls["text"] == 1
    assert mime.calls["image_data"] == 0
    assert any(row.message_type == "warning" for row in window._info_area_vm.history_rows)


def test_table_html_bitmap_preview_creates_no_document_image(
    window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    mime = _CountingMimeData()
    mime.setText("A\tB\n1\t2")
    mime.setHtml("<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>")
    preview = QImage(320, 180, QImage.Format.Format_RGBA8888)
    preview.fill(QColor(10, 20, 30, 255))
    mime.setImageData(preview)
    qapp.clipboard().setMimeData(mime)
    mime.reset_calls()

    window._on_paste_requested()

    selected = _selected(window, qtbot)
    assert selected.format == "clipboard_document"
    document = load_clipboard_document_bytes(Path(selected.path).read_bytes())
    assert document.resources == ()
    assert not any(isinstance(item, ClipboardImageRef) for item in iter_clipboard_inlines(document))
    assert mime.calls["image_data"] == 0


def test_standalone_transparent_dpr2_image_preserves_physical_pixels_and_store_lifecycle(
    window: MainWindow,
    qapp: QApplication,
    qtbot,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    image = QImage(96, 64, QImage.Format.Format_RGBA8888)
    image.fill(QColor(12, 34, 56, 0))
    image.setDevicePixelRatio(2.0)
    frozen = freeze_qimage(image)
    assert (frozen.width, frozen.height, frozen.device_pixel_ratio) == (96, 64, 2.0)
    mime = QMimeData()
    mime.setImageData(image)
    qapp.clipboard().setMimeData(mime)

    observed: list[tuple[int, int, float]] = []
    original = window._paste_standalone_clipboard_image

    def record_and_publish(frozen) -> None:
        observed.append((frozen.width, frozen.height, frozen.device_pixel_ratio))
        original(frozen)

    monkeypatch.setattr(window, "_paste_standalone_clipboard_image", record_and_publish)
    window._on_paste_requested()
    assert observed == [(96, 64, 2.0)], [(row.message_type, row.message) for row in window._info_area_vm.history_rows]
    first = _selected(window, qtbot)
    first_path = Path(first.path)
    assert observed == [(96, 64, 2.0)]
    with Image.open(first_path) as decoded:
        assert decoded.size == (96, 64)
        pixel = decoded.convert("RGBA").getpixel((0, 0))
        assert isinstance(pixel, tuple)
        assert pixel[3] == 0

    window._on_paste_requested()
    second = _selected(window, qtbot)
    second_path = Path(second.path)
    assert second_path != first_path
    qtbot.waitUntil(lambda: not first_path.exists())
    assert observed == [(96, 64, 2.0), (96, 64, 2.0)]

    qtbot.mouseClick(window.input_area.clear_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: window.view_model.files == [])
    qtbot.waitUntil(lambda: not second_path.exists())


def test_structured_store_failure_is_danger_not_plain_fallback(
    window: MainWindow,
    qapp: QApplication,
    qtbot,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    qapp.clipboard().setText("original")
    window._on_paste_requested()
    original = _selected(window, qtbot)
    original_path = Path(original.path)

    def unavailable_store():
        raise PermissionError("test-only structured store failure")

    monkeypatch.setattr(window, "_clipboard_store_for_paste", unavailable_store)
    mime = QMimeData()
    mime.setText("A\tB\n1\t2")
    mime.setHtml("<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>")
    qapp.clipboard().setMimeData(mime)
    warnings_before = sum(row.message_type == "warning" for row in window._info_area_vm.history_rows)

    window._on_paste_requested()
    qapp.processEvents()

    selected = window.view_model.selected_file
    assert selected is not None and selected.path == str(original_path)
    assert original_path.is_file()
    assert any(row.message_type == "danger" for row in window._info_area_vm.history_rows)
    assert sum(row.message_type == "warning" for row in window._info_area_vm.history_rows) == warnings_before


def test_plain_data_uri_images_bind_by_context_without_losing_authored_plain(
    window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    first_payload, first_uri = _png_data_uri((255, 0, 0, 128))
    second_payload, second_uri = _png_data_uri((0, 0, 255, 255))
    plain = "  before https://example.test/a **md** middle after  "
    mime = QMimeData()
    mime.setText(plain)
    mime.setHtml(
        f"<p>before<img src='{first_uri}' alt='alpha'> https://example.test/a **md** middle"
        f"<img src='{second_uri}' alt='blue'> after</p>"
    )
    qapp.clipboard().setMimeData(mime)

    window._on_paste_requested()

    selected = _selected(window, qtbot)
    assert selected.format == "clipboard_document"
    document = load_clipboard_document_bytes(Path(selected.path).read_bytes())
    assert _authored_plain(document) == plain
    images = [item for item in iter_clipboard_inlines(document) if isinstance(item, ClipboardImageRef)]
    assert len(images) == 2
    resource_sha = {item.resource_id: item.sha256 for item in document.resources}
    assert [resource_sha[item.resource_id] for item in images if item.resource_id is not None] == [
        hashlib.sha256(first_payload).hexdigest(),
        hashlib.sha256(second_payload).hexdigest(),
    ]


def test_html_only_data_uri_keeps_inline_order_and_warns(window: MainWindow, qapp: QApplication, qtbot) -> None:
    _payload, uri = _png_data_uri((1, 2, 3, 255))
    mime = QMimeData()
    mime.setHtml(f"<p>before<img src='{uri}' alt='inline'>after</p>")
    qapp.clipboard().setMimeData(mime)

    window._on_paste_requested()

    selected = _selected(window, qtbot)
    assert selected.format == "clipboard_document"
    document = load_clipboard_document_bytes(Path(selected.path).read_bytes())
    inlines = list(iter_clipboard_inlines(document))
    assert [type(item).__name__ for item in inlines] == ["ClipboardText", "ClipboardImageRef", "ClipboardText"]
    assert any(row.message_type == "warning" for row in window._info_area_vm.history_rows)


def test_external_image_keeps_alt_missing_diagnostic_and_plain_without_resource(
    window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    plain = "PREPOST"
    mime = QMimeData()
    mime.setText(plain)
    mime.setHtml(
        "<html><head><base href='https://example.invalid/base/'></head><body>"
        "<p>PRE<img src='file:///C:/private/image.png' alt='kept-alt'>POST</p></body></html>"
    )
    qapp.clipboard().setMimeData(mime)

    window._on_paste_requested()

    selected = _selected(window, qtbot)
    assert selected.format == "clipboard_document"
    document = load_clipboard_document_bytes(Path(selected.path).read_bytes())
    assert document.resources == ()
    image_refs = [item for item in iter_clipboard_inlines(document) if isinstance(item, ClipboardImageRef)]
    assert len(image_refs) == 1
    assert image_refs[0].resource_id is None
    assert image_refs[0].alt == "kept-alt"
    assert image_refs[0].missing_reason == "clipboard_resource_unavailable"
    assert _authored_plain(document) == plain
    assert any(row.message_type == "warning" for row in window._info_area_vm.history_rows)


def test_ambiguous_image_alignment_falls_back_to_complete_plain(window: MainWindow, qapp: QApplication, qtbot) -> None:
    _payload, uri = _png_data_uri((20, 30, 40, 255))
    plain = "same same same"
    mime = QMimeData()
    mime.setText(plain)
    mime.setHtml(f"<p>same <img src='{uri}' alt='ambiguous'>same</p>")
    qapp.clipboard().setMimeData(mime)

    window._on_paste_requested()

    selected = _selected(window, qtbot)
    assert selected.format == "markdown"
    assert Path(selected.path).read_text(encoding="utf-8") == plain
    assert any(row.message_type == "warning" for row in window._info_area_vm.history_rows)
