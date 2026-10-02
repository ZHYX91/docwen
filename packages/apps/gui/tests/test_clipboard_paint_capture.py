"""Captured Paint bitmap must not be mistaken for an Office document preview."""

import hashlib
import struct
from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData
from PySide6.QtGui import QImage

from docwen_core.models.clipboard_document import ClipboardTable, clipboard_cell_text
from docwen_gui.clipboard_capture import freeze_clipboard_mime
from docwen_gui.clipboard_office_provider import WORD_EMBED_SOURCE_MIME
from docwen_gui.clipboard_paint_provider import paint_bitmap_dimensions
from docwen_gui.clipboard_rich_document import project_frozen_rich_document

pytestmark = pytest.mark.contract
_SAMPLES = Path(__file__).resolve().parents[4] / "tests/fixtures/files/clipboard-paint"


class _PreviewCountingMime(QMimeData):
    def __init__(self) -> None:
        super().__init__()
        self.image_reads = 0

    def imageData(self) -> object:
        self.image_reads += 1
        return super().imageData()


def test_known_paint_source_does_not_block_html_table_or_read_preview(qapp) -> None:
    mime = _PreviewCountingMime()
    mime.setData(WORD_EMBED_SOURCE_MIME, _paint_payload())
    mime.setImageData(QImage(str(_SAMPLES / "paint-qt-image.png")))
    mime.setHtml("<table><tr><th>Name</th><th>Value</th></tr><tr><td>00123</td><td>=1+1</td></tr></table>")

    capture = freeze_clipboard_mime(mime)
    decision = project_frozen_rich_document(capture)

    assert capture.image is None and mime.image_reads == 0
    assert decision.projection is not None and not decision.plain_fallback
    table = decision.projection.document.blocks[0]
    assert isinstance(table, ClipboardTable)
    assert [clipboard_cell_text(cell) for cell in table.cells] == ["Name", "Value", "00123", "=1+1"]
    assert decision.projection.resources == ()


def test_known_paint_source_does_not_trigger_plain_fallback_or_read_preview(qapp) -> None:
    mime = _PreviewCountingMime()
    mime.setData(WORD_EMBED_SOURCE_MIME, _paint_payload())
    mime.setImageData(QImage(str(_SAMPLES / "paint-qt-image.png")))
    mime.setText("preserve this body exactly")

    capture = freeze_clipboard_mime(mime)
    decision = project_frozen_rich_document(capture)

    assert capture.plain_text == "preserve this body exactly"
    assert capture.image is None and mime.image_reads == 0
    assert not decision.handled and not decision.plain_fallback


def _paint_payload() -> bytes:
    payload = (_SAMPLES / "paint-native.ole").read_bytes()
    assert hashlib.sha256(payload).hexdigest() == "ee74060b3182d0ba4e63788f7a16bb61f91c4462d7835af75e5bd6362e095515"
    return payload


def test_actual_paint_bitmap_capture_is_not_an_office_document(qapp) -> None:
    mime = QMimeData()
    mime.setData(WORD_EMBED_SOURCE_MIME, _paint_payload())
    mime.setImageData(QImage(str(_SAMPLES / "paint-qt-image.png")))

    capture = freeze_clipboard_mime(mime)

    assert capture.image is not None
    assert (capture.image.width, capture.image.height) == (96, 64)
    assert capture.image.rgba_sha256 == "313143bd124d48f05ab6219fe45a7678c5e24366432d33aa0045391f8abc5be9"
    assert not project_frozen_rich_document(capture).handled


@pytest.mark.parametrize(
    "defect", ["class", "ambiguous", "length", "padding", "dimensions", "compression", "truncated"]
)
def test_malformed_or_ambiguous_paint_does_not_authorize_a_bitmap(defect: str) -> None:
    payload = bytearray(_paint_payload())
    directory = 512 + 512 * struct.unpack_from("<I", payload, 48)[0]
    native = payload.index(struct.pack("<I", 24640) + b"BM")
    if defect == "class":
        payload[directory + 80] ^= 1
    elif defect == "ambiguous":
        entry = directory + 128
        name = "Package"
        label = (name + "\0").encode("utf-16le")
        payload[entry : entry + 64] = label.ljust(64, b"\0")
        struct.pack_into("<H", payload, entry + 64, len(label))
    elif defect == "length":
        struct.pack_into("<I", payload, native, 24639)
    elif defect == "padding":
        payload[native + 4 + 24630] = 1
    elif defect == "dimensions":
        struct.pack_into("<i", payload, native + 4 + 18, 20000)
    elif defect == "compression":
        struct.pack_into("<I", payload, native + 4 + 30, 1)
    else:
        payload = payload[: native + 100]
    assert paint_bitmap_dimensions(bytes(payload)) is None


@pytest.mark.parametrize("other", ["text", "html", "wps"])
def test_paint_like_generic_source_does_not_override_other_content(qapp, other: str) -> None:
    mime = QMimeData()
    mime.setData(WORD_EMBED_SOURCE_MIME, _paint_payload())
    mime.setImageData(QImage(str(_SAMPLES / "paint-qt-image.png")))
    if other == "text":
        mime.setText("preserve this body")
    elif other == "html":
        mime.setHtml("<p>document content</p>")
    else:
        mime.setData('application/x-qt-windows-mime;value="Kingsoft WPS 9.0 Format"', b"incomplete provider")
    capture = freeze_clipboard_mime(mime)
    assert capture.image is None
    if other == "text":
        assert capture.plain_text == "preserve this body"
        assert not project_frozen_rich_document(capture).handled


def test_paint_bitmap_dimensions_must_match_qimage(qapp) -> None:
    mime = QMimeData()
    mime.setData(WORD_EMBED_SOURCE_MIME, _paint_payload())
    image = QImage(3, 2, QImage.Format.Format_RGBA8888)
    image.fill(0)
    mime.setImageData(image)
    capture = freeze_clipboard_mime(mime)
    assert capture.image is None
    assert capture.image_error_code == "clipboard.image_provider_mismatch"


@pytest.mark.parametrize("name", ["word-derived.ole", "paint-native.ole"])
def test_unknown_or_office_embed_source_is_not_assumed_to_be_a_bitmap(qapp, name: str) -> None:
    payload = (
        (_SAMPLES.parent / "clipboard-office" / name).read_bytes() if name == "word-derived.ole" else b"unknown OLE"
    )
    mime = QMimeData()
    mime.setData(WORD_EMBED_SOURCE_MIME, payload)
    mime.setImageData(QImage(str(_SAMPLES / "paint-qt-image.png")))
    capture = freeze_clipboard_mime(mime)
    assert capture.image is None
    assert capture.get(WORD_EMBED_SOURCE_MIME) == payload
