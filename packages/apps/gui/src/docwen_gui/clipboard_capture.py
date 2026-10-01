"""Owner-thread clipboard snapshot used by all non-file paste paths."""

from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtCore import QMimeData

from docwen_gui.clipboard_image_bytes import ClipboardImageBytesError, FrozenPng, freeze_qimage
from docwen_gui.clipboard_office_provider import (
    WORD_EMBED_SOURCE_MIME,
    WPS_DOCUMENT_MIME,
    WPS_IMAGE_DATA_MIME,
)

_CAPTURE_FORMAT_LIMITS = {
    "text/html": 8 * 1024 * 1024,
    WORD_EMBED_SOURCE_MIME: 32 * 1024 * 1024,
    WPS_DOCUMENT_MIME: 32 * 1024 * 1024,
    WPS_IMAGE_DATA_MIME: 32 * 1024 * 1024,
}
_CAPTURE_TOTAL_LIMIT = 64 * 1024 * 1024


@dataclass(frozen=True, slots=True)
class FrozenClipboardCapture:
    """One immutable owner-thread snapshot; downstream code never rereads Qt."""

    plain_text: str | None
    html_bytes: bytes
    binary_formats: tuple[tuple[str, bytes], ...]
    image: FrozenPng | None
    image_error_code: str = ""
    rich_error_code: str = ""

    def get(self, mime_type: str) -> bytes | None:
        return next((payload for key, payload in self.binary_formats if key == mime_type), None)


def freeze_clipboard_mime(mime_data: QMimeData) -> FrozenClipboardCapture:
    """Freeze allowlisted clipboard values exactly once on the Qt owner thread."""

    has_text = bool(mime_data.hasText())
    plain_text = str(mime_data.text()) if has_text else None
    formats = {str(value) for value in mime_data.formats()}
    frozen: list[tuple[str, bytes]] = []
    selected_formats = set(_CAPTURE_FORMAT_LIMITS).intersection(formats)
    if WPS_DOCUMENT_MIME in formats or WPS_IMAGE_DATA_MIME in formats:
        # A WPS-specific format is stronger identity than the generic OLE name,
        # including incomplete WPS captures, which must not be guessed as Word.
        selected_formats.discard(WORD_EMBED_SOURCE_MIME)
    total_bytes = 0
    rich_error_code = ""
    for mime_type in sorted(selected_formats):
        array = mime_data.data(mime_type)
        size = int(array.size())
        if size < 0 or size > _CAPTURE_FORMAT_LIMITS[mime_type] or total_bytes + size > _CAPTURE_TOTAL_LIMIT:
            rich_error_code = "clipboard.capture_budget_exceeded"
            frozen.clear()
            break
        payload = bytes(array.data())
        if len(payload) != size:
            rich_error_code = "clipboard.capture_invalid"
            frozen.clear()
            break
        total_bytes += size
        frozen.append((mime_type, payload))

    html_bytes = next((payload for key, payload in frozen if key == "text/html"), b"")
    image: FrozenPng | None = None
    image_error_code = ""
    # Office exports a bitmap preview alongside text/rich formats. Only a
    # standalone bitmap has authority to become a separate image input.
    if not has_text and not selected_formats and bool(mime_data.hasImage()):
        try:
            image = freeze_qimage(mime_data.imageData())
        except ClipboardImageBytesError as exc:
            image_error_code = exc.code
        except Exception:
            image_error_code = "clipboard.image_qt_invalid"
    return FrozenClipboardCapture(
        plain_text=plain_text,
        html_bytes=html_bytes,
        binary_formats=tuple(frozen),
        image=image,
        image_error_code=image_error_code,
        rich_error_code=rich_error_code,
    )


__all__ = ["FrozenClipboardCapture", "freeze_clipboard_mime"]
