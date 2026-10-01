"""Owner-thread clipboard snapshot used by all non-file paste paths."""

from __future__ import annotations

from dataclasses import dataclass

from docwen_gui.clipboard_image_bytes import ClipboardImageBytesError, FrozenPng, freeze_qimage
from docwen_gui.clipboard_office_provider import (
    WORD_EMBED_SOURCE_MIME,
    WPS_DOCUMENT_MIME,
    WPS_IMAGE_DATA_MIME,
)

_ALLOWED_BINARY_MIME = frozenset(
    {
        "text/html",
        WORD_EMBED_SOURCE_MIME,
        WPS_DOCUMENT_MIME,
        WPS_IMAGE_DATA_MIME,
    }
)


@dataclass(frozen=True, slots=True)
class FrozenClipboardCapture:
    """One immutable owner-thread snapshot; downstream code never rereads Qt."""

    plain_text: str | None
    html_bytes: bytes
    binary_formats: tuple[tuple[str, bytes], ...]
    image: FrozenPng | None
    image_error_code: str = ""

    def get(self, mime_type: str) -> bytes | None:
        return next((payload for key, payload in self.binary_formats if key == mime_type), None)


def freeze_clipboard_mime(mime_data: object) -> FrozenClipboardCapture:
    """Freeze allowlisted clipboard values exactly once on the Qt owner thread."""

    has_text = bool(getattr(mime_data, "hasText")())
    plain_text = str(getattr(mime_data, "text")()) if has_text else None
    formats = set(str(value) for value in getattr(mime_data, "formats")())
    frozen: list[tuple[str, bytes]] = []
    for mime_type in sorted(_ALLOWED_BINARY_MIME.intersection(formats)):
        payload = bytes(getattr(mime_data, "data")(mime_type))
        if payload:
            frozen.append((mime_type, payload))

    html_bytes = next((payload for key, payload in frozen if key == "text/html"), b"")
    image: FrozenPng | None = None
    image_error_code = ""
    if bool(getattr(mime_data, "hasImage")()):
        try:
            image = freeze_qimage(getattr(mime_data, "imageData")())
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
    )


__all__ = ["FrozenClipboardCapture", "freeze_clipboard_mime"]
