"""Project one frozen clipboard capture into a structured document when provable."""

from __future__ import annotations

from dataclasses import dataclass

from docwen_core.models.clipboard_document import (
    MAX_CLIPBOARD_TEXT_CODEPOINTS,
    ClipboardDocument,
    clipboard_document_to_bytes,
)
from docwen_gui.clipboard_capture import FrozenClipboardCapture
from docwen_gui.clipboard_image_binding import bind_provider_images
from docwen_gui.clipboard_office_provider import (
    WORD_EMBED_SOURCE_MIME,
    WPS_DOCUMENT_MIME,
    WPS_IMAGE_DATA_MIME,
    ClipboardOfficeProviderError,
    parse_embed_source_images,
    parse_wps_writer,
)
from docwen_gui.clipboard_structured import (
    StructuredClipboardProjection,
    html_contains_image,
    html_contains_table,
    project_structured_clipboard_html_with_resources,
    splice_plain_text_with_structured_tables,
)


class ClipboardRichDocumentError(ValueError):
    """Stable rich-document projection failure."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)


@dataclass(frozen=True, slots=True)
class RichDocumentProjection:
    document: ClipboardDocument
    resources: tuple[tuple[str, str, str, bytes], ...]
    html_only: bool
    provider: str = ""


@dataclass(frozen=True, slots=True)
class RichDocumentDecision:
    projection: RichDocumentProjection | None
    handled: bool
    plain_fallback: bool = False


def _project_frozen_rich_document(capture: FrozenClipboardCapture) -> RichDocumentDecision:
    """Use plain text as body authority and provider bytes only for proven images."""

    html_bytes = capture.html_bytes
    if capture.plain_text is not None and len(capture.plain_text) > MAX_CLIPBOARD_TEXT_CODEPOINTS:
        return RichDocumentDecision(None, True, plain_fallback=True)
    word_payload = capture.get(WORD_EMBED_SOURCE_MIME)
    wps_document = capture.get(WPS_DOCUMENT_MIME)
    wps_images = capture.get(WPS_IMAGE_DATA_MIME)
    has_provider = word_payload is not None or wps_document is not None or wps_images is not None
    if capture.rich_error_code:
        raise ClipboardRichDocumentError(capture.rich_error_code, "Clipboard rich capture could not be frozen safely.")
    has_wps = wps_document is not None or wps_images is not None
    if has_wps and (wps_document is None or wps_images is None):
        raise ClipboardRichDocumentError(
            "clipboard.provider_incomplete", "WPS clipboard image provider data is incomplete."
        )

    if not html_bytes:
        if has_provider:
            raise ClipboardRichDocumentError(
                "clipboard.provider_html_missing",
                "Rich clipboard provider data has no captured HTML structure.",
            )
        return RichDocumentDecision(None, False)

    try:
        contains_table = html_contains_table(html_bytes)
        contains_image = html_contains_image(html_bytes)
    except (UnicodeDecodeError, ValueError) as exc:
        raise ClipboardRichDocumentError(
            "clipboard.structured_invalid",
            "Clipboard rich document structure is invalid.",
        ) from exc
    if not contains_table and not contains_image and not has_provider:
        return RichDocumentDecision(None, False)

    try:
        provider_projection = None
        if wps_document is not None and wps_images is not None:
            provider_projection = parse_wps_writer(wps_document, wps_images)
        elif word_payload is not None:
            provider_projection = parse_embed_source_images(word_payload)

        html_projection: StructuredClipboardProjection = project_structured_clipboard_html_with_resources(
            html_bytes, decode_inline_images=provider_projection is None
        )
        document = html_projection.document

        resources = html_projection.resources
        provider_name = ""
        if provider_projection is not None:
            document, resources = bind_provider_images(document, provider_projection, resources)
            provider_name = provider_projection.provider

        if capture.plain_text is not None:
            merged = splice_plain_text_with_structured_tables(document, capture.plain_text)
            if merged is None:
                return RichDocumentDecision(None, True, plain_fallback=True)
            document = merged

        clipboard_document_to_bytes(document)
        return RichDocumentDecision(
            RichDocumentProjection(
                document=document,
                resources=resources,
                html_only=capture.plain_text is None,
                provider=provider_name,
            ),
            True,
        )
    except ClipboardRichDocumentError:
        raise
    except ClipboardOfficeProviderError as exc:
        raise ClipboardRichDocumentError(exc.code, str(exc)) from exc
    except (OSError, TypeError, ValueError) as exc:
        raise ClipboardRichDocumentError(
            "clipboard.structured_invalid",
            "Clipboard rich document could not be projected safely.",
        ) from exc


def project_frozen_rich_document(capture: FrozenClipboardCapture) -> RichDocumentDecision:
    """Preserve complete captured plain text when rich representation is invalid."""

    try:
        return _project_frozen_rich_document(capture)
    except ClipboardRichDocumentError:
        if capture.plain_text is not None:
            return RichDocumentDecision(None, True, plain_fallback=True)
        raise


__all__ = [
    "ClipboardRichDocumentError",
    "RichDocumentDecision",
    "RichDocumentProjection",
    "project_frozen_rich_document",
]
