"""Project one frozen clipboard capture into a structured document when provable."""

from __future__ import annotations

from dataclasses import dataclass

from docwen_core.models.clipboard_document import ClipboardDocument, clipboard_document_to_bytes
from docwen_gui.clipboard_capture import FrozenClipboardCapture
from docwen_gui.clipboard_image_binding import bind_provider_images
from docwen_gui.clipboard_office_provider import (
    WORD_EMBED_SOURCE_MIME,
    WPS_DOCUMENT_MIME,
    WPS_IMAGE_DATA_MIME,
    ClipboardOfficeProviderError,
    parse_word_embed_source,
    parse_wps_writer,
)
from docwen_gui.clipboard_structured import (
    StructuredClipboardProjection,
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


def project_frozen_rich_document(capture: FrozenClipboardCapture) -> RichDocumentDecision:
    """Use plain text as body authority and provider bytes only for proven images."""

    html_bytes = capture.html_bytes
    word_payload = capture.get(WORD_EMBED_SOURCE_MIME)
    wps_document = capture.get(WPS_DOCUMENT_MIME)
    wps_images = capture.get(WPS_IMAGE_DATA_MIME)
    has_provider = word_payload is not None or wps_document is not None or wps_images is not None

    if not html_bytes:
        if has_provider:
            raise ClipboardRichDocumentError(
                "clipboard.provider_html_missing",
                "Rich clipboard provider data has no captured HTML structure.",
            )
        return RichDocumentDecision(None, False)

    try:
        contains_table = html_contains_table(html_bytes)
    except (UnicodeDecodeError, ValueError) as exc:
        raise ClipboardRichDocumentError(
            "clipboard.structured_invalid",
            "Clipboard rich document structure is invalid.",
        ) from exc
    if not contains_table and not has_provider:
        return RichDocumentDecision(None, False)

    try:
        html_projection: StructuredClipboardProjection = project_structured_clipboard_html_with_resources(html_bytes)
        document = html_projection.document
        if capture.plain_text is not None:
            merged = splice_plain_text_with_structured_tables(document, capture.plain_text)
            if merged is None:
                return RichDocumentDecision(None, True, plain_fallback=True)
            document = merged

        provider_projection = None
        if word_payload is not None:
            if wps_document is not None or wps_images is not None:
                raise ClipboardRichDocumentError(
                    "clipboard.provider_ambiguous",
                    "Clipboard exposes more than one rich document provider.",
                )
            provider_projection = parse_word_embed_source(word_payload)
        elif wps_document is not None or wps_images is not None:
            if wps_document is None or wps_images is None:
                raise ClipboardRichDocumentError(
                    "clipboard.provider_incomplete",
                    "WPS clipboard image provider data is incomplete.",
                )
            provider_projection = parse_wps_writer(wps_document, wps_images)

        resources = html_projection.resources
        provider_name = ""
        if provider_projection is not None:
            document, resources = bind_provider_images(document, provider_projection, resources)
            provider_name = provider_projection.provider

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


__all__ = [
    "ClipboardRichDocumentError",
    "RichDocumentDecision",
    "RichDocumentProjection",
    "project_frozen_rich_document",
]
