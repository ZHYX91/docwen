"""Clipboard image capture, provider selection, and predecode budget contracts."""

from __future__ import annotations

import base64
import hashlib
import json
import struct
import zipfile
from io import BytesIO
from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData

from docwen_core.models.clipboard_document import (
    CLIPBOARD_DOCUMENT_SCHEMA,
    ClipboardDocumentError,
    ClipboardImageRef,
    iter_clipboard_inlines,
)
from docwen_gui import clipboard_capture as capture_module
from docwen_gui import clipboard_inputs as inputs_module
from docwen_gui import clipboard_office_provider as provider
from docwen_gui import clipboard_rich_document as rich
from docwen_gui import clipboard_structured as structured
from docwen_gui.clipboard_capture import FrozenClipboardCapture, freeze_clipboard_mime
from docwen_gui.clipboard_office_provider import (
    WORD_EMBED_SOURCE_MIME,
    WPS_DOCUMENT_MIME,
    WPS_IMAGE_DATA_MIME,
)

pytestmark = pytest.mark.contract

_SAMPLES = Path(__file__).resolve().parents[4] / "tests/fixtures/files/clipboard-office"


def _fake_png_header(width: int, height: int, suffix: bytes = b"") -> bytes:
    return (
        b"\x89PNG\r\n\x1a\n"
        + struct.pack(">I", 13)
        + b"IHDR"
        + struct.pack(">II", width, height)
        + b"\x08\x06\x00\x00\x00"
        + b"\x00\x00\x00\x00"
        + suffix
    )


def _data_uri(payload: bytes) -> str:
    return "data:image/png;base64," + base64.b64encode(payload).decode("ascii")


def _wps_document(raw_ids: list[bytes]) -> bytes:
    paragraphs = []
    for raw_id in raw_ids:
        identifier = base64.b64encode(raw_id).decode("ascii")
        paragraphs.append(
            "<w:p><w:r><w:drawing><wp:inline><wp:extent cx='914400' cy='609600'/>"
            "<a:graphic><a:graphicData>"
            f"<a:blip r:embed='{identifier}'/>"
            "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
        )
    xml = (
        "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' "
        "xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' "
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' "
        "xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'>"
        f"<w:body>{''.join(paragraphs)}</w:body></w:document>"
    ).encode()
    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", xml)
    return output.getvalue()


def _wps_records(items: list[tuple[bytes, bytes]]) -> bytes:
    output = bytearray()
    for raw_id, payload in items:
        record = raw_id + payload
        output += struct.pack("<I", len(record))
        output += record
    return bytes(output)


def _cf_html(fragment: str) -> bytes:
    prefix = (
        "Version:1.0\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_fragment:010d}\r\n"
        "EndFragment:{end_fragment:010d}\r\n"
        "SourceURL:file:///C:/private/source.docx\r\n"
    )
    html = (
        "<html><head><base href='https://example.invalid/base/'></head><body><!--StartFragment-->"
        f"{fragment}<!--EndFragment--></body></html>"
    ).encode()
    placeholder = prefix.format(start_html=0, end_html=0, start_fragment=0, end_fragment=0).encode()
    start_html = len(placeholder)
    marker_start = html.index(b"<!--StartFragment-->") + len(b"<!--StartFragment-->")
    marker_end = html.index(b"<!--EndFragment-->")
    header = prefix.format(
        start_html=start_html,
        end_html=start_html + len(html),
        start_fragment=start_html + marker_start,
        end_fragment=start_html + marker_end,
    ).encode()
    return header + html


class _OversizeArray:
    def size(self) -> int:
        return capture_module._CAPTURE_FORMAT_LIMITS[WORD_EMBED_SOURCE_MIME] + 1

    def data(self):
        pytest.fail("oversized QByteArray was copied before its size was checked")


class _OversizeMimeData(QMimeData):
    def formats(self) -> list[str]:
        return ["text/plain", WORD_EMBED_SOURCE_MIME]

    def data(self, mime_type: str):
        assert mime_type == WORD_EMBED_SOURCE_MIME
        return _OversizeArray()


def test_capture_checks_qbytearray_size_before_copy_and_keeps_plain() -> None:
    mime = _OversizeMimeData()
    mime.setText("complete plain fallback")

    frozen = freeze_clipboard_mime(mime)

    assert frozen.plain_text == "complete plain fallback"
    assert frozen.binary_formats == ()
    assert frozen.rich_error_code == "clipboard.capture_budget_exceeded"


def test_complete_wps_pair_does_not_copy_generic_embed_source() -> None:
    reads: list[str] = []

    class TrackingMimeData(QMimeData):
        def data(self, mime_type: str):
            reads.append(mime_type)
            if mime_type == WORD_EMBED_SOURCE_MIME:
                pytest.fail("generic Embed Source must not be copied when the complete WPS pair is present")
            return super().data(mime_type)

    mime = TrackingMimeData()
    mime.setData(WORD_EMBED_SOURCE_MIME, b"controlled sentinel; not a real captured generic OLE payload")
    mime.setData(WPS_DOCUMENT_MIME, b"wps-document")
    mime.setData(WPS_IMAGE_DATA_MIME, b"wps-images")

    frozen = freeze_clipboard_mime(mime)

    assert WORD_EMBED_SOURCE_MIME not in reads
    assert frozen.get(WORD_EMBED_SOURCE_MIME) is None
    assert frozen.get(WPS_DOCUMENT_MIME) == b"wps-document"
    assert frozen.get(WPS_IMAGE_DATA_MIME) == b"wps-images"


def test_rich_adapter_uses_complete_wps_pair_even_if_generic_embed_source_is_present(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    seen: list[str] = []
    actual_parse_wps = rich.parse_wps_writer

    def parse_wps(document_payload: bytes, image_payload: bytes):
        projection = actual_parse_wps(document_payload, image_payload)
        seen.append(projection.provider)
        return projection

    def reject_word(_payload: bytes):
        pytest.fail("generic Embed Source was parsed even though the WPS-specific pair was complete")

    def accept_binding(document, projection, html_resources):
        assert projection.provider == "wps_writer"
        return document, html_resources

    monkeypatch.setattr(rich, "parse_wps_writer", parse_wps)
    monkeypatch.setattr(rich, "parse_embed_source_images", reject_word)
    monkeypatch.setattr(rich, "bind_provider_images", accept_binding)
    frozen = FrozenClipboardCapture(
        plain_text=None,
        html_bytes=b"<p>before<img src='unavailable' alt='placeholder'>after</p>",
        binary_formats=(
            (WORD_EMBED_SOURCE_MIME, b"controlled sentinel; not a real captured generic OLE payload"),
            (WPS_DOCUMENT_MIME, (_SAMPLES / "wps-writer-derived.zip").read_bytes()),
            (WPS_IMAGE_DATA_MIME, (_SAMPLES / "wps-images.bin").read_bytes()),
        ),
        image=None,
    )

    decision = rich.project_frozen_rich_document(frozen)

    assert decision.handled and decision.projection is not None
    assert decision.projection.provider == "wps_writer"
    assert seen == ["wps_writer"]


def test_incomplete_wps_pair_with_generic_embed_source_falls_back_without_guessing_word(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        rich,
        "parse_embed_source_images",
        lambda _payload: pytest.fail("incomplete WPS pair must not guess that generic Embed Source is Word"),
    )
    frozen = FrozenClipboardCapture(
        plain_text="keep this exact plain body",
        html_bytes=b"<p>keep this exact plain body</p>",
        binary_formats=(
            (WORD_EMBED_SOURCE_MIME, b"controlled generic sentinel"),
            (WPS_DOCUMENT_MIME, b"incomplete-wps-document"),
        ),
        image=None,
    )

    decision = rich.project_frozen_rich_document(frozen)

    assert decision.handled and decision.plain_fallback and decision.projection is None


def test_external_html_sources_never_become_resources_and_keep_alt_diagnostics() -> None:
    projection = structured.project_structured_clipboard_html_with_resources(
        _cf_html(
            "<p>before"
            "<img src='https://example.invalid/a.png' alt='http-alt'>"
            "<img src='file:///C:/private/b.png' alt='file-alt'>"
            "<img src='relative/c.png' alt='relative-alt'>after</p>"
        )
    )

    assert projection.resources == ()
    images = [item for item in iter_clipboard_inlines(projection.document) if isinstance(item, ClipboardImageRef)]
    assert [item.alt for item in images] == ["http-alt", "file-alt", "relative-alt"]
    assert {item.missing_reason for item in images} == {"clipboard_resource_unavailable"}


@pytest.mark.parametrize("budget", ["count", "bytes", "pixels"])
def test_html_data_uri_budget_is_checked_for_all_resources_before_decoder(
    budget: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        structured,
        "inspect_png_bytes",
        lambda _payload: pytest.fail("pixel decoder entered before the complete HTML resource preflight"),
    )
    payloads = [_fake_png_header(5000, 5000, bytes([index])) for index in range(3)]
    if budget == "count":
        monkeypatch.setattr(structured, "MAX_CLIPBOARD_RESOURCES", 2)
        payloads = [_fake_png_header(1, 1, bytes([index])) for index in range(3)]
    elif budget == "bytes":
        monkeypatch.setattr(structured, "MAX_CLIPBOARD_RESOURCE_BYTES", 60)
        payloads = [_fake_png_header(1, 1, bytes([index])) for index in range(2)]
    html = "<p>" + "".join(f"<img src='{_data_uri(payload)}' alt='x'>" for payload in payloads) + "</p>"

    with pytest.raises(ClipboardDocumentError) as failure:
        structured.project_structured_clipboard_html_with_resources(html.encode("utf-8"))

    assert failure.value.code == "clipboard.image_data_uri_invalid"
    assert str(failure.value) == "Clipboard PNG data URI is invalid or exceeds the image budget."


@pytest.mark.parametrize("budget", ["count", "bytes", "pixels"])
def test_wps_provider_budget_is_checked_before_any_pixel_decoder(
    budget: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        provider,
        "inspect_png_bytes",
        lambda _payload: pytest.fail("pixel decoder entered before the complete provider resource preflight"),
    )
    raw_ids = [(index + 1).to_bytes(16, "big") for index in range(3)]
    payloads = [_fake_png_header(5000, 5000, bytes([index])) for index in range(3)]
    if budget == "count":
        monkeypatch.setattr(provider, "MAX_CLIPBOARD_RESOURCES", 2)
        payloads = [_fake_png_header(1, 1, bytes([index])) for index in range(3)]
    elif budget == "bytes":
        monkeypatch.setattr(provider, "MAX_CLIPBOARD_RESOURCE_BYTES", 60)
        raw_ids = raw_ids[:2]
        payloads = [_fake_png_header(1, 1, bytes([index])) for index in range(2)]

    with pytest.raises(provider.ClipboardOfficeProviderError) as failure:
        provider.parse_wps_writer(
            _wps_document(raw_ids),
            _wps_records(list(zip(raw_ids, payloads, strict=True))),
        )

    assert failure.value.code == "clipboard.provider_image_budget_exceeded"
    assert str(failure.value) == "Clipboard provider images exceed the supported resource budget."


@pytest.mark.parametrize(
    ("case", "expected_message"),
    [
        ("duplicate", "WPS clipboard image identifier is duplicated."),
        ("length", "WPS clipboard image record length is invalid."),
        ("tail", "WPS clipboard image data has trailing bytes."),
        ("bad_magic", "Clipboard provider image is invalid."),
    ],
)
def test_wps_provider_malformed_resources_fail_with_stable_sanitized_diagnostics(
    case: str,
    expected_message: str,
) -> None:
    raw_id = b"A" * 16
    document = _wps_document([raw_id])
    if case == "duplicate":
        image_data = _wps_records([(raw_id, b"x"), (raw_id, b"y")])
        expected_code = "clipboard.provider_wps_image_data_invalid"
    elif case == "length":
        image_data = struct.pack("<I", 16) + raw_id
        expected_code = "clipboard.provider_wps_image_data_invalid"
    elif case == "tail":
        image_data = _wps_records([(raw_id, b"x")]) + b"\x00\x01"
        expected_code = "clipboard.provider_wps_image_data_invalid"
    else:
        image_data = _wps_records([(raw_id, b"not-a-png")])
        expected_code = "clipboard.provider_image_invalid"

    with pytest.raises(provider.ClipboardOfficeProviderError) as failure:
        provider.parse_wps_writer(document, image_data)

    assert failure.value.code == expected_code
    assert str(failure.value) == expected_message
    assert "not-a-png" not in str(failure.value)


def test_store_preflights_actual_png_headers_before_decoding_understated_declarations(tmp_path, monkeypatch):
    monkeypatch.setattr(
        inputs_module,
        "inspect_png_bytes",
        lambda _payload: pytest.fail("Store decoded before complete resource preflight"),
    )
    declarations = []
    supplied = []
    for index in range(3):
        payload = _fake_png_header(5000, 5000, bytes([index]))
        resource_id = f"image-{index}"
        logical_path = f"resources/{index}.png"
        declarations.append(
            {
                "resourceId": resource_id,
                "logicalPath": logical_path,
                "mediaType": "image/png",
                "sizeBytes": len(payload),
                "sha256": hashlib.sha256(payload).hexdigest(),
                "pixelWidth": 1,
                "pixelHeight": 1,
                "rgbaSha256": "0" * 64,
            }
        )
        supplied.append((resource_id, logical_path, "image/png", payload))
    document = json.dumps({"schema": CLIPBOARD_DOCUMENT_SCHEMA, "blocks": [], "resources": declarations}).encode()
    store = inputs_module.ClipboardInputStore(tmp_path / "clipboard")
    try:
        with pytest.raises(ValueError, match="PNG resource is invalid"):
            store.create_bundle(
                document, display_name_template="Clipboard {index}.dwclip", preview="", resources=tuple(supplied)
            )
        assert not list(store.session_root.glob("bundle-*"))
    finally:
        store.close()


def test_wps_aliases_decode_one_unique_png_once(monkeypatch):
    from PIL import Image

    with BytesIO() as stream:
        Image.new("RGBA", (2, 2), (1, 2, 3, 120)).save(stream, format="PNG")
        payload = stream.getvalue()
    original = provider.inspect_png_bytes
    calls = []

    def decode_once(data):
        calls.append(data)
        return original(data)

    monkeypatch.setattr(provider, "inspect_png_bytes", decode_once)
    raw_ids = [b"A" * 16, b"B" * 16]
    projection = provider.parse_wps_writer(_wps_document(raw_ids), _wps_records([(item, payload) for item in raw_ids]))
    assert calls == [payload]
    assert len(projection.resources) == 1
    assert len(projection.occurrences) == 2
    assert len({resource_id for _identifier, resource_id in projection.identifier_to_resource_id}) == 1
