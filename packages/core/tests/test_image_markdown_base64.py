"""Base64 image export semantics consumed by the shared Markdown helper."""

from __future__ import annotations

import base64
import re
from io import BytesIO
from pathlib import Path
from typing import cast

import pytest
from PIL import Image

from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_core.text.image_markdown import build_base64_image_data_uri, generate_image_markdown

pytestmark = pytest.mark.unit


def _parse_data_uri(markdown: str) -> tuple[str, bytes]:
    match = re.search(r"data:([^;]+);base64,([A-Za-z0-9+/=]+)", markdown)
    assert match is not None
    return match.group(1), base64.b64decode(match.group(2))


def _noise_png(path: Path, *, mode: str = "RGB") -> None:
    image = Image.effect_noise((512, 512), 100).convert(mode)
    if mode == "RGBA":
        image.putalpha(0)
    image.save(path, format="PNG")


def _image_bytes(image_format: str) -> bytes:
    with BytesIO() as buffer:
        Image.new("RGB", (8, 8), (40, 80, 120)).save(buffer, format=image_format)
        return buffer.getvalue()


def _padded_png_bytes(size: int) -> bytes:
    payload = _image_bytes("PNG")
    assert len(payload) <= size
    return payload + b"\x00" * (size - len(payload))


def test_base64_jpg_uses_canonical_mime_and_preserves_below_threshold_bytes(
    tmp_path: Path,
) -> None:
    source = tmp_path / "photo.jpg"
    Image.new("RGB", (8, 8), (40, 80, 120)).save(source, format="JPEG")
    semantics = MarkdownExportSemantics(
        export_base64_compress_enabled=True,
        export_base64_compress_threshold_kb=10_000,
    )

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        alt_text="photo.jpg",
        export_semantics=semantics,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/jpeg"
    assert payload == source.read_bytes()


def test_base64_explicit_legacy_jpg_media_type_is_canonicalized(tmp_path: Path) -> None:
    source = tmp_path / "resource.bin"
    source_bytes = _image_bytes("JPEG")
    source.write_bytes(source_bytes)

    data_uri = build_base64_image_data_uri(
        image_path=source,
        media_type="image/jpg; charset=binary",
        export_semantics=MarkdownExportSemantics(),
    )

    media_type, payload = _parse_data_uri(data_uri)
    assert media_type == "image/jpeg"
    assert payload == source_bytes


def test_base64_declared_media_type_cannot_override_image_content(tmp_path: Path) -> None:
    source = tmp_path / "resource.jpg"
    source_bytes = _image_bytes("PNG")
    source.write_bytes(source_bytes)

    data_uri = build_base64_image_data_uri(
        image_path=source,
        media_type="image/jpeg",
        export_semantics=MarkdownExportSemantics(),
    )

    media_type, payload = _parse_data_uri(data_uri)
    assert media_type == "image/png"
    assert payload == source_bytes


@pytest.mark.parametrize(
    ("image_format", "expected_media_type"),
    [
        ("JPEG", "image/jpeg"),
        ("PNG", "image/png"),
        ("GIF", "image/gif"),
        ("BMP", "image/bmp"),
        ("TIFF", "image/tiff"),
        ("WEBP", "image/webp"),
    ],
)
def test_base64_without_explicit_media_type_uses_image_content(
    tmp_path: Path,
    image_format: str,
    expected_media_type: str,
) -> None:
    source = tmp_path / f"{image_format.lower()}.resource"
    source_bytes = _image_bytes(image_format)
    source.write_bytes(source_bytes)

    data_uri = build_base64_image_data_uri(
        image_path=source,
        export_semantics=MarkdownExportSemantics(),
    )

    media_type, payload = _parse_data_uri(data_uri)
    assert media_type == expected_media_type
    assert payload == source_bytes


@pytest.mark.parametrize("media_type", [None, "image/png"], ids=["detected", "declared"])
@pytest.mark.parametrize(
    "payload",
    [b"not an image", b"\x89PNG\r\n\x1a\ntruncated-image"],
    ids=["fake-image", "corrupt-image"],
)
def test_base64_rejects_invalid_image_content_even_with_declared_media_type(
    tmp_path: Path,
    payload: bytes,
    media_type: str | None,
) -> None:
    source = tmp_path / "disguised.png"
    source.write_bytes(payload)

    with pytest.raises(ValueError, match=r"not a supported image|corrupt or unsupported"):
        build_base64_image_data_uri(
            image_path=source,
            media_type=media_type,
            export_semantics=MarkdownExportSemantics(),
        )


def test_base64_compression_disabled_preserves_large_png_bytes(tmp_path: Path) -> None:
    source = tmp_path / "noise.png"
    _noise_png(source)
    semantics = MarkdownExportSemantics(
        export_base64_compress_enabled=False,
        export_base64_compress_threshold_kb=1,
    )

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        export_semantics=semantics,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/png"
    assert payload == source.read_bytes()


def test_base64_explicit_semantics_are_the_only_compression_source(tmp_path: Path) -> None:
    source = tmp_path / "noise.png"
    _noise_png(source)
    explicit = MarkdownExportSemantics(
        export_base64_compress_enabled=False,
        export_base64_compress_threshold_kb=1,
    )

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        export_semantics=explicit,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/png"
    assert payload == source.read_bytes()


def test_base64_rejects_missing_request_semantics(tmp_path: Path) -> None:
    source = tmp_path / "noise.png"
    _noise_png(source)
    with pytest.raises(TypeError, match="export_semantics"):
        build_base64_image_data_uri(
            image_path=source,
            export_semantics=None,  # type: ignore[arg-type]
        )


def test_base64_exact_threshold_does_not_call_compressor(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.image_markdown as image_markdown

    source = tmp_path / "exact.png"
    source.write_bytes(_padded_png_bytes(1024))
    semantics = MarkdownExportSemantics(
        export_base64_compress_enabled=True,
        export_base64_compress_threshold_kb=1,
    )

    def _unexpected(_path: Path, _target_bytes: int) -> bytes:
        raise AssertionError("strictly-equal payload must not be compressed")

    monkeypatch.setattr(image_markdown, "_compress_image_to_jpeg_bytes", _unexpected)

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        export_semantics=semantics,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/png"
    assert payload == source.read_bytes()


def test_base64_over_threshold_compresses_to_smaller_jpeg(tmp_path: Path) -> None:
    source = tmp_path / "noise.png"
    _noise_png(source)
    semantics = MarkdownExportSemantics(
        export_base64_compress_enabled=True,
        export_base64_compress_threshold_kb=100,
    )

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        alt_text="noise.png",
        export_semantics=semantics,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/jpeg"
    assert payload.startswith(b"\xff\xd8\xff")
    assert len(payload) <= 100 * 1024
    assert len(payload) < source.stat().st_size
    assert "![noise.png](data:image/jpeg;base64," in markdown


def test_base64_transparent_source_is_flattened_to_white_jpeg(tmp_path: Path) -> None:
    source = tmp_path / "transparent-noise.png"
    _noise_png(source, mode="RGBA")
    semantics = MarkdownExportSemantics(
        export_base64_compress_enabled=True,
        export_base64_compress_threshold_kb=10,
    )

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        export_semantics=semantics,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/jpeg"
    with Image.open(BytesIO(payload)) as compressed:
        assert compressed.format == "JPEG"
        assert compressed.mode == "RGB"
        red, green, blue = cast(tuple[int, int, int], compressed.getpixel((256, 256)))
        assert min(red, green, blue) >= 245


def test_base64_compression_failure_falls_back_to_original(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    import docwen_core.text.image_markdown as image_markdown

    source = tmp_path / "noise.png"
    _noise_png(source)
    semantics = MarkdownExportSemantics(
        export_base64_compress_enabled=True,
        export_base64_compress_threshold_kb=1,
    )

    def _fail(_path: Path, _target_bytes: int) -> bytes:
        raise OSError("codec unavailable")

    monkeypatch.setattr(image_markdown, "_compress_image_to_jpeg_bytes", _fail)
    caplog.set_level("WARNING")

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        export_semantics=semantics,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/png"
    assert payload == source.read_bytes()
    assert "Base64 image compression failed" in caplog.text


def test_base64_non_shrinking_compression_falls_back_to_original(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    import docwen_core.text.image_markdown as image_markdown

    source = tmp_path / "noise.png"
    source.write_bytes(_padded_png_bytes(2048))
    semantics = MarkdownExportSemantics(
        export_base64_compress_enabled=True,
        export_base64_compress_threshold_kb=1,
    )
    monkeypatch.setattr(
        image_markdown,
        "_compress_image_to_jpeg_bytes",
        lambda _path, _target_bytes: b"j" * 2048,
    )
    caplog.set_level("WARNING")

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        export_semantics=semantics,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/png"
    assert payload == source.read_bytes()
    assert "did not reduce" in caplog.text


def test_base64_smaller_jpeg_above_threshold_is_kept_with_warning(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    import docwen_core.text.image_markdown as image_markdown

    source = tmp_path / "noise.png"
    source.write_bytes(_padded_png_bytes(4096))
    compressed = b"\xff\xd8\xff" + b"j" * 1497
    semantics = MarkdownExportSemantics(
        export_base64_compress_enabled=True,
        export_base64_compress_threshold_kb=1,
    )
    monkeypatch.setattr(
        image_markdown,
        "_compress_image_to_jpeg_bytes",
        lambda _path, _target_bytes: compressed,
    )
    caplog.set_level("WARNING")

    markdown = generate_image_markdown(
        image_path=source,
        image_mode="base64",
        image_link_style="markdown_embed",
        export_semantics=semantics,
    )

    mime, payload = _parse_data_uri(markdown)
    assert mime == "image/jpeg"
    assert payload == compressed
    assert "remains above configured threshold" in caplog.text
