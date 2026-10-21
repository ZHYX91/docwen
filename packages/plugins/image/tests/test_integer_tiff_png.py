"""Integer TIFF exports keep their explicit 16-bit PNG projection."""

from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from docwen_core.text.ocr import OcrOutcome, OcrStatus
from docwen_plugin_image.format_conversion.converter import ImageFormatConverter
from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

from ._image_conversions_support import _build_fake_context

pytestmark = [pytest.mark.integration, pytest.mark.release_gate]

SAMPLES = [-1, 0, 1, 65535, 65536, 2147483647]
PNG_SAMPLES = [0, 0, 1, 65535, 65535, 65535]


def _write_integer_tiff(path: Path, frames: int, samples: list[int] | None = None) -> None:
    images = [Image.new("I", (6, 1)) for _ in range(frames)]
    try:
        for image in images:
            image.putdata(samples if samples is not None else SAMPLES)
        images[0].save(path, format="TIFF", save_all=True, append_images=images[1:])
    finally:
        for image in images:
            image.close()


def _assert_png_samples(path: Path, expected: list[int] | None = None) -> None:
    with Image.open(path) as image:
        image.load()
        assert image.format == "PNG"
        assert image.mode == "I;16"
        assert [image.getpixel((x, 0)) for x in range(6)] == (expected if expected is not None else PNG_SAMPLES)


@pytest.mark.parametrize("frames", [1, 2])
@pytest.mark.parametrize("clipped", [False, True])
def test_integer_tiff_to_png_preserves_16bit_projection(tmp_path: Path, frames: int, clipped: bool) -> None:
    source = tmp_path / "integer.tif"
    samples = SAMPLES if clipped else [0, 1, 255, 256, 65534, 65535]
    expected = PNG_SAMPLES if clipped else samples
    _write_integer_tiff(source, frames, samples)
    original = source.read_bytes()
    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(str(source), str(staging), "png", source_format="tif")

    result = ImageFormatConverter().convert(context)

    assert result.success, result.error
    assert len(result.artifacts) == frames
    for artifact in result.artifacts:
        _assert_png_samples(Path(artifact.staging_path), expected)
    warnings = [item for item in result.diagnostics if item.code == "IMAGE-PNG-16BIT-CLIPPED"]
    assert len(warnings) == (frames if clipped else 0)
    assert all(item.level == "warning" and "Use TIFF" in item.message for item in warnings)
    assert source.read_bytes() == original


@pytest.mark.parametrize("keep_images", [False, True])
def test_integer_tiff_ocr_decodes_each_private_png_page(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, keep_images: bool
) -> None:
    import docwen_plugin_image.to_markdown.converter as converter

    source = tmp_path / "integer.tif"
    _write_integer_tiff(source, 2)
    staging = tmp_path / "staging"
    staging.mkdir()
    decoded: list[Path] = []

    def read_page(path: str, **_kwargs: object) -> OcrOutcome:
        _assert_png_samples(Path(path))
        decoded.append(Path(path))
        return OcrOutcome(OcrStatus.SUCCESS, text=f"PAGE {len(decoded)}")

    monkeypatch.setattr(converter, "run_ocr_outcome", read_page)
    context = _build_fake_context(
        str(source),
        str(staging),
        "md",
        {"to_md_enable_ocr": True, "to_md_keep_images": keep_images},
        source_format="tif",
    )

    result = ImageToMarkdownConverter().convert(context)

    assert result.success, result.error
    assert len(decoded) == 2
    fragments = [artifact for artifact in result.artifacts if artifact.kind == "auxiliary"]
    assert all(
        f"PAGE {index}" in Path(artifact.staging_path).read_text(encoding="utf-8")
        for index, artifact in enumerate(fragments, 1)
    )
    assert len(list(staging.glob("*.png"))) == (2 if keep_images else 0)
    warnings = [item for item in result.diagnostics if item.code == "IMAGE-PNG-16BIT-CLIPPED"]
    assert [item.location for item in warnings] == ["integer.tif:frame-1", "integer.tif:frame-2"]
