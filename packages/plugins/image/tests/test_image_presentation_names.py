"""Public image outputs keep the admitted name separate from physical storage."""

from pathlib import Path

import pytest
from PIL import Image

from docwen_core.models.file_ref import SOURCE_PRESENTATION_NAME_METADATA_KEY
from docwen_plugin_image.format_conversion.converter import ImageFormatConverter
from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter
from docwen_plugin_image.to_pdf.converter import ImageToPdfConverter

from ._markdown_assets_support import _build_fake_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("target", ["tif", "pdf", "md"])
@pytest.mark.parametrize("source_format", ["png", "tif"])
def test_public_image_names_and_markdown_use_admitted_presentation(
    tmp_path: Path, target: str, source_format: str
) -> None:
    source = tmp_path / f"clipboard-opaque-storage-id.{source_format}"
    Image.new("RGB", (8, 6), (23, 71, 144)).save(source)
    before = source.read_bytes()
    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(
        str(source), str(staging), target, {"to_md_enable_ocr": False, "to_md_keep_images": True}
    )
    label = "剪贴板图片 7"
    context.request.input_refs[0].metadata[SOURCE_PRESENTATION_NAME_METADATA_KEY] = f"{label}.{source_format}"
    converter = {"tif": ImageFormatConverter, "pdf": ImageToPdfConverter, "md": ImageToMarkdownConverter}[target]()

    result = converter.convert(context)

    assert result.success, result.error
    assert source.read_bytes() == before
    assert context.workspace.input_path == str(source)
    assert all(artifact.suggested_name.startswith(label) for artifact in result.artifacts)
    for artifact in result.artifacts:
        if artifact.media_type == "text/markdown":
            text = Path(artifact.staging_path).read_text(encoding="utf-8")
            assert label in text
            assert "clipboard-opaque-storage-id" not in text
        elif artifact.media_type == "image/png":
            with Image.open(artifact.staging_path) as image:
                assert image.getpixel((0, 0)) == (23, 71, 144)


def test_split_image_frames_use_admitted_name(tmp_path: Path) -> None:
    source = tmp_path / "clipboard-opaque-storage-id.tif"
    frames = [Image.new("RGB", (8, 6), color) for color in ((255, 0, 0), (0, 255, 0))]
    frames[0].save(source, save_all=True, append_images=frames[1:])
    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(str(source), str(staging), "png")
    context.request.input_refs[0].metadata[SOURCE_PRESENTATION_NAME_METADATA_KEY] = "Clipboard Image 7.tif"

    result = ImageFormatConverter().convert(context)

    assert result.success
    assert [artifact.suggested_name for artifact in result.artifacts] == [
        "Clipboard Image 7_page1.png",
        "Clipboard Image 7_page2.png",
    ]
    for artifact, expected in zip(result.artifacts, ((255, 0, 0), (0, 255, 0)), strict=True):
        with Image.open(artifact.staging_path) as image:
            assert image.getpixel((0, 0)) == expected
