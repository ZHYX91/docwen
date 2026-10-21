"""Exercise upgraded image codecs through the public Markdown conversion route."""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from docx import Document
from PIL import Image

from tests.integration._round_trip_helper import md_to_docx

pytestmark = [pytest.mark.integration, pytest.mark.release_gate]


@pytest.mark.parametrize("codec,suffix", [("JPEG", "jpg"), ("TIFF", "tif"), ("GIF", "gif")])
def test_markdown_embeds_real_decoded_images(round_trip_runtime: Any, tmp_path: Path, codec: str, suffix: str) -> None:
    image_path = tmp_path / f"image.{suffix}"
    with Image.new("RGB", (32, 24), (180, 80, 40)) as image:
        image.save(image_path, format=codec)
    original = image_path.read_bytes()
    source = tmp_path / "image.md"
    source.write_text(f"Before\n\n![sample]({image_path.name})\n\nAfter\n", encoding="utf-8")

    output = md_to_docx(round_trip_runtime, source, tmp_path / "output")

    document = Document(str(output))
    assert len(document.inline_shapes) == 1
    assert document.inline_shapes[0].width > 0 and document.inline_shapes[0].height > 0
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    assert "Before" in text and "After" in text
    with ZipFile(output) as archive:
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        assert len(media) == 1
        with Image.open(io.BytesIO(archive.read(media[0]))) as embedded:
            embedded.load()
            assert embedded.size == (32, 24)
    assert image_path.read_bytes() == original
