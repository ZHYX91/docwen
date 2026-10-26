"""Image metadata uses the same capacity boundary as other XLSX sheets."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest
from openpyxl import load_workbook
from PIL import Image

from docwen_core.models.clipboard_document import ClipboardImageRef, ClipboardResource
from docwen_plugin_markdown.structured_clipboard_images import BoundClipboardImage, ClipboardImageOccurrenceInfo
from docwen_plugin_markdown.structured_clipboard_output_images import add_xlsx_images
from docwen_plugin_markdown.structured_clipboard_xlsx import prepare_projection_workbook
from docwen_plugin_markdown.xlsx_text import XlsxCellTextTooLongError

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("units", [32767, 32768])
def test_image_semantics_checks_final_logical_path(tmp_path: Path, units: int) -> None:
    image = tmp_path / "image.png"
    Image.new("RGB", (2, 2), (120, 20, 20)).save(image)
    payload = image.read_bytes()
    logical_path = "a" * (units - 4) + ".png"
    resource = ClipboardResource(
        "image-1", logical_path, "image/png", len(payload), hashlib.sha256(payload).hexdigest()
    )
    occurrence = ClipboardImageOccurrenceInfo(1, ClipboardImageRef("image-1", "image"), "", "")
    resources = {"image-1": BoundClipboardImage(resource, image, "image.png")}
    projection = prepare_projection_workbook(None)
    if units > 32767:
        with pytest.raises(XlsxCellTextTooLongError, match=r"Image Semantics.*row 2, column 4"):
            add_xlsx_images(projection, (occurrence,), resources, {})
    else:
        add_xlsx_images(projection, (occurrence,), resources, {})
        output = tmp_path / "images.xlsx"
        projection.workbook.save(output)
        workbook = load_workbook(output)
        assert workbook["Image Semantics"]["D2"].value == logical_path
        workbook.close()
