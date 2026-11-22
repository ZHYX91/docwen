"""The same opening lines must own the same source through actual exports."""

import csv
from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from PIL import Image

from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_spreadsheet.converter import MdToCsvConverter

from .conftest import make_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("prefix", "visible"),
    [
        ("---\n\n```txt\n---\n```\n\n", "---"),
        ("---\n---\n", None),
        ("---\ntitle: hidden property\n...\n", None),
        ("---\n\nfirst paragraph\n\n", "first paragraph"),
    ],
)
def test_word_keeps_body_image_and_heading_after_leading_dashes(
    tmp_path: Path, prefix: str, visible: str | None
) -> None:
    source = tmp_path / "note.md"
    image = tmp_path / "linked.png"
    Image.new("RGB", (3, 3), "blue").save(image)
    data = (prefix + "# Heading ^anchor\n\n![photo](linked.png)\n\n---\n\nLast line\n").replace("\n", "\r\n").encode()
    source.write_bytes(data)
    ctx, _ = make_context(str(source), target_format="docx")
    result = MdToDocxConverter().convert(ctx)
    assert result.success, result.error
    document = Document(result.artifacts[0].staging_path)
    texts = document.element.xpath("//w:t/text()")
    assert "Heading" in texts and "Last line" in texts and "hidden property" not in " ".join(texts)
    assert document.element.xpath("//w:bookmarkStart")
    if visible:
        assert visible in texts
    with ZipFile(result.artifacts[0].staging_path) as package:
        assert [sha256(package.read(p)).digest() for p in package.namelist() if p.startswith("word/media/")] == [
            sha256(image.read_bytes()).digest()
        ]
    assert source.read_bytes() == data


def test_csv_keeps_table_between_initial_and_later_rule(tmp_path: Path) -> None:
    source = tmp_path / "table.md"
    source.write_text("---\n\n| Item | Count |\n| --- | --- |\n| ALPHA | 9 |\n\n---\n", encoding="utf-8")
    ctx, _ = make_context(str(source), target_format="csv")
    result = MdToCsvConverter().convert(ctx)
    assert result.success, result.error
    with Path(result.artifacts[0].staging_path).open(encoding="utf-8-sig", newline="") as stream:
        assert list(csv.reader(stream)) == [["Item", "Count"], ["ALPHA", "9"]]


def test_materialized_body_image_is_never_reclassified_as_properties(tmp_path: Path) -> None:
    source = tmp_path / "note.md"
    source.write_text("---\n\n![visible](linked.png)\n\n---\n", encoding="utf-8")
    image = tmp_path / "linked.png"
    Image.new("RGB", (2, 2), "green").save(image)
    ctx, _ = make_context(str(source), target_format="docx")
    result = MdToDocxConverter().convert(ctx)
    assert result.success, result.error
    with ZipFile(result.artifacts[0].staging_path) as package:
        assert b"{{IMAGE" not in package.read("word/document.xml")
        assert [package.read(p) for p in package.namelist() if p.startswith("word/media/")] == [image.read_bytes()]
