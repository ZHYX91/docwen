"""Actual Word text and notes across generated Markdown presentation boundaries."""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from ._link_processing_routes_support import _TINY_PNG, _convert_docx, _link_config

pytestmark = pytest.mark.contract

_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
_NOTES = "\n\nFoot[^f], end[^endnote:e].\n\n[^f]: Foot.\n[^endnote:e]: End.\n"
_OPTIONS = {"markdown_extensions": {"input": {"typed_endnotes": True, "structural_tables": True}}}


@pytest.mark.parametrize(
    "syntax,visible",
    [
        ("[](local(foo)%%draft)", "local(foo)%%draft"),
        ("[](local[^fake])", "local[^fake]"),
        ("![remote](https://example.com/image.png)", "https://example.com/image.png"),
        ("![remote](http://example.com/image.png)", "http://example.com/image.png"),
        ("![remote](//example.com/image.png)", "//example.com/image.png"),
        ("![[https://example.com/remote.md]]", "https://example.com/remote.md"),
        ("![[chart.png]]", "chart.png"),
    ],
)
def test_generated_destination_word_text_and_typed_notes(syntax: str, visible: str, tmp_path: Path) -> None:
    source = tmp_path / "presentation.md"
    original = (syntax + _NOTES).encode()
    source.write_bytes(original)
    config = _link_config(markdown_mode="extract_text")
    if syntax == "![[chart.png]]":
        config = _link_config(wiki_image_mode="extract_text")
    output = _convert_docx(source, config, options=_OPTIONS)
    assert visible in output.text
    assert "\\" not in output.text
    root = ET.fromstring(output.document_xml)
    for kind in ("footnote", "endnote"):
        [reference] = root.findall(f".//w:{kind}Reference", _NS)
        assert int(reference.attrib[f"{{{_NS['w']}}}id"]) > 0
    assert source.read_bytes() == original


@pytest.mark.parametrize("prefix", ["", "> ", "- "])
@pytest.mark.parametrize("gap", ["\n", "\n\n"])
@pytest.mark.parametrize("content", ["[Visible](https://example.test/page)", "![[child.md]]", "![Image](image.png)"])
@pytest.mark.parametrize("block_comment", [False, True])
def test_comment_adjacent_content_remains_active_in_word(
    prefix: str, gap: str, content: str, block_comment: bool, tmp_path: Path
) -> None:
    child = tmp_path / "child.md"
    child.write_text("Expanded child.", encoding="utf-8")
    image = tmp_path / "image.png"
    image.write_bytes(_TINY_PNG)
    source = tmp_path / "adjacent.md"
    # The frozen renderer hides standalone delimiter-line comments only.
    continuation = "  " if prefix == "- " else prefix
    comment = f"{prefix}%%\n{continuation}hidden\n{continuation}%%" if block_comment else f"{prefix}%% hidden %%"
    original = (f"{comment}{gap}{prefix}{content}" + _NOTES).encode()
    source.write_bytes(original)
    output = _convert_docx(source, _link_config(markdown_mode="extract_text"), options=_OPTIONS)
    assert ("hidden" in output.text) is not block_comment
    assert ("%%" in output.text) is not block_comment
    if content.startswith("![Image]"):
        assert len(output.media_names) == 1
    else:
        assert ("Expanded child." if content.startswith("!") else "Visible") in output.text
    assert "https://example.test/page" not in output.text
    root = ET.fromstring(output.document_xml)
    assert len(root.findall(".//w:footnoteReference", _NS)) == 1
    assert len(root.findall(".//w:endnoteReference", _NS)) == 1
    assert source.read_bytes() == original
    assert child.read_text(encoding="utf-8") == "Expanded child."
    assert image.read_bytes() == _TINY_PNG


@pytest.mark.parametrize("slashes", [0, 1, 2, 3, 4])
def test_actual_word_notes_follow_source_escape_parity(slashes: int, tmp_path: Path) -> None:
    source = tmp_path / "note-parity.md"
    prefix = "\\" * slashes
    original = f"Foot {prefix}[^f], end {prefix}[^endnote:e].\n\n[^f]: Foot.\n[^endnote:e]: End.\n".encode()
    source.write_bytes(original)
    output = _convert_docx(source, _link_config(), options=_OPTIONS)
    root = ET.fromstring(output.document_xml)
    for kind in ("footnote", "endnote"):
        references = root.findall(f".//w:{kind}Reference", _NS)
        assert len(references) == (0 if slashes % 2 else 1)
        assert all(int(reference.attrib[f"{{{_NS['w']}}}id"]) > 0 for reference in references)
    assert source.read_bytes() == original


@pytest.mark.parametrize("slashes", [0, 1, 2, 3, 4])
def test_actual_word_reference_follows_source_escape_parity(slashes: int, tmp_path: Path) -> None:
    source = tmp_path / "reference-parity.md"
    original = ("Figure: Caption ^target\n\nText " + "\\" * slashes + "@[[#^target]].\n").encode()
    source.write_bytes(original)
    options = {"markdown_extensions": {"input": {"captions_references": True}}}
    output = _convert_docx(source, _link_config(), options=options)
    root = ET.fromstring(output.document_xml)
    instructions = root.findall(".//w:instrText", _NS)
    assert sum("REF " in (node.text or "") for node in instructions) == (0 if slashes % 2 else 1)
    assert source.read_bytes() == original
