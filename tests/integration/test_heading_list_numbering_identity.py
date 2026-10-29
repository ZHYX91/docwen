"""A native list owns Heading blocks without changing authored Heading text."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from markdown_it import MarkdownIt

from tests.integration.test_formula_and_gongwen_source_content import _body, _bundle

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS = {"w": _W}


@pytest.mark.parametrize("remove_numbering", [False, True])
@pytest.mark.parametrize("start", [None, 0, 2], ids=["bullet", "zero", "ordered"])
def test_native_container_marker_precedes_heading_and_counts_once(
    round_trip_runtime: Any, tmp_path: Path, remove_numbering: bool, start: int | None
) -> None:
    markers = ["-"] * 3 if start is None else [f"{value}." for value in range(start, start + 3)]
    source = tmp_path / "headings-in-list.md"
    source.write_text(
        f"{markers[0]} # First ^first\n{markers[1]} # Second ^second\n{markers[2]} Tail.\n\nOutside.\n",
        encoding="utf-8",
    )
    output, _ = _bundle(
        round_trip_runtime,
        source,
        tmp_path / "written",
        source_format="markdown",
        capability="convert.markdown_source.to_docx",
        options={"remove_numbering": False, "markdown_extensions": {"input": {"captions_references": True}}},
    )
    with ZipFile(output) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        numbering = ET.fromstring(package.read("word/numbering.xml"))
    headings = [
        paragraph
        for paragraph in document.findall(".//w:p", _NS)
        if paragraph.find("w:pPr/w:pStyle[@w:val='Heading1']", _NS) is not None
    ]
    assert ["".join(node.text or "" for node in p.findall(".//w:t", _NS)) for p in headings] == ["First", "Second"]
    num_ids = [p.find("w:pPr/w:numPr/w:numId", _NS).get(f"{{{_W}}}val") for p in headings]
    assert len(set(num_ids)) == 1
    num = numbering.find(f"w:num[@w:numId='{num_ids[0]}']", _NS)
    abstract_id = num.find("w:abstractNumId", _NS).get(f"{{{_W}}}val")
    abstract = numbering.find(f"w:abstractNum[@w:abstractNumId='{abstract_id}']", _NS)
    assert abstract.find("w:multiLevelType", _NS).get(f"{{{_W}}}val") == "hybridMultilevel"
    level = abstract.find("w:lvl[@w:ilvl='0']", _NS)
    assert level.find("w:pStyle", _NS) is None
    assert level.find("w:numFmt", _NS).get(f"{{{_W}}}val") == ("bullet" if start is None else "decimal")
    assert level.find("w:start", _NS).get(f"{{{_W}}}val") == str(1 if start is None else start)
    recovered, _ = _bundle(
        round_trip_runtime,
        output,
        tmp_path / "read",
        source_format="docx",
        capability="convert.docx.to_markdown",
        options={
            "remove_numbering": remove_numbering,
            "markdown_extensions": {"output": {"captions_references": True}},
        },
    )
    _, body = _body(recovered.read_text(encoding="utf-8"))
    expected_lines = [f"{markers[0]} # First ^first", f"{markers[1]} # Second ^second", f"{markers[2]} Tail."]
    assert [line for line in body.splitlines() if line] == [*expected_lines, "Outside."]
    tokens = MarkdownIt("commonmark").parse(body)
    assert sum(token.type == "list_item_open" for token in tokens) == 3
    assert [tokens[i + 1].content for i, token in enumerate(tokens) if token.type == "heading_open"] == [
        "First ^first",
        "Second ^second",
    ]
    if start is not None:
        [list_token] = [token for token in tokens if token.type == "ordered_list_open"]
        assert list_token.attrGet("start") == start
    assert tokens[-2].type == "inline" and tokens[-2].content == "Outside."
