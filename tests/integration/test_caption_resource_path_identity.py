"""Raw caption resource names remain distinct after final Markdown publication."""

from __future__ import annotations

import hashlib
import posixpath
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import Any
from urllib.parse import quote, unquote, urlsplit
from zipfile import ZipFile

import pytest
from docx import Document
from lxml import etree
from markdown_it import MarkdownIt
from PIL import Image

from docwen_core.markdown_extensions import MarkdownExtensions
from tests.integration._round_trip_helper import _primary_path, _run
from tests.integration.test_resource_identity_pipeline import _bundle

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@pytest.mark.parametrize("order", [(0, 1, 2, 3), (1, 2, 0, 3)])
@pytest.mark.parametrize("configured_style", ["markdown_embed", "wiki_link"])
def test_caption_paths_resolve_to_original_occurrence_bytes(
    tmp_path: Path, round_trip_runtime: Any, order: tuple[int, ...], configured_style: str
) -> None:
    names = ["photo%20one.png", "photo one.png", "photo#part.png", "photo(one).png"]
    payloads: list[bytes] = []
    source_lines: list[str] = []
    for index, color in enumerate(("red", "green", "blue", "yellow")):
        buffer = BytesIO()
        with Image.new("RGB", (3, 2), color) as image:
            image.save(buffer, format="PNG")
        payloads.append(buffer.getvalue())
        (tmp_path / names[index]).write_bytes(payloads[-1])
    for index in order:
        source_lines.append(
            f"Figure: Image {index} ^figure-{index}\n\n![Occurrence {index}]({quote(names[index], safe='')})\n\n"
            f"End of occurrence {index}.\n"
        )
    source = tmp_path / "captions.md"
    source.write_text("\n".join(source_lines), encoding="utf-8")
    result = _run(
        round_trip_runtime,
        "caption-path-writer",
        source,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "forward",
    )
    native = _primary_path(result)
    fresh = Document(BytesIO(native.read_bytes()))
    assert len(fresh.inline_shapes) == len(order)
    with ZipFile(native) as package:
        root = etree.fromstring(package.read("word/document.xml"))
        relationships = etree.fromstring(package.read("word/_rels/document.xml.rels"))
        targets = {node.get("Id"): node.get("Target") for node in relationships}
        assert [node.get("title") for node in root.iter(f"{{{_WP}}}docPr")] == [names[i] for i in order]
        assert [
            package.read(posixpath.normpath(posixpath.join("word", targets[node.get(f"{{{_R}}}embed")])))
            for node in root.iter(f"{{{_A}}}blip")
        ] == [payloads[i] for i in order]
    bundle, markdown, primary = _bundle(
        round_trip_runtime,
        native,
        tmp_path / "final",
        image_link_style=configured_style,
        markdown_extensions={"output": MarkdownExtensions(captions_references=True).to_dict()},
    )
    tokens = MarkdownIt("commonmark").parse(markdown)
    images = [child for token in tokens for child in (token.children or []) if child.type == "image"]
    assert [image.content for image in images] == [f"Occurrence {i}" for i in order]
    actual_hashes: list[str] = []
    for image in images:
        target = urlsplit(image.attrGet("src") or "")
        assert not target.scheme and not target.netloc and not target.query and not target.fragment
        logical_path = posixpath.normpath(
            posixpath.join(str(PurePosixPath(primary.logical_path).parent), unquote(target.path))
        )
        [artifact] = [artifact for artifact in bundle.artifacts if artifact.logical_path == logical_path]
        published = (tmp_path / "final" / artifact.locator).read_bytes()
        assert artifact.sha256 == hashlib.sha256(published).hexdigest()
        actual_hashes.append(artifact.sha256)
    assert actual_hashes == [hashlib.sha256(payloads[i]).hexdigest() for i in order]
    assert len(set(actual_hashes)) == 4
