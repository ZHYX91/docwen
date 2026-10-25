"""Exercise malformed spine references through the real EPUB reader and Runtime."""

from __future__ import annotations

import hashlib
from pathlib import Path
from zipfile import ZipFile

import pytest
from ebooklib import epub

from ._input_routes_support import _run_request
from ._input_routes_support import pipeline as pipeline

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def _write_epub(path: Path, spine: list[str]) -> None:
    chapters = ["one", "two", "three"]
    manifest = "".join(
        f'<item id="{name}" href="{name}.xhtml" media-type="application/xhtml+xml"/>' for name in chapters
    )
    itemrefs = "".join(f'<itemref idref="{item_id}"/>' for item_id in spine)
    package = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="book-id">'
        '<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">'
        '<dc:identifier id="book-id">spine-integrity</dc:identifier>'
        "<dc:title>Spine integrity</dc:title><dc:language>en</dc:language>"
        '<meta property="dcterms:modified">2020-01-01T00:00:00Z</meta></metadata>'
        f'<manifest>{manifest}<item id="nav" href="nav.xhtml" '
        'media-type="application/xhtml+xml" properties="nav"/></manifest>'
        f"<spine>{itemrefs}</spine></package>"
    )
    with ZipFile(path, "w") as archive:
        archive.writestr("mimetype", "application/epub+zip")
        archive.writestr(
            "META-INF/container.xml",
            '<?xml version="1.0"?><container version="1.0" '
            'xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles>'
            '<rootfile full-path="EPUB/content.opf" media-type="application/oebps-package+xml"/>'
            "</rootfiles></container>",
        )
        archive.writestr("EPUB/content.opf", package)
        for name in chapters:
            archive.writestr(
                f"EPUB/{name}.xhtml",
                '<html xmlns="http://www.w3.org/1999/xhtml"><head><title>Chapter</title></head>'
                f"<body><p>Readable body {name}.</p></body></html>",
            )
        archive.writestr(
            "EPUB/nav.xhtml",
            '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">'
            "<head><title>Navigation</title></head><body><p>Navigation sentinel.</p>"
            '<nav epub:type="toc"><ol><li><a href="one.xhtml">Chapter one</a></li></ol></nav></body></html>',
        )


@pytest.mark.parametrize(
    ("spine", "expected", "missing"),
    [
        (["two", "one", "three"], ["two", "one", "three"], []),
        (["two", "missing"], ["two", "one", "three"], ["missing"]),
        (["missing-one", "missing-two"], ["one", "two", "three"], ["missing-one", "missing-two"]),
        (["two", "missing", "missing"], ["two", "one", "three"], ["missing", "missing"]),
        (["two", "missing", "two"], ["two", "two", "one", "three"], ["missing"]),
        (["nav", "two", "missing"], ["two", "one", "three"], ["missing"]),
    ],
)
def test_missing_spine_ids_preserve_readable_chapters_and_diagnostics(
    tmp_path: Path, pipeline, spine: list[str], expected: list[str], missing: list[str]
) -> None:
    source = tmp_path / "chapters.epub"
    _write_epub(source, spine)
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    book = epub.read_epub(str(source))
    assert book.get_item_with_id("missing") is None

    _plugin, task_manager, _workspace = pipeline
    output = tmp_path / "out"
    output.mkdir()
    result = _run_request(task_manager, source, "epub", output, to_md_keep_images=False)

    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    paragraphs = [line.strip() for line in markdown.splitlines() if line.startswith("Readable body ")]
    assert paragraphs == [f"Readable body {name}." for name in expected]
    assert "Navigation sentinel." not in markdown
    diagnostics = [entry for entry in result.diagnostics if entry.code == "EPUB2MD-SPINE-ITEM-FALLBACK"]
    assert [entry.location for entry in diagnostics] == missing
    assert all(entry.level == "warning" and entry.location in entry.message for entry in diagnostics)
    assert result.metrics.extra["spine_lookup_failure_count"] == len(missing)
    assert result.metrics.extra["degradation_count"] == len(missing)
    assert result.metrics.extra["chapter_skip_count"] == 0
    assert result.artifacts[0].metadata["degradation_count"] == len(missing)
    assert hashlib.sha256(source.read_bytes()).hexdigest() == source_hash
