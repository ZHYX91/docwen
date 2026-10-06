"""Cross-plugin note dialect round-trip regressions."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree

from tests.integration._round_trip_helper import docx_to_md, md_to_docx

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def test_ordinary_endnote_dash_id_unicode_ids_and_multiline_note_round_trip(
    round_trip_runtime,
    tmp_path: Path,
) -> None:
    source_path = tmp_path / "notes.md"
    source_path.write_text(
        "Dash[^endnote-topic], sharp[^Straße], ss[^Strasse], multiline[^multi].\n\n"
        "[^endnote-topic]: Ordinary footnote.\n"
        "[^Straße]: Sharp-s identity.\n"
        "[^Strasse]: SS identity.\n"
        "[^multi]: First line\n"
        "  second line\n"
        "  **third line**\n",
        encoding="utf-8",
    )

    output = md_to_docx(
        round_trip_runtime,
        source_path,
        tmp_path / "docx",
        request_id="note-dialect-forward",
    )

    with ZipFile(output) as package:
        document = etree.fromstring(package.read("word/document.xml"))
        footnotes = etree.fromstring(package.read("word/footnotes.xml"))
        assert not list(document.iter(f"{{{WML_NS}}}endnoteReference"))
        assert len(list(document.iter(f"{{{WML_NS}}}footnoteReference"))) == 4
        authored_notes = [
            item
            for item in footnotes.findall(f"{{{WML_NS}}}footnote")
            if int(item.get(f"{{{WML_NS}}}id", "0")) > 0
        ]
        assert len(authored_notes) == 4

    markdown = docx_to_md(
        round_trip_runtime,
        output,
        tmp_path / "markdown",
        request_id="note-dialect-reverse",
        preserve_numbering=False,
    )

    assert "Ordinary footnote." in markdown
    assert "Sharp-s identity." in markdown
    assert "SS identity." in markdown
    assert "First line\n    second line\n    **third line**" in markdown
