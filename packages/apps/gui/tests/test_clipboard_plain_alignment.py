"""Lossless table text alignment and provider paragraph identity."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from docwen_core.models.clipboard_document import (
    ClipboardHardBreak,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardTable,
    ClipboardText,
    iter_clipboard_inlines,
)
from docwen_gui import clipboard_office_provider as provider
from docwen_gui.clipboard_capture import FrozenClipboardCapture
from docwen_gui.clipboard_image_binding import bind_provider_images
from docwen_gui.clipboard_office_provider import WORD_EMBED_SOURCE_MIME, WPS_DOCUMENT_MIME, WPS_IMAGE_DATA_MIME
from docwen_gui.clipboard_plain_alignment import match_plain_table
from docwen_gui.clipboard_rich_document import project_frozen_rich_document
from docwen_gui.clipboard_structured import (
    project_structured_clipboard_html,
    project_structured_clipboard_html_with_resources,
)

pytestmark = pytest.mark.contract
_SAMPLES = Path(__file__).resolve().parents[4] / "tests/fixtures/files/clipboard-office"


def _text(blocks) -> str:
    values = []
    for block in blocks:
        assert isinstance(block, ClipboardParagraph)
        values.append(
            "".join(
                inline.value if isinstance(inline, ClipboardText) else "\n"
                for inline in block.inlines
                if isinstance(inline, (ClipboardText, ClipboardHardBreak))
            )
        )
    return "\n".join(values)


def _table(html: str) -> ClipboardTable:
    block = project_structured_clipboard_html(html.encode()).blocks[0]
    assert isinstance(block, ClipboardTable)
    return block


@pytest.mark.parametrize("producer", ["word", "wps"])
def test_derived_provider_binds_distinct_paragraphs_and_nested_cell(producer) -> None:
    # Compact HTML/plain are synthesized regression inputs, not native bytes.
    html = (_SAMPLES / "rich-derived.html").read_bytes()
    formats = (
        ((WORD_EMBED_SOURCE_MIME, (_SAMPLES / "word-derived.ole").read_bytes()),)
        if producer == "word"
        else (
            (WPS_DOCUMENT_MIME, (_SAMPLES / "wps-writer-derived.zip").read_bytes()),
            (WPS_IMAGE_DATA_MIME, (_SAMPLES / "wps-images.bin").read_bytes()),
        )
    )
    decision = project_frozen_rich_document(FrozenClipboardCapture(None, html, formats, None))
    assert decision.projection is not None and not decision.plain_fallback
    images = [
        inline
        for inline in iter_clipboard_inlines(decision.projection.document)
        if isinstance(inline, ClipboardImageRef)
    ]
    assert [image.alt for image in images] == ["alpha-first", "nested-blue", "alpha-repeat"]
    assert images[0].resource_id == images[2].resource_id != images[1].resource_id
    assert all(image.resource_id is not None for image in images)
    assert [(image.extent_cx_emu, image.extent_cy_emu) for image in images] == [
        (914400, 609600),
        (457200, 304800),
        (914400, 609600),
    ]


@pytest.mark.parametrize("newlines", ["\n", "\r\n"])
def test_full_grid_merge_and_quoted_multiline_keep_actual_cell_values(newlines: str) -> None:
    table = _table(
        '<table><tr><th colspan="2" id="heading">Heading</th><th>Tail</th></tr>'
        '<tr><td rowspan="2">00123</td><td>line one<br>\n  line two &quot;quoted&quot;</td><td></td></tr>'
        "<tr><td>&nbsp; spaces&nbsp; / NBSP:&nbsp;</td><td>=1+1</td></tr></table>"
    )
    plain = 'Heading\t\tTail\n00123\t"line one\nline two ""quoted"""\t\n\t  spaces  / NBSP:\u00a0\t=1+1'
    plain = plain.replace("\n", newlines)
    match = match_plain_table(table, plain)
    assert match is not None and (match.start, match.end) == (0, len(plain))
    cells = {(cell.row, cell.column): cell for cell in match.table.cells}
    assert cells[(0, 0)].header and cells[(0, 0)].html_id == "heading" and cells[(0, 0)].column_span == 2
    assert cells[(1, 0)].row_span == 2 and _text(cells[(1, 0)].blocks) == "00123"
    assert _text(cells[(1, 1)].blocks) == "line one" + newlines + 'line two "quoted"'
    assert _text(cells[(2, 1)].blocks) == "  spaces  / NBSP:\u00a0"
    assert _text(cells[(2, 2)].blocks) == "=1+1"


@pytest.mark.parametrize("grid", [False, True])
def test_nested_table_anchor_or_grid_and_original_padding(grid: bool) -> None:
    table = _table(
        '<table><tr><td colspan="2">A</td><td>Z</td></tr><tr><td><p>before</p>'
        "<table><tr><td>00123</td><td>  value </td></tr></table>"
        "<p>&nbsp;</p><p>after</p></td><td>line one<br>\n line two</td><td></td></tr></table>"
    )
    first = "A\t\tZ" if grid else "A\tZ"
    plain = first + "\nbefore\n00123\t  value \n \nafter\tline one\nline two\t"
    match = match_plain_table(table, plain)
    assert match is not None
    cell = next(cell for cell in match.table.cells if (cell.row, cell.column) == (1, 0))
    inner = cell.blocks[1]
    assert isinstance(inner, ClipboardTable)
    assert _text(inner.cells[1].blocks) == "  value "
    assert _text((cell.blocks[2],)) == " "


@pytest.mark.parametrize("plain", ["prefix A\tB suffix", "A\tB\nA\tB", "A\tB\nA\tB\ntrailing", '"A\tB'])
def test_unproven_or_repeated_table_representation_is_refused(plain: str) -> None:
    assert match_plain_table(_table("<table><tr><td>A</td><td>B</td></tr></table>"), plain) is None


def test_nonempty_covered_tsv_field_cannot_be_discarded() -> None:
    table = _table('<table><tr><td colspan="2">A</td><td>line one<br>line two</td></tr></table>')
    assert match_plain_table(table, 'A\tlost value\t"line one\nline two"') is None


def test_empty_table_keeps_captured_space_characters() -> None:
    match = match_plain_table(_table("<table><tr><td></td></tr></table>"), " \u00a0 ")
    assert match is not None
    assert _text(match.table.cells[0].blocks) == " \u00a0 "


def test_provider_two_inline_images_match_distinct_text_segments_without_ordinal_guess() -> None:
    xml = b"""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body><w:p>
        <w:r><w:t>left</w:t><w:br/><w:t>line</w:t></w:r>
        <w:r><w:drawing><wp:extent cx="914400" cy="609600"/><a:blip r:embed="rId7"/></w:drawing></w:r>
        <w:r><w:t>middle</w:t></w:r>
        <w:r><w:drawing><wp:extent cx="457200" cy="304800"/><a:blip r:embed="rId8"/></w:drawing></w:r>
        <w:r><w:t>right</w:t></w:r></w:p></w:body></w:document>"""
    occurrences = provider._collect_occurrences(xml, lambda value: value)
    assert [(item.before_text, item.after_text) for item in occurrences] == [
        ("left\nline", "middle"),
        ("middle", "right"),
    ]
    original = provider.parse_word_embed_source((_SAMPLES / "word-derived.ole").read_bytes())
    html = project_structured_clipboard_html(
        b'<p>left<br>line<img src="clipboard:first" alt="first">middle<img src="clipboard:second" alt="second">right</p>'
    )
    document, _resources = bind_provider_images(html, replace(original, occurrences=occurrences))
    images = [item for item in iter_clipboard_inlines(document) if isinstance(item, ClipboardImageRef)]
    assert [image.alt for image in images] == ["first", "second"]
    identities = dict(original.identifier_to_resource_id)
    assert [image.resource_id for image in images] == [identities["rId7"], identities["rId8"]]


def test_duplicate_neighbor_contexts_fail_instead_of_using_order_or_image_count() -> None:
    original = provider.parse_word_embed_source((_SAMPLES / "word-derived.ole").read_bytes())
    html = project_structured_clipboard_html(
        b"<p>BEFORE / 00123 / | &lt; ^ / alpha image follows</p>"
        b'<p><img src="clipboard:first"></p><p><img src="clipboard:second"></p>'
        b"<p>AFTER / repeated alpha image follows</p>"
    )
    with pytest.raises(provider.ClipboardOfficeProviderError) as failure:
        bind_provider_images(html, replace(original, occurrences=(original.occurrences[0],)))
    assert failure.value.code == "clipboard.provider_binding_invalid"


def test_unique_contexts_in_reversed_provider_order_are_refused() -> None:
    original = provider.parse_word_embed_source((_SAMPLES / "word-derived.ole").read_bytes())
    html = project_structured_clipboard_html_with_resources(
        (_SAMPLES / "rich-derived.html").read_bytes(), decode_inline_images=False
    ).document
    with pytest.raises(provider.ClipboardOfficeProviderError) as failure:
        bind_provider_images(html, replace(original, occurrences=tuple(reversed(original.occurrences))))
    assert failure.value.code == "clipboard.provider_binding_invalid"


@pytest.mark.parametrize("producer", ["word", "wps"])
@pytest.mark.parametrize("table", [False, True])
@pytest.mark.parametrize("plain", [None, "A"])
def test_provider_images_cannot_disappear_when_html_omits_all_image_nodes(producer, table, plain) -> None:
    formats = (
        ((WORD_EMBED_SOURCE_MIME, (_SAMPLES / "word-derived.ole").read_bytes()),)
        if producer == "word"
        else (
            (WPS_DOCUMENT_MIME, (_SAMPLES / "wps-writer-derived.zip").read_bytes()),
            (WPS_IMAGE_DATA_MIME, (_SAMPLES / "wps-images.bin").read_bytes()),
        )
    )
    html = b"<table><tr><td>A</td></tr></table>" if table else b"<p>A</p>"
    capture = FrozenClipboardCapture(plain, html, formats, None)
    if plain is None:
        from docwen_gui.clipboard_rich_document import ClipboardRichDocumentError

        with pytest.raises(ClipboardRichDocumentError) as failure:
            project_frozen_rich_document(capture)
        assert failure.value.code == "clipboard.provider_binding_invalid"
    else:
        decision = project_frozen_rich_document(capture)
        assert decision.handled and decision.plain_fallback and decision.projection is None


def test_plain_budget_precedes_rich_parsing_and_alignment_offsets(monkeypatch) -> None:
    from docwen_gui import clipboard_plain_alignment as alignment
    from docwen_gui import clipboard_rich_document as rich

    def forbidden(*_args, **_kwargs):
        pytest.fail("Over-budget plain text reached rich parsing or offset allocation")

    monkeypatch.setattr(alignment, "MAX_CLIPBOARD_TEXT_CODEPOINTS", 4)
    monkeypatch.setattr(rich, "MAX_CLIPBOARD_TEXT_CODEPOINTS", 4)
    table = _table("<table><tr><td>valid</td></tr></table>")
    monkeypatch.setattr(alignment, "canonical_clipboard_text", forbidden)
    monkeypatch.setattr(rich, "html_contains_table", forbidden)
    assert alignment.match_plain_table(table, "valid") is None
    decision = rich.project_frozen_rich_document(FrozenClipboardCapture("valid", b"<table>", (), None))
    assert decision.handled and decision.plain_fallback and decision.projection is None
