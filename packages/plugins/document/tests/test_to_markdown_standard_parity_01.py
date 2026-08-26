"""Focused tests split from test_to_markdown_standard_parity.py."""

from __future__ import annotations

import pytest

from ._to_markdown_standard_parity_support import (
    WD_COLOR_INDEX,
    WD_STYLE_TYPE,
    Document,
    DocxMarkdownSyntaxConfig,
    DocxToMarkdownConverter,
    MagicMock,
    NoteExtractor,
    OxmlElement,
    ParagraphStyle,
    Path,
    _convert_document_fixture_to_markdown,
    _inject_numpr,
    _inject_outline_level,
    cast,
    qn,
    re,
)

pytestmark = pytest.mark.contract


def test_note_extractor_reference_texts():
    extractor = NoteExtractor.__new__(NoteExtractor)
    extractor.footnotes = {7: "脚注内容"}
    extractor.endnotes = {8: "尾注内容"}
    extractor.footnote_id_map = {7: "1"}
    extractor.endnote_id_map = {8: "endnote:1"}

    assert extractor.get_reference_text("footnote", 7) == "[^1]"
    assert extractor.get_reference_text("endnote", 8) == "[^endnote:1]"
    assert "[^1]: 脚注内容" in extractor.build_definitions_block()
    assert "[^endnote:1]: 尾注内容" in extractor.build_definitions_block()


def test_hyperlink_uses_relationship_target_not_rid():
    doc = Document()
    para = doc.add_paragraph()
    rel_id = para.part.relate_to(
        "https://example.com/path",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "Example"
    run.append(text)
    hyperlink.append(run)
    para._p.append(hyperlink)

    from docwen_plugin_document.shared.markdown_runs import render_paragraph_runs

    rendered = render_paragraph_runs(para, syntax_config=DocxMarkdownSyntaxConfig())

    assert rendered == "[Example](https://example.com/path)"
    assert rel_id not in rendered


def test_revision_and_simple_field_xml_boundary_is_explicit():
    """Current renderer keeps inserted text, skips tracked deletions, and reads fldSimple display text.

    Old Tk/PySide6 fixtures treat comments/revisions/fields as broader final
    artifact scope, so this locks the current clean XML handling without
    claiming full Word review-feature parity.
    """
    doc = Document()
    para = doc.add_paragraph()
    para.add_run("Before ")

    inserted = OxmlElement("w:ins")
    inserted_run = OxmlElement("w:r")
    inserted_text = OxmlElement("w:t")
    inserted_text.text = "Inserted"
    inserted_run.append(inserted_text)
    inserted.append(inserted_run)
    para._p.append(inserted)

    deleted = OxmlElement("w:del")
    deleted_run = OxmlElement("w:r")
    deleted_text = OxmlElement("w:delText")
    deleted_text.text = "Deleted"
    deleted_run.append(deleted_text)
    deleted.append(deleted_run)
    para._p.append(deleted)

    moved_from = OxmlElement("w:moveFrom")
    moved_from_run = OxmlElement("w:r")
    moved_from_text = OxmlElement("w:t")
    moved_from_text.text = "MovedFrom"
    moved_from_run.append(moved_from_text)
    moved_from.append(moved_from_run)
    para._p.append(moved_from)

    moved_to = OxmlElement("w:moveTo")
    moved_to_run = OxmlElement("w:r")
    moved_to_text = OxmlElement("w:t")
    moved_to_text.text = "MovedTo"
    moved_to_run.append(moved_to_text)
    moved_to.append(moved_to_run)
    para._p.append(moved_to)

    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), 'HYPERLINK "https://example.com"')
    field_run = OxmlElement("w:r")
    field_text = OxmlElement("w:t")
    field_text.text = "FieldDisplay"
    field_run.append(field_text)
    field.append(field_run)
    para._p.append(field)

    from docwen_plugin_document.shared.markdown_runs import render_paragraph_runs

    rendered = render_paragraph_runs(para, syntax_config=DocxMarkdownSyntaxConfig())

    assert rendered == "Before InsertedMovedToFieldDisplay"
    assert "Deleted" not in rendered
    assert "MovedFrom" not in rendered
    assert "HYPERLINK" not in rendered


def test_visible_special_run_characters_and_hidden_text_follow_word_display_semantics() -> None:
    doc = Document()
    para = doc.add_paragraph()
    first = para.add_run("anti")
    first._r.append(OxmlElement("w:noBreakHyphen"))
    first.add_text("VEGF")
    second = para.add_run(" hidden")
    second._r.get_or_add_rPr().append(OxmlElement("w:vanish"))
    para.add_run(" treatment")

    from docwen_plugin_document.shared.markdown_runs import render_paragraph_runs

    assert (
        render_paragraph_runs(
            para,
            preserve_formatting=False,
            syntax_config=DocxMarkdownSyntaxConfig(),
        )
        == "anti‑VEGF treatment"
    )


def test_image_refs_merge_once_after_last_nonempty_rendered_line() -> None:
    """One element's image references must not be duplicated across every line."""
    lines = ["first", "", "second"]
    refs = ["![[image-a.png]]\n", "![[image-b.png]]\n"]

    merged = DocxToMarkdownConverter._merge_img_refs_into_lines(lines, refs)

    assert merged == ["first", "", "second![[image-a.png]]\n![[image-b.png]]\n"]


def test_image_only_element_still_emits_reachable_reference_line() -> None:
    """An empty image-owning paragraph/SDT must not leave a finalized orphan."""
    refs = ["![[image-only.png]]\n"]

    assert DocxToMarkdownConverter._merge_img_refs_into_lines([""], refs) == [
        "![[image-only.png]]\n",
        "",
    ]


def test_wrapped_revision_text_reaches_paragraph_processing() -> None:
    doc = Document()
    para = doc.add_paragraph()

    inserted = OxmlElement("w:ins")
    inserted_run = OxmlElement("w:r")
    inserted_text = OxmlElement("w:t")
    inserted_text.text = "Inserted only"
    inserted_run.append(inserted_text)
    inserted.append(inserted_run)
    para._p.append(inserted)

    deleted = OxmlElement("w:del")
    deleted_run = OxmlElement("w:r")
    deleted_text = OxmlElement("w:delText")
    deleted_text.text = "Deleted only"
    deleted_run.append(deleted_text)
    deleted.append(deleted_run)
    para._p.append(deleted)

    lines, stats = DocxToMarkdownConverter()._process_paragraph(
        para._p,
        {id(para._p): para},
    )

    assert lines == ["Inserted only", ""]
    assert stats == {"paragraphs": 1, "headings": 0}


def test_mixed_direct_and_inserted_revision_text_reaches_paragraph_processing() -> None:
    """A direct numbering run must not suppress a sibling accepted insertion."""
    doc = Document()
    para = doc.add_paragraph()
    para.add_run("2.\t")

    deleted = OxmlElement("w:del")
    deleted_run = OxmlElement("w:r")
    deleted_text = OxmlElement("w:delText")
    deleted_text.text = "Rejected heading"
    deleted_run.append(deleted_text)
    deleted.append(deleted_run)
    para._p.append(deleted)

    inserted = OxmlElement("w:ins")
    inserted_run = OxmlElement("w:r")
    inserted_text = OxmlElement("w:t")
    inserted_text.text = "Accepted heading"
    inserted_run.append(inserted_text)
    inserted.append(inserted_run)
    para._p.append(inserted)

    lines, stats = DocxToMarkdownConverter()._process_paragraph(
        para._p,
        {id(para._p): para},
    )

    assert lines == ["2.\tAccepted heading", ""]
    assert "Rejected heading" not in "\n".join(lines)
    assert stats == {"paragraphs": 1, "headings": 0}


def test_deleted_page_break_does_not_erase_accepted_inserted_paragraph() -> None:
    doc = Document()
    para = doc.add_paragraph()
    inserted = OxmlElement("w:ins")
    inserted_run = OxmlElement("w:r")
    inserted_text = OxmlElement("w:t")
    inserted_text.text = "Accepted field display"
    inserted_run.append(inserted_text)
    inserted.append(inserted_run)
    para._p.append(inserted)

    deleted = OxmlElement("w:del")
    deleted_run = OxmlElement("w:r")
    deleted_break = OxmlElement("w:br")
    deleted_break.set(qn("w:type"), "page")
    deleted_run.append(deleted_break)
    deleted.append(deleted_run)
    para._p.append(deleted)

    lines, stats = DocxToMarkdownConverter()._process_paragraph(
        para._p,
        {id(para._p): para},
    )

    assert lines == ["Accepted field display", ""]
    assert stats == {"paragraphs": 1, "headings": 0}


def test_remove_numbering_strips_heading_text_prefix():
    from docwen_core.text.heading_numbering import strip_heading_prefix

    rules = (("chinese_顿号", re.compile(r"^[一二三四五六七八九十百千万]+、"), 1),)
    numbering, text = strip_heading_prefix("一、总体要求", rules=rules)
    assert numbering == "一、"
    assert text == "总体要求"


def test_remove_numbering_preserves_heading_formatting_around_prefix():
    rules = (("chinese_顿号", re.compile(r"^[一二三四五六七八九十百千万]+、"), 1),)
    doc = Document()
    para = doc.add_paragraph(style="Heading 1")
    run = para.add_run("一、总体要求")
    run.bold = True
    converter = DocxToMarkdownConverter()
    converter._preserve_heading_formatting = True  # pyright: ignore[reportPrivateUsage]

    lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        remove_numbering=True,
        heading_cleanup_rules=rules,
    )

    assert lines == ["# **总体要求**", ""]


def test_docx_to_md_full_convert_uses_request_cleanup_rules(tmp_path):
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy
    from docwen_core.text.heading_numbering import (
        compile_clean_rules_from_data,
    )

    request_rules = compile_clean_rules_from_data(
        [{"id": "request", "enabled": True, "pattern": r"^REQ:\s*", "level": 1}]
    )
    doc = Document()
    doc.add_heading("REQ: Request title", level=1)
    doc.add_heading("GLOBAL: Global title", level=1)
    input_path = tmp_path / "request-cleanup.docx"
    doc.save(str(input_path))
    staging_dir = tmp_path / "staging-request-cleanup"
    staging_dir.mkdir()
    context = FakeExecutionContext(
        request=ConversionRequest(
            request_id="docx-request-cleanup",
            input_refs=[FileRef(path=str(input_path), format="docx", category="document")],
            target_format="md",
            options={"remove_numbering": True},
            output_policy=OutputPolicy(),
        ),
        workspace=FakeWorkspaceHandle(str(input_path), str(staging_dir)),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=None,
        heading_cleanup_rules=request_rules,
    )

    result = DocxToMarkdownConverter().convert(context)

    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "# Request title" in markdown
    assert "# GLOBAL: Global title" in markdown


def test_standard_table_uses_formatted_cell_text():
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    para = table.cell(0, 0).paragraphs[0]
    run = para.add_run("加粗")
    run.bold = True

    converter = DocxToMarkdownConverter()
    converter._preserve_table_header_formatting = True  # pyright: ignore[reportPrivateUsage]
    lines, image_count = converter._process_table(
        table._tbl,
        para_by_element={id(para._element): para},
    )

    assert "**加粗**" in "\n".join(lines)
    assert image_count == 0


def test_vertical_merge_continuation_does_not_add_a_phantom_column():
    """A vMerge continuation cell occupies its grid column; it is not extra."""
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Merged"
    table.cell(0, 1).text = "B"
    table.cell(1, 1).text = "C"
    table.cell(0, 0).merge(table.cell(1, 0))

    lines, image_count = DocxToMarkdownConverter()._process_table(
        table._tbl,
        table_merge_strategy="empty",
    )

    assert lines[:3] == [
        "| Merged | B |",
        "| --- | --- |",
        "|  | C |",
    ]
    assert image_count == 0


def test_inline_markdown_syntax_consumes_request_config():
    """DOCX->MD inline markers honor syntax frozen onto the converter request."""
    from docwen_core.docx_parsing.format_features import (
        docx_markdown_syntax_config_from_conversion_config,
    )

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("bold")
    run.bold = True
    para.add_run(" ")
    run = para.add_run("italic")
    run.italic = True
    para.add_run(" ")
    run = para.add_run("strike")
    run.font.strike = True
    para.add_run(" ")
    run = para.add_run("mark")
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    para.add_run(" ")
    run = para.add_run("sup")
    run.font.superscript = True
    para.add_run(" ")
    run = para.add_run("sub")
    run.font.subscript = True

    converter = DocxToMarkdownConverter()
    converter._syntax_config = docx_markdown_syntax_config_from_conversion_config(  # pyright: ignore[reportPrivateUsage]
        {
            "syntax": {
                "bold": "underscore",
                "italic": "underscore",
                "strikethrough": "html",
                "highlight": "html",
                "superscript": "html",
                "subscript": "html",
            }
        }
    )
    rendered = converter._extract_paragraph_text_formatted(para)  # pyright: ignore[reportPrivateUsage]

    assert rendered == ("__bold__ _italic_ <del>strike</del> <mark>mark</mark> <sup>sup</sup> <sub>sub</sub>")


def test_builtin_heading_styles_one_through_nine_export_matching_markdown_levels(tmp_path: Path):
    document = Document()
    for level in range(1, 10):
        document.add_heading(f"Heading {level}", level=level)

    markdown = _convert_document_fixture_to_markdown(tmp_path, document, request_id="heading-levels-1-9")

    for level in range(1, 10):
        assert f"{'#' * level} Heading {level}" in markdown


def test_inherited_outline_level_exports_the_effective_heading_level(tmp_path: Path):
    document = Document()
    inherited = cast(
        ParagraphStyle,
        document.styles.add_style("Inherited deep heading", WD_STYLE_TYPE.PARAGRAPH),
    )
    inherited.base_style = document.styles["Heading 8"]
    document.add_paragraph("Inherited heading", style=inherited)

    markdown = _convert_document_fixture_to_markdown(tmp_path, document, request_id="inherited-heading-level")

    assert "######## Inherited heading" in markdown


def test_word_body_text_outline_sentinel_does_not_become_h6():
    """Real government DOCX producers may emit outlineLvl=9 on every body paragraph."""
    doc = Document()
    para = doc.add_paragraph("ordinary government document body")
    _inject_outline_level(para, 9)

    lines, stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )

    assert lines[0] == "ordinary government document body"
    assert stats == {"paragraphs": 1, "headings": 0}


def test_real_outline_level_zero_still_becomes_h1():
    doc = Document()
    para = doc.add_paragraph("explicit outline heading")
    _inject_outline_level(para, 0)

    lines, stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )

    assert lines[0] == "# explicit outline heading"
    assert stats == {"paragraphs": 0, "headings": 1}


def test_standard_list_item_uses_explicit_marker_and_indent():
    """Standard DOCX list items honor the request-owned list syntax."""
    from docwen_plugin_document.shared.list_processing import ListCounterManager

    doc = Document()
    para = doc.add_paragraph("nested item")
    _inject_numpr(para, num_id="9", ilvl=2)

    level = MagicMock()
    level.num_fmt = "bullet"
    numbering_index = MagicMock()
    numbering_index.lookup.return_value = level
    numbering_index.lookup_by_style_id.return_value = None

    syntax_config = DocxMarkdownSyntaxConfig(unordered_list="plus", indent_spaces=2)
    converter = DocxToMarkdownConverter()
    converter._unordered_list_marker_type = syntax_config.unordered_list  # pyright: ignore[reportPrivateUsage]
    converter._list_indent_spaces = syntax_config.indent_spaces  # pyright: ignore[reportPrivateUsage]
    lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        numbering_index=numbering_index,
        list_counter=ListCounterManager(),
    )

    assert lines[0] == "    + nested item"


def test_docx_to_md_convert_smart_joins_adjacent_list_items(tmp_path: Path) -> None:
    """F-E1-011/F-E1-012: adjacent list items join, then separate from ordinary body."""
    document = Document()
    first = document.add_paragraph("list first")
    second = document.add_paragraph("list second")
    _inject_numpr(first, num_id="42", ilvl=0)
    _inject_numpr(second, num_id="42", ilvl=0)
    document.add_paragraph("after list")

    markdown = _convert_document_fixture_to_markdown(tmp_path, document, request_id="smart-join-list")

    assert "- list first\n- list second\n\nafter list" in markdown
    assert "- list first\n\n- list second" not in markdown


def test_docx_to_md_convert_smart_joins_adjacent_quote_blocks(tmp_path: Path) -> None:
    """F-E1-011/F-E1-013: adjacent quotes join, then separate from ordinary body."""
    document = Document()
    document.add_paragraph("quote first", style="Quote")
    document.add_paragraph("quote second", style="Quote")
    document.add_paragraph("after quote")

    markdown = _convert_document_fixture_to_markdown(tmp_path, document, request_id="smart-join-quote")

    assert "> quote first\n> quote second\n\nafter quote" in markdown
    assert "> quote first\n\n> quote second" not in markdown
