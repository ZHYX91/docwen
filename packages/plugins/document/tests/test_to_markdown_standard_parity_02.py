"""Focused tests split from test_to_markdown_standard_parity.py."""

from __future__ import annotations

from ._to_markdown_standard_parity_support import (
    WD_BREAK,
    WD_STYLE_TYPE,
    Any,
    Document,
    DocxToMarkdownConverter,
    MagicMock,
    OxmlElement,
    Path,
    Twips,
    _inject_numpr,
    _inject_outline_level,
    _inject_pPr_without_numPr,
    _make_pstyle_numbering_index,
    _parse_markdown_ast,
    patch,
    pytest,
    qn,
    tomllib,
)

pytestmark = pytest.mark.contract


def test_docx_to_md_convert_consumes_request_list_syntax_config(tmp_path):
    """Full convert path reads list syntax from the admitted request snapshot."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    doc = Document()
    doc.add_paragraph("body")
    input_path = tmp_path / "list_syntax.docx"
    doc.save(str(input_path))
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()

    config_snapshot = {"conversion": {"syntax": {"unordered_list": "plus", "indent_spaces": 2}}}
    ctx = FakeExecutionContext(
        request=ConversionRequest(
            request_id="list-syntax-test",
            input_refs=[FileRef(path=str(input_path), format="docx", category="document")],
            target_format="md",
            options={},
            output_policy=OutputPolicy(),
            config_snapshot=config_snapshot,
        ),
        workspace=FakeWorkspaceHandle(str(input_path), str(staging_dir)),
        config=FakeConfigView(config_snapshot),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=None,
    )

    captured: dict[str, Any] = {}

    def _fake_parse(self: DocxToMarkdownConverter, *_args: Any, **_kwargs: Any) -> tuple[str, dict[str, int]]:
        captured["unordered_list"] = self._unordered_list_marker_type  # pyright: ignore[reportPrivateUsage]
        captured["indent_spaces"] = self._list_indent_spaces  # pyright: ignore[reportPrivateUsage]
        return "body", {"paragraphs": 1, "headings": 0, "tables": 0, "images": 0}

    with patch.object(DocxToMarkdownConverter, "_parse_docx", _fake_parse):
        result = DocxToMarkdownConverter().convert(ctx)

    assert result.success, result.error.message if result.error else "conversion failed"
    assert captured == {"unordered_list": "plus", "indent_spaces": 2}


def test_docx_to_md_convert_consumes_request_style_alias_overrides(tmp_path):
    """Plugin-level style alias request options should override runtime style config."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    doc = Document()
    doc.styles.add_style("My Direct Code Alias Style", WD_STYLE_TYPE.PARAGRAPH)
    doc.styles.add_style("My Direct Quote Alias Style", WD_STYLE_TYPE.PARAGRAPH)
    doc.add_paragraph("print('request')", style="My Direct Code Alias Style")
    doc.add_paragraph("quoted by request", style="My Direct Quote Alias Style")

    input_path = tmp_path / "request_style_alias.docx"
    doc.save(str(input_path))
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()

    ctx = FakeExecutionContext(
        request=ConversionRequest(
            request_id="request-style-alias-test",
            input_refs=[FileRef(path=str(input_path), format="docx", category="document")],
            target_format="md",
            options={
                "code_block_style_aliases": ["Direct Code Alias"],
                "quote_style_aliases": ["Direct Quote Alias"],
            },
            output_policy=OutputPolicy(),
        ),
        workspace=FakeWorkspaceHandle(str(input_path), str(staging_dir)),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=None,
    )

    result = DocxToMarkdownConverter().convert(ctx)

    assert result.success, result.error.message if result.error else "conversion failed"
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "```\nprint('request')\n```" in content
    assert "> quoted by request" in content


def test_docx_to_md_convert_uses_every_bundled_locale_template_style(tmp_path):
    """Bundled localized styles must survive the real DOCX-to-Markdown path."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    repo_root = Path(__file__).resolve().parents[4]
    locale_paths = sorted((repo_root / "i18n" / "locales").glob("*.toml"))
    template_paths = sorted((repo_root / "templates").glob("*.docx"))
    assert len(locale_paths) == 11
    assert len(template_paths) == 11

    template_style_names = {path: {style.name for style in Document(str(path)).styles} for path in template_paths}
    for locale_path in locale_paths:
        styles = tomllib.loads(locale_path.read_text(encoding="utf-8"))["styles"]
        code_name = styles["code_block"]
        matching_templates = [path for path, names in template_style_names.items() if code_name in names]
        assert len(matching_templates) == 1, (locale_path.name, code_name, matching_templates)

        locale_id = locale_path.stem
        code_text = f"print('{locale_id}')"
        quote_text = f"quote-{locale_id}"
        document = Document(str(matching_templates[0]))
        document.add_paragraph(code_text, style=code_name)
        document.add_paragraph(quote_text, style=styles["quote_3"])

        input_path = tmp_path / f"localized-{locale_id}.docx"
        document.save(str(input_path))
        staging_dir = tmp_path / f"staging-{locale_id}"
        staging_dir.mkdir()
        context = FakeExecutionContext(
            request=ConversionRequest(
                request_id=f"localized-style-{locale_id}",
                input_refs=[FileRef(path=str(input_path), format="docx", category="document")],
                target_format="md",
                options={},
                output_policy=OutputPolicy(),
            ),
            workspace=FakeWorkspaceHandle(str(input_path), str(staging_dir)),
            config=FakeConfigView(),
            progress=FakeProgressSink(),
            cancellation=CancellationToken().view(),
            logger=FakePluginLogger(),
            numbering_registry=None,
        )

        result = DocxToMarkdownConverter().convert(context)

        assert result.success, (
            locale_path.name,
            result.error.message if result.error else "conversion failed",
        )
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert f"```\n{code_text}\n```" in content, (locale_path.name, code_name, content)
        assert f">>> {quote_text}" in content, (locale_path.name, styles["quote_3"], content)


def test_docx_to_md_convert_indents_extra_indent_list_continuation(tmp_path):
    """Extra-indented paragraph immediately after a list item remains inside that list item."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    doc = Document()
    list_para = doc.add_paragraph("parent item")
    _inject_numpr(list_para, num_id="42", ilvl=0)
    continuation = doc.add_paragraph("continued block")
    continuation.paragraph_format.left_indent = Twips(900)

    input_path = tmp_path / "list_continuation.docx"
    doc.save(str(input_path))
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()

    ctx = FakeExecutionContext(
        request=ConversionRequest(
            request_id="list-continuation-test",
            input_refs=[FileRef(path=str(input_path), format="docx", category="document")],
            target_format="md",
            options={},
            output_policy=OutputPolicy(),
            config_snapshot={"conversion": {"syntax": {"indent_spaces": 2}}},
        ),
        workspace=FakeWorkspaceHandle(str(input_path), str(staging_dir)),
        config=FakeConfigView({"conversion": {"syntax": {"indent_spaces": 2}}}),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=None,
    )

    result = DocxToMarkdownConverter().convert(ctx)

    assert result.success, result.error.message if result.error else "conversion failed"
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "- parent item" in content
    assert "  continued block" in content


def test_docx_to_md_convert_indents_extra_indent_table_after_list(tmp_path):
    """Extra-indented table immediately after a list item remains inside that list item."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    doc = Document()
    list_para = doc.add_paragraph("parent item")
    _inject_numpr(list_para, num_id="43", ilvl=0)
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "nested cell"
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tbl_pr)
    tbl_ind = OxmlElement("w:tblInd")
    tbl_ind.set(qn("w:w"), "900")
    tbl_ind.set(qn("w:type"), "dxa")
    tbl_pr.append(tbl_ind)

    input_path = tmp_path / "list_table_continuation.docx"
    doc.save(str(input_path))
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()

    ctx = FakeExecutionContext(
        request=ConversionRequest(
            request_id="list-table-continuation-test",
            input_refs=[FileRef(path=str(input_path), format="docx", category="document")],
            target_format="md",
            options={},
            output_policy=OutputPolicy(),
            config_snapshot={"conversion": {"syntax": {"indent_spaces": 2}}},
        ),
        workspace=FakeWorkspaceHandle(str(input_path), str(staging_dir)),
        config=FakeConfigView({"conversion": {"syntax": {"indent_spaces": 2}}}),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=None,
    )

    result = DocxToMarkdownConverter().convert(ctx)

    assert result.success, result.error.message if result.error else "conversion failed"
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "- parent item" in content
    assert "  | nested cell |" in content
    assert "  | --- |" in content


def test_sdt_internal_list_context_indents_continuation_and_table():
    """List context inside one SDT content block applies to continuation paragraphs and tables."""
    from docwen_plugin_document.shared.list_processing import ListCounterManager

    doc = Document()
    list_para = doc.add_paragraph("parent item")
    _inject_numpr(list_para, num_id="44", ilvl=0)
    continuation = doc.add_paragraph("continued block")
    continuation.paragraph_format.left_indent = Twips(900)
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "sdt cell"
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tbl_pr)
    tbl_ind = OxmlElement("w:tblInd")
    tbl_ind.set(qn("w:w"), "900")
    tbl_ind.set(qn("w:type"), "dxa")
    tbl_pr.append(tbl_ind)

    sdt_content = OxmlElement("w:sdtContent")
    sdt_content.append(list_para._element)
    sdt_content.append(continuation._element)
    sdt_content.append(table._tbl)

    level = MagicMock()
    level.num_fmt = "bullet"
    numbering_index = MagicMock()
    numbering_index.lookup.return_value = level
    numbering_index.lookup_by_style_id.return_value = None

    para_by_element = {
        id(list_para._element): list_para,
        id(continuation._element): continuation,
        id(table.cell(0, 0).paragraphs[0]._element): table.cell(0, 0).paragraphs[0],
    }
    converter = DocxToMarkdownConverter()
    converter._list_indent_spaces = 2
    lines, _stats = converter._process_sdt(
        sdt_content,
        para_by_element=para_by_element,
        numbering_index=numbering_index,
        list_counter=ListCounterManager(),
    )

    output = "\n".join(lines)
    assert "- parent item" in output
    assert "  continued block" in output
    assert "  | sdt cell |" in output
    assert "  | --- |" in output


def test_paragraph_page_break_emits_separator():
    """Page break in a standard DOCX paragraph → ``---`` separator emitted."""
    doc = Document()
    para = doc.add_paragraph()
    para.add_run("before page break")
    run = para.add_run(" ")
    run.add_break(WD_BREAK.PAGE)
    para.add_run("after page break")

    converter = DocxToMarkdownConverter()
    lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )
    assert lines == ["before page break", "", "---", "", "after page break", ""]

    ast = _parse_markdown_ast("\n".join(lines))
    assert [node["type"] for node in ast] == ["paragraph", "thematic_break", "paragraph"]


def test_heading_page_break_preserves_heading_structure_on_both_sides():
    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")
    _inject_outline_level(para, 0)

    lines, stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )

    assert lines == ["# Before", "", "---", "", "# After", ""]
    assert stats == {"paragraphs": 0, "headings": 1}


def test_standard_list_page_break_keeps_one_marker_and_continuation_indent():
    from docwen_plugin_document.shared.list_processing import ListCounterManager

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")
    _inject_numpr(para, num_id="9", ilvl=0)
    level = MagicMock()
    level.num_fmt = "bullet"
    numbering_index = MagicMock()
    numbering_index.lookup.return_value = level
    numbering_index.lookup_by_style_id.return_value = None

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        numbering_index=numbering_index,
        list_counter=ListCounterManager(),
    )

    assert lines == ["- Before", "", "    ---", "", "    After", ""]

    ast = _parse_markdown_ast("\n".join(lines))
    assert [node["type"] for node in ast] == ["list"]
    item_children = [node for node in ast[0]["children"][0]["children"] if node["type"] != "blank_line"]
    assert [node["type"] for node in item_children] == ["paragraph", "thematic_break", "paragraph"]


def test_standard_list_page_break_supports_wide_ordered_marker():
    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")
    _inject_numpr(para, num_id="9", ilvl=0)
    level = MagicMock()
    level.num_fmt = "decimal"
    numbering_index = MagicMock()
    numbering_index.lookup.return_value = level
    numbering_index.lookup_by_style_id.return_value = None
    list_counter = MagicMock()
    list_counter.next.return_value = 100

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        numbering_index=numbering_index,
        list_counter=list_counter,
    )

    assert lines == ["100. Before", "", "     ---", "", "     After", ""]

    ast = _parse_markdown_ast("\n".join(lines))
    assert [node["type"] for node in ast] == ["list"]
    item_children = [node for node in ast[0]["children"][0]["children"] if node["type"] != "blank_line"]
    assert [node["type"] for node in item_children] == ["paragraph", "thematic_break", "paragraph"]


def test_standard_list_page_break_before_text_does_not_invent_an_empty_item():
    from docwen_plugin_document.shared.list_processing import ListCounterManager

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run()
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")
    _inject_numpr(para, num_id="9", ilvl=0)
    level = MagicMock()
    level.num_fmt = "bullet"
    numbering_index = MagicMock()
    numbering_index.lookup.return_value = level
    numbering_index.lookup_by_style_id.return_value = None

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        numbering_index=numbering_index,
        list_counter=ListCounterManager(),
    )

    assert lines == ["---", "", "- After", ""]

    ast = _parse_markdown_ast("\n".join(lines))
    assert [node["type"] for node in ast] == ["thematic_break", "list"]


def test_pstyle_list_page_break_keeps_separator_inside_same_list_item():
    from docwen_plugin_document.shared.list_processing import ListCounterManager

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")
    _inject_pPr_without_numPr(para)
    numbering_index = _make_pstyle_numbering_index(
        style_id="1ListStyle",
        num_fmt="decimal",
        lvl_text="%1.",
    )
    mock_style = MagicMock(name="Normal")
    mock_style.name = "Normal"
    mock_style.style_id = "1ListStyle"

    with patch.object(type(para), "style", property(lambda _self: mock_style)):
        lines, _stats = DocxToMarkdownConverter()._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            numbering_index=numbering_index,
            list_counter=ListCounterManager(),
        )

    assert lines == ["1. Before", "", "   ---", "", "   After", ""]

    ast = _parse_markdown_ast("\n".join(lines))
    item_children = [node for node in ast[0]["children"][0]["children"] if node["type"] != "blank_line"]
    assert [node["type"] for node in item_children] == ["paragraph", "thematic_break", "paragraph"]


@pytest.mark.parametrize("with_section", (False, True))
def test_non_commonmark_pstyle_page_break_never_turns_tail_into_code(with_section: bool):
    from docwen_plugin_document.shared.list_processing import ListCounterManager

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")
    _inject_pPr_without_numPr(para)
    if with_section:
        para._p.get_or_add_pPr().append(OxmlElement("w:sectPr"))
    numbering_index = _make_pstyle_numbering_index(
        style_id="ArticleListStyle",
        num_fmt="decimal",
        lvl_text="Article %1.",
    )
    mock_style = MagicMock(name="Normal")
    mock_style.name = "Normal"
    mock_style.style_id = "ArticleListStyle"

    with patch.object(type(para), "style", property(lambda _self: mock_style)):
        lines, _stats = DocxToMarkdownConverter()._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            numbering_index=numbering_index,
            list_counter=ListCounterManager(),
        )

    expected = ["Article 1. Before", "", "---", "", "After"]
    if with_section:
        expected.extend(["", "***"])
    expected.append("")
    assert lines == expected

    ast = _parse_markdown_ast("\n".join(lines))
    expected_types = ["paragraph", "thematic_break", "paragraph"]
    if with_section:
        expected_types.append("thematic_break")
    assert [node["type"] for node in ast] == expected_types
    assert "block_code" not in {node["type"] for node in ast}
