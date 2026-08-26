"""Focused tests split from test_to_markdown_standard_parity.py."""

from __future__ import annotations

from ._to_markdown_standard_parity_support import (
    _MC_NS,
    WD_BREAK,
    WD_STYLE_TYPE,
    Any,
    Document,
    DocxToMarkdownConverter,
    MagicMock,
    OxmlElement,
    Path,
    _append_alternate_content_formula,
    _append_formula,
    _inject_numpr,
    _inject_outline_level,
    _mock_para_style,
    _parse_markdown_ast,
    etree,
    nullcontext,
    pytest,
    qn,
)

pytestmark = pytest.mark.contract


def test_list_continuation_page_break_keeps_both_sides_in_parent_item():
    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        continuation_list_level=0,
    )

    assert lines == ["    Before", "", "    ---", "", "    After", ""]

    markdown = "- Parent\n\n" + "\n".join(lines)
    ast = _parse_markdown_ast(markdown)
    assert [node["type"] for node in ast] == ["list"]
    item_children = [node for node in ast[0]["children"][0]["children"] if node["type"] != "blank_line"]
    assert [node["type"] for node in item_children] == [
        "paragraph",
        "paragraph",
        "thematic_break",
        "paragraph",
    ]


def test_code_block_page_break_closes_and_reopens_fence():
    from docwen_core.docx_parsing.format_features import CodeBlockAccumulator

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")
    accumulator = CodeBlockAccumulator()

    with _mock_para_style(para, "Code Block"):
        lines, _stats = DocxToMarkdownConverter()._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            code_block_acc=accumulator,
        )

    assert lines == ["```\nBefore\n```", "", "---", ""]
    assert accumulator.finalize() == "```\nAfter\n```"


@pytest.mark.parametrize("mode", ("style", "shading", "full_run"))
def test_nested_code_page_break_indents_every_physical_line_for_all_detection_modes(mode: str):
    from docwen_core.docx_parsing.format_features import CodeBlockAccumulator, StyleDetectorConfig

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("One")
    run.add_break(WD_BREAK.LINE)
    run.add_text("Two")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("Three")
    run.add_break(WD_BREAK.LINE)
    run.add_text("Four")
    _inject_numpr(para, num_id="9", ilvl=0)
    level = MagicMock()
    level.num_fmt = "bullet"
    numbering_index = MagicMock()
    numbering_index.lookup.return_value = level
    numbering_index.lookup_by_style_id.return_value = None
    style_context = nullcontext()
    config = None
    if mode == "style":
        style_context = _mock_para_style(para, "Code Block")
    elif mode == "shading":
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "D9D9D9")
        shading.set(qn("w:val"), "clear")
        para._p.get_or_add_pPr().append(shading)
    else:
        doc.styles.add_style("HTML Code", WD_STYLE_TYPE.CHARACTER)
        run.style = "HTML Code"
        config = StyleDetectorConfig(code_character_style_names=frozenset({"HTML Code"}))
    accumulator = CodeBlockAccumulator()

    with style_context:
        lines, _stats = DocxToMarkdownConverter()._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            numbering_index=numbering_index,
            code_block_acc=accumulator,
            style_detector_config=config,
        )

    assert lines == ["    ```\n    One\n    Two\n    ```", "", "    ---", ""]
    assert accumulator.finalize() == "    ```\n    Three\n    Four\n    ```"


def test_quote_page_break_preserves_quote_structure_on_both_sides():
    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")

    with _mock_para_style(para, "Quote 2"):
        lines, _stats = DocxToMarkdownConverter()._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
        )

    assert lines == [">> Before", "", "---", "", ">> After", ""]


@pytest.mark.parametrize("separate_runs", (False, True))
def test_full_paragraph_quote_character_style_page_break_does_not_emit_code_ticks(separate_runs: bool):
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig

    doc = Document()
    doc.styles.add_style("DocWen Quote Char", WD_STYLE_TYPE.CHARACTER)
    para = doc.add_paragraph()
    before = para.add_run("Before")
    before.style = "DocWen Quote Char"
    before.add_break(WD_BREAK.PAGE)
    if separate_runs:
        after = para.add_run("After")
        after.style = "DocWen Quote Char"
    else:
        before.add_text("After")
    config = StyleDetectorConfig(quote_character_style_names=frozenset({"DocWen Quote Char"}))

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        style_detector_config=config,
    )

    assert lines == ["> Before", "", "---", "", "> After", ""]
    assert "`" not in "\n".join(lines)


@pytest.mark.parametrize("kind", ("normal", "heading", "list", "continuation", "code", "quote"))
def test_page_and_section_breaks_preserve_content_then_boundary_order(kind: str):
    from docwen_core.docx_parsing.format_features import CodeBlockAccumulator
    from docwen_plugin_document.shared.list_processing import ListCounterManager

    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")
    para._p.get_or_add_pPr().append(OxmlElement("w:sectPr"))
    kwargs: dict[str, Any] = {}
    style_context = nullcontext()
    accumulator = None
    if kind == "heading":
        _inject_outline_level(para, 0)
    elif kind == "list":
        _inject_numpr(para, num_id="9", ilvl=0)
        level = MagicMock()
        level.num_fmt = "bullet"
        numbering_index = MagicMock()
        numbering_index.lookup.return_value = level
        numbering_index.lookup_by_style_id.return_value = None
        kwargs.update(numbering_index=numbering_index, list_counter=ListCounterManager())
    elif kind == "continuation":
        kwargs["continuation_list_level"] = 0
    elif kind == "code":
        accumulator = CodeBlockAccumulator()
        kwargs["code_block_acc"] = accumulator
        style_context = _mock_para_style(para, "Code Block")
    elif kind == "quote":
        style_context = _mock_para_style(para, "Quote")

    with style_context:
        lines, _stats = DocxToMarkdownConverter()._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            **kwargs,
        )

    before_index = next(index for index, line in enumerate(lines) if "Before" in line)
    page_index = next(index for index, line in enumerate(lines) if line.strip() == "---")
    after_index = next(index for index, line in enumerate(lines) if "After" in line)
    section_index = next(index for index, line in enumerate(lines) if line.strip() == "***")
    assert before_index < page_index < after_index < section_index
    if kind == "heading":
        assert "# Before" in lines and "# After" in lines
    elif kind == "list":
        assert "- Before" in lines and "    After" in lines and "    ---" in lines
    elif kind == "continuation":
        assert "    Before" in lines and "    After" in lines and "    ---" in lines
    elif kind == "code":
        assert accumulator is not None and accumulator.in_code_block is False
        assert sum(line.count("```") for line in lines) == 4
    elif kind == "quote":
        assert "> Before" in lines and "> After" in lines


def test_paragraph_section_break_emits_separator():
    """Paragraph with w:sectPr → ``***`` separator in output."""
    doc = Document()
    para = doc.add_paragraph("section-end paragraph")
    # Inject sectPr
    sect_pr = OxmlElement("w:sectPr")
    para._p.append(sect_pr)

    converter = DocxToMarkdownConverter()
    lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )
    output = "\n".join(lines)
    assert "***" in output


def test_paragraph_section_break_ignore_omits_separator():
    """Ignored section breaks must not leak a sentinel/empty separator."""
    doc = Document()
    para = doc.add_paragraph("section-end paragraph")
    sect_pr = OxmlElement("w:sectPr")
    para._p.append(sect_pr)

    converter = DocxToMarkdownConverter()
    converter._section_break_separator = ""
    lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )
    output = "\n".join(lines)
    assert "section-end paragraph" in output
    assert "***" not in output
    assert "ignore" not in output


def test_paragraph_page_break_ignore_omits_separator():
    """Ignored page breaks split text without emitting a separator token."""
    doc = Document()
    para = doc.add_paragraph()
    para.add_run("before page break")
    run = para.add_run(" ")
    run.add_break(WD_BREAK.PAGE)
    para.add_run("after page break")

    converter = DocxToMarkdownConverter()
    converter._page_break_separator = ""
    lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )
    output = "\n".join(lines)
    assert "before page break" in output
    assert "after page break" in output
    assert "---" not in output
    assert "ignore" not in output


def test_docx_to_md_convert_consumes_request_break_separator_config(tmp_path):
    """Full convert path consumes request-owned export separators."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    doc = Document()
    para = doc.add_paragraph()
    para.add_run("before")
    run = para.add_run(" ")
    run.add_break(WD_BREAK.PAGE)
    para.add_run("after")
    input_path = tmp_path / "page_break.docx"
    doc.save(str(input_path))

    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()
    config_snapshot = {
        "conversion": {
            "horizontal_rule": {
                "docx_to_md": {
                    "page_break": "___",
                    "section_break": "***",
                    "horizontal_rule": "---",
                }
            }
        }
    }
    ctx = FakeExecutionContext(
        request=ConversionRequest(
            request_id="break-separator-test",
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

    result = DocxToMarkdownConverter().convert(ctx)
    assert result.success, result.error.message if result.error else "conversion failed"
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "before" in content
    assert "after" in content
    assert "___" in content


def test_standalone_omath_paragraph_without_omathpara_emits_block_formula():
    """Legacy Word producers may store a display formula as a bare oMath paragraph."""
    doc = Document()
    para = doc.add_paragraph()
    omath = OxmlElement("m:oMath")
    math_run = OxmlElement("m:r")
    math_text = OxmlElement("m:t")
    math_text.text = "x"
    math_run.append(math_text)
    omath.append(math_run)
    para._p.append(omath)

    converter = DocxToMarkdownConverter()
    output = converter._build_paragraph_text_with_formulas(para._element, para)

    assert output == "$$x$$"


def test_mixed_text_and_omath_paragraph_remains_inline_formula():
    doc = Document()
    para = doc.add_paragraph("before ")
    omath = OxmlElement("m:oMath")
    math_run = OxmlElement("m:r")
    math_text = OxmlElement("m:t")
    math_text.text = "x"
    math_run.append(math_text)
    omath.append(math_run)
    para._p.append(omath)
    para.add_run(" after")

    converter = DocxToMarkdownConverter()
    output = converter._build_paragraph_text_with_formulas(para._element, para)

    assert output == "before $x$ after"


def test_alternate_content_formula_prefers_choice_without_duplicate_fallback():
    doc = Document()
    para = doc.add_paragraph("before ")
    _append_alternate_content_formula(
        para,
        choice_text="CHOICE_SENTINEL",
        fallback_text="FALLBACK_SENTINEL",
    )
    para.add_run(" after")

    output = DocxToMarkdownConverter()._build_paragraph_text_with_formulas(para._element, para)

    assert output == "before $CHOICE_SENTINEL$ after"
    assert "FALLBACK_SENTINEL" not in output


def test_alternate_content_formula_uses_fallback_when_choice_is_not_renderable():
    doc = Document()
    para = doc.add_paragraph("before ")
    _append_alternate_content_formula(
        para,
        choice_text=None,
        fallback_text="FALLBACK_SENTINEL",
    )
    para.add_run(" after")

    output = DocxToMarkdownConverter()._build_paragraph_text_with_formulas(para._element, para)

    assert output == "before $FALLBACK_SENTINEL$ after"


def test_alternate_content_formula_uses_fallback_for_empty_choice_hyperlink():
    doc = Document()
    para = doc.add_paragraph("before ")
    rel_id = para.part.relate_to(
        "https://example.com/empty-choice",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    alternate = etree.Element(f"{{{_MC_NS}}}AlternateContent", nsmap={"mc": _MC_NS})
    choice = etree.SubElement(alternate, f"{{{_MC_NS}}}Choice")
    choice.set("Requires", "w14")
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    choice.append(hyperlink)
    fallback = etree.SubElement(alternate, f"{{{_MC_NS}}}Fallback")
    _append_formula(fallback, "FALLBACK_SENTINEL")
    para._p.append(alternate)
    para.add_run(" after")

    output = DocxToMarkdownConverter()._build_paragraph_text_with_formulas(para._element, para)

    assert output == "before $FALLBACK_SENTINEL$ after"
    assert "empty-choice" not in output


def test_page_break_renderer_keeps_selected_alternate_content_formula():
    doc = Document()
    para = doc.add_paragraph("before ")
    _append_alternate_content_formula(
        para,
        choice_text="CHOICE_SENTINEL",
        fallback_text="FALLBACK_SENTINEL",
    )
    page_break = para.add_run()
    page_break.add_break(WD_BREAK.PAGE)
    para.add_run("after")

    parts = DocxToMarkdownConverter()._render_page_break_parts(para)

    assert parts == [
        ("text", "before $CHOICE_SENTINEL$"),
        ("separator", "---"),
        ("text", "after"),
    ]
    assert "FALLBACK_SENTINEL" not in repr(parts)


def test_full_convert_preserves_one_effective_alternate_content_formula(tmp_path: Path) -> None:
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    document = Document()
    paragraph = document.add_paragraph()
    _append_formula(paragraph._p, "DIRECT_SENTINEL")
    _append_alternate_content_formula(
        paragraph,
        choice_text="CHOICE_SENTINEL",
        fallback_text="FALLBACK_SENTINEL",
    )
    source = tmp_path / "alternate-content-formula.docx"
    document.save(str(source))
    staging = tmp_path / "staging"
    staging.mkdir()
    context = FakeExecutionContext(
        request=ConversionRequest(
            request_id="alternate-content-formula",
            input_refs=[FileRef(path=str(source), format="docx", category="document")],
            target_format="md",
            options={},
            output_policy=OutputPolicy(),
        ),
        workspace=FakeWorkspaceHandle(str(source), str(staging)),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=None,
    )

    result = DocxToMarkdownConverter().convert(context)

    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert markdown.count("DIRECT_SENTINEL") == 1
    assert markdown.count("CHOICE_SENTINEL") == 1
    assert "FALLBACK_SENTINEL" not in markdown


def test_code_block_style_accumulates_lines():
    """Paragraph with code-block style → lines accumulate in CodeBlockAccumulator."""
    from docwen_core.docx_parsing.format_features import CodeBlockAccumulator

    doc = Document()
    para = doc.add_paragraph("print('hello')")

    acc = CodeBlockAccumulator()
    converter = DocxToMarkdownConverter()

    with _mock_para_style(para, "Code Block"):
        _lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            code_block_acc=acc,
        )

    # Code block paragraphs should NOT emit lines directly
    assert acc.in_code_block
    # Finalize and check content
    result = acc.finalize()
    assert result is not None
    assert "print('hello')" in result
    assert "```" in result


def test_code_block_style_consumes_explicit_document_style_config():
    """DOCX→MD uses the style aliases owned by the active request."""
    from docwen_core.docx_parsing.format_features import (
        CodeBlockAccumulator,
        StyleDetectorConfig,
    )

    doc = Document()
    para = doc.add_paragraph("configured code")

    acc = CodeBlockAccumulator()
    converter = DocxToMarkdownConverter()

    with _mock_para_style(para, "Runtime Code Alias"):
        _lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            code_block_acc=acc,
            style_detector_config=StyleDetectorConfig(code_block_style_fragments=("Runtime Code Alias",)),
        )

    assert acc.in_code_block
    result = acc.finalize()
    assert result is not None
    assert "configured code" in result
