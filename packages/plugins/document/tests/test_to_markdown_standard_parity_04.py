"""Focused tests split from test_to_markdown_standard_parity.py."""

from __future__ import annotations

from ._to_markdown_standard_parity_support import (
    WD_BREAK,
    WD_STYLE_TYPE,
    Document,
    DocxToMarkdownConverter,
    MagicMock,
    OxmlElement,
    _inject_numpr,
    _inject_pPr_without_numPr,
    _make_pstyle_numbering_index,
    _mock_para_style,
    patch,
    pytest,
    qn,
)

pytestmark = pytest.mark.contract


def test_all_code_character_runs_become_configured_code_block():
    from docwen_core.docx_parsing.format_features import (
        CodeBlockAccumulator,
        StyleDetectorConfig,
    )

    doc = Document()
    character_style = "HTML Code"
    doc.styles.add_style(character_style, WD_STYLE_TYPE.CHARACTER)
    para = doc.add_paragraph()
    para.add_run("first").style = character_style
    para.add_run(" second").style = character_style
    accumulator = CodeBlockAccumulator()

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        code_block_acc=accumulator,
        style_detector_config=StyleDetectorConfig(
            code_character_style_names=frozenset({"HTML Code"}),
            code_full_paragraph_as_block=True,
        ),
    )

    assert lines == []
    assert accumulator.finalize() == "```\nfirst second\n```"


def test_mixed_character_styles_keep_inline_code_position():
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig

    doc = Document()
    character_style = "HTML Code"
    doc.styles.add_style(character_style, WD_STYLE_TYPE.CHARACTER)
    para = doc.add_paragraph()
    para.add_run("code").style = character_style
    para.add_run(" plain")

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        style_detector_config=StyleDetectorConfig(
            code_character_style_names=frozenset({"HTML Code"}),
            code_full_paragraph_as_block=True,
        ),
    )

    assert lines == ["`code` plain", ""]


def test_all_quote_character_runs_become_configured_quote_block():
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig

    doc = Document()
    character_style = "Quote Char"
    para = doc.add_paragraph()
    para.add_run("quoted").style = character_style

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        style_detector_config=StyleDetectorConfig(
            quote_character_style_names=frozenset({"Quote Char"}),
            quote_full_paragraph_as_block=True,
        ),
    )

    assert lines == ["> quoted", ""]


def test_code_block_style_inside_list_uses_list_context_indent():
    """Code blocks that are themselves list paragraphs keep old list-context indentation."""
    from docwen_core.docx_parsing.format_features import CodeBlockAccumulator

    doc = Document()
    para = doc.add_paragraph("print('nested')")
    _inject_numpr(para, num_id="7", ilvl=1)

    level = MagicMock()
    level.num_fmt = "bullet"
    numbering_index = MagicMock()
    numbering_index.lookup.return_value = level

    acc = CodeBlockAccumulator(indent_spaces=2)
    converter = DocxToMarkdownConverter()

    with _mock_para_style(para, "Code Block"):
        lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            numbering_index=numbering_index,
            code_block_acc=acc,
        )

    assert lines == []
    result = acc.finalize()
    assert result == "    ```\n    print('nested')\n    ```"


def test_quote_style_emits_blockquote():
    """Quote-style paragraph → ``>`` prefixed Markdown."""
    doc = Document()
    para = doc.add_paragraph("This is a quote")

    converter = DocxToMarkdownConverter()
    with _mock_para_style(para, "Quote 1"):
        lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
        )
    output = "\n".join(lines)
    assert output.startswith(">")
    assert "This is a quote" in output


def test_quote_style_nested_level():
    """Quote level 2 → ``>>`` prefix."""
    doc = Document()
    para = doc.add_paragraph("nested quote")

    converter = DocxToMarkdownConverter()
    with _mock_para_style(para, "Quote 2"):
        lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
        )
    output = "\n".join(lines)
    assert output.startswith(">>")
    assert "nested quote" in output


def test_multiline_quote_applies_level_marker_to_every_line():
    """A line break inside Quote 2 must retain the marker on both lines."""
    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("first line")
    run.add_break(WD_BREAK.LINE)
    run.add_text("second line")

    converter = DocxToMarkdownConverter()
    with _mock_para_style(para, "Quote 2"):
        lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
        )

    assert lines == [">> first line", ">> second line", ""]


def test_quote_style_inside_list_uses_list_context_indent():
    """Quote paragraphs that are themselves list paragraphs keep list-context indentation."""
    doc = Document()
    para = doc.add_paragraph("nested quote")
    _inject_numpr(para, num_id="8", ilvl=1)

    level = MagicMock()
    level.num_fmt = "bullet"
    numbering_index = MagicMock()
    numbering_index.lookup.return_value = level

    converter = DocxToMarkdownConverter()
    converter._list_indent_spaces = 2

    with _mock_para_style(para, "Quote 1"):
        lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            numbering_index=numbering_index,
        )

    assert lines[0] == "    > nested quote"


def test_shading_based_code_block():
    """Gray paragraph shading → paragraph treated as code block."""
    doc = Document()
    para = doc.add_paragraph("shaded code")
    # Inject gray paragraph shading
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "D9D9D9")
    pPr.append(shd)

    from docwen_core.docx_parsing.format_features import CodeBlockAccumulator

    acc = CodeBlockAccumulator()
    converter = DocxToMarkdownConverter()
    _lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        code_block_acc=acc,
    )

    assert acc.in_code_block
    result = acc.finalize()
    assert "shaded code" in (result or "")


@pytest.mark.parametrize(
    ("fill", "value", "config"),
    [
        ("D9D9D9", "clear", {"wps_shading_enabled": False}),
        ("FFFFFF", "pct15", {"word_shading_enabled": False}),
    ],
)
def test_paragraph_shading_switches_disable_each_producer(fill, value, config):
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig

    doc = Document()
    para = doc.add_paragraph("ordinary")
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    shading.set(qn("w:val"), value)
    para._p.get_or_add_pPr().append(shading)

    lines, _stats = DocxToMarkdownConverter()._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
        style_detector_config=StyleDetectorConfig(**config),
    )

    assert lines == ["ordinary", ""]


def test_normal_paragraph_output():
    """Normal paragraph without special styles → plain text output."""
    doc = Document()
    para = doc.add_paragraph("ordinary paragraph")

    converter = DocxToMarkdownConverter()
    with _mock_para_style(para, "Normal"):
        lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
        )
    output = "\n".join(lines)
    assert "ordinary paragraph" in output


def test_code_block_closes_before_normal_paragraph():
    """A normal paragraph after a code block emits the fenced code block."""
    from docwen_core.docx_parsing.format_features import CodeBlockAccumulator

    doc = Document()
    code_para = doc.add_paragraph("code line 1")
    normal_para = doc.add_paragraph("after code")

    converter = DocxToMarkdownConverter()
    acc = CodeBlockAccumulator()

    # First paragraph: code block style
    with _mock_para_style(code_para, "Code Block"):
        lines1, _ = converter._process_paragraph(
            code_para._element,
            para_by_element={id(code_para._element): code_para},
            code_block_acc=acc,
        )
    assert acc.in_code_block
    assert lines1 == [] or lines1 == [""]  # no visible output

    # Second paragraph: normal → should close code block
    with _mock_para_style(normal_para, "Normal"):
        lines2, _ = converter._process_paragraph(
            normal_para._element,
            para_by_element={id(normal_para._element): normal_para},
            code_block_acc=acc,
        )
    output = "\n".join(lines2)
    assert "```" in output
    assert "code line 1" in output
    assert not acc.in_code_block


def test_body_formatting_switch_discards_inline_markers():
    """Configured DOCX→MD body formatting discard removes inline Markdown markers."""
    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("body bold")
    run.bold = True

    converter = DocxToMarkdownConverter()
    converter._preserve_formatting = False  # pyright: ignore[reportPrivateUsage]

    lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )

    output = "\n".join(lines)
    assert "body bold" in output
    assert "**body bold**" not in output


def test_heading_formatting_switch_preserves_inline_markers():
    """Configured DOCX→MD heading formatting preserve keeps inline markers."""
    doc = Document()
    para = doc.add_paragraph()
    para.style = "Heading 1"
    run = para.add_run("heading bold")
    run.bold = True

    converter = DocxToMarkdownConverter()
    converter._preserve_heading_formatting = True  # pyright: ignore[reportPrivateUsage]

    lines, _stats = converter._process_paragraph(
        para._element,
        para_by_element={id(para._element): para},
    )

    output = "\n".join(lines)
    assert "# **heading bold**" in output


def test_table_header_formatting_switch_only_affects_header_row():
    """Table header formatting has its own switch separate from body cells."""
    doc = Document()
    table = doc.add_table(rows=2, cols=1)
    table.cell(0, 0).text = ""
    header_run = table.cell(0, 0).paragraphs[0].add_run("Header Bold")
    header_run.bold = True
    table.cell(1, 0).text = ""
    body_run = table.cell(1, 0).paragraphs[0].add_run("Body Bold")
    body_run.bold = True
    para_by_element = {
        id(paragraph._element): paragraph for row in table.rows for cell in row.cells for paragraph in cell.paragraphs
    }

    converter = DocxToMarkdownConverter()
    converter._preserve_formatting = True  # pyright: ignore[reportPrivateUsage]
    converter._preserve_table_header_formatting = False  # pyright: ignore[reportPrivateUsage]

    lines, _stats = converter._process_table(
        table._tbl,  # pyright: ignore[reportPrivateUsage]
        para_by_element=para_by_element,
    )

    output = "\n".join(lines)
    assert "| Header Bold |" in output
    assert "| **Header Bold** |" not in output
    assert "| **Body Bold** |" in output


def test_heading_with_pstyle_numbering_renders_prefix():
    """Heading whose style maps to a numbering definition via pStyle
    gets the resolved numbering prefix prepended (step 7)."""
    doc = Document()
    para = doc.add_paragraph("概述")

    num_idx = _make_pstyle_numbering_index(
        style_id="1Heading1",
        num_fmt="chineseCountingThousand",
        lvl_text="%1、",
    )

    with _mock_para_style(para, "Heading 1"):
        # Also mock style_id on the style object for pStyle lookup
        mock_style = MagicMock()
        mock_style.name = "Heading 1"
        mock_style.style_id = "1Heading1"
        with patch.object(type(para), "style", property(lambda s: mock_style)):
            converter = DocxToMarkdownConverter()
            from docwen_plugin_document.shared.list_processing import ListCounterManager

            lines, _stats = converter._process_paragraph(
                para._element,
                para_by_element={id(para._element): para},
                numbering_index=num_idx,
                list_counter=ListCounterManager(),
                remove_numbering=False,
            )

    output = "\n".join(lines)
    assert "一、" in output
    assert "概述" in output


def test_heading_pstyle_numbering_respects_remove_numbering():
    """When remove_numbering=True, pStyle prefix is stripped."""
    doc = Document()
    para = doc.add_paragraph("总体要求")

    num_idx = _make_pstyle_numbering_index(
        style_id="1Heading1",
        num_fmt="chineseCountingThousand",
        lvl_text="%1、",
    )

    with _mock_para_style(para, "Heading 1"):
        mock_style = MagicMock()
        mock_style.name = "Heading 1"
        mock_style.style_id = "1Heading1"
        with patch.object(type(para), "style", property(lambda s: mock_style)):
            converter = DocxToMarkdownConverter()
            from docwen_plugin_document.shared.list_processing import ListCounterManager

            lines, _stats = converter._process_paragraph(
                para._element,
                para_by_element={id(para._element): para},
                numbering_index=num_idx,
                list_counter=ListCounterManager(),
                remove_numbering=True,
            )

    output = "\n".join(lines)
    # Numbering prefix should be stripped by strip_heading_prefix
    assert "总体要求" in output


def test_regular_paragraph_pstyle_numbering_via_abs_num_id():
    """Regular paragraph whose style maps to numbering via pStyle
    gets the numbering prefix rendered (step 8, abs_ numId path)."""
    doc = Document()
    para = doc.add_paragraph("第一条内容")
    _inject_pPr_without_numPr(para)

    num_idx = _make_pstyle_numbering_index(
        style_id="1ListStyle",
        num_fmt="chineseCountingThousand",
        lvl_text="%1、",
        num_id="10",
        abs_id="0",
    )

    mock_style = MagicMock()
    mock_style.name = "Normal"
    mock_style.style_id = "1ListStyle"
    with patch.object(type(para), "style", property(lambda s: mock_style)):
        converter = DocxToMarkdownConverter()
        from docwen_plugin_document.shared.list_processing import ListCounterManager

        lines, _stats = converter._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            numbering_index=num_idx,
            list_counter=ListCounterManager(),
        )

    output = "\n".join(lines)
    assert "一、" in output
    assert "第一条内容" in output


def test_regular_paragraph_pstyle_numbering_respects_remove_numbering():
    doc = Document()
    para = doc.add_paragraph("第一条内容")
    _inject_pPr_without_numPr(para)
    num_idx = _make_pstyle_numbering_index(style_id="1ListStyle")

    mock_style = MagicMock()
    mock_style.name = "Normal"
    mock_style.style_id = "1ListStyle"
    with patch.object(type(para), "style", property(lambda s: mock_style)):
        from docwen_plugin_document.shared.list_processing import ListCounterManager

        lines, _stats = DocxToMarkdownConverter()._process_paragraph(
            para._element,
            para_by_element={id(para._element): para},
            numbering_index=num_idx,
            list_counter=ListCounterManager(),
            remove_numbering=True,
        )

    assert lines == ["第一条内容", ""]


def test_paragraph_pstyle_numbering_counter_increments():
    """Successive pStyle paragraphs increment the counter."""
    doc = Document()
    para1 = doc.add_paragraph("第一条")
    para2 = doc.add_paragraph("第二条")
    _inject_pPr_without_numPr(para1)
    _inject_pPr_without_numPr(para2)

    num_idx = _make_pstyle_numbering_index(
        style_id="1ListStyle",
        num_fmt="chineseCountingThousand",
        lvl_text="%1、",
    )

    mock_style1 = MagicMock()
    mock_style1.name = "Normal"
    mock_style1.style_id = "1ListStyle"
    mock_style2 = MagicMock()
    mock_style2.name = "Normal"
    mock_style2.style_id = "1ListStyle"

    converter = DocxToMarkdownConverter()
    from docwen_plugin_document.shared.list_processing import ListCounterManager

    lc = ListCounterManager()

    with patch.object(type(para1), "style", property(lambda s: mock_style1)):
        lines1, _ = converter._process_paragraph(
            para1._element,
            para_by_element={id(para1._element): para1},
            numbering_index=num_idx,
            list_counter=lc,
        )
    output1 = "\n".join(lines1)
    assert "一、" in output1

    with patch.object(type(para2), "style", property(lambda s: mock_style2)):
        lines2, _ = converter._process_paragraph(
            para2._element,
            para_by_element={id(para2._element): para2},
            numbering_index=num_idx,
            list_counter=lc,
        )
    output2 = "\n".join(lines2)
    assert "二、" in output2
