import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_core.docx_parsing.format_features import (
    DocxMarkdownSyntaxConfig,
    StyleDetectorConfig,
)
from docwen_plugin_document.shared.markdown_runs import (
    append_formatted_run_text,
    render_paragraph_runs,
    render_paragraph_runs_split_on_page_breaks,
    resolve_hyperlink_target,
)

pytestmark = pytest.mark.unit


def test_resolve_hyperlink_target_returns_none_for_missing_rel():
    doc = Document()
    para = doc.add_paragraph("plain")

    class Link:
        def get(self, key, default=None):
            return "rId999"

    assert resolve_hyperlink_target(para, Link()) is None


def test_resolve_hyperlink_target_resolves_real_url():
    doc = Document()
    para = doc.add_paragraph()
    rel_id = para.part.relate_to(
        "https://example.com/path",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    # Simulate a hyperlink XML element with r:id=rel_id
    from lxml import etree

    hl = etree.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink")
    hl.set(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
        rel_id,
    )

    result = resolve_hyperlink_target(para, hl)
    assert result == "https://example.com/path"


@pytest.mark.parametrize(
    ("apply_format", "expected"),
    [
        (lambda run: setattr(run, "bold", True), "__value__"),
        (lambda run: setattr(run, "italic", True), "_value_"),
        (lambda run: setattr(run, "underline", True), "<u>value</u>"),
        (lambda run: setattr(run.font, "strike", True), "<del>value</del>"),
        (lambda run: setattr(run.font, "superscript", True), "^value^"),
        (lambda run: setattr(run.font, "subscript", True), "~value~"),
        (
            lambda run: setattr(run.font, "highlight_color", WD_COLOR_INDEX.YELLOW),
            "<mark>value</mark>",
        ),
    ],
)
def test_append_formatted_run_text_honors_explicit_syntax_config(apply_format, expected):
    doc = Document()
    run = doc.add_paragraph().add_run("value")
    apply_format(run)
    syntax = DocxMarkdownSyntaxConfig(
        bold="underscore",
        italic="underscore",
        strikethrough="html",
        highlight="html",
        superscript="extended",
        subscript="extended",
    )
    parts: list[str] = []

    append_formatted_run_text(parts, run.text, run._r, syntax_config=syntax)

    assert "".join(parts) == expected


def test_inline_code_uses_safe_backtick_delimiter_and_coalesces_adjacent_runs():
    doc = Document()
    paragraph = doc.add_paragraph()
    first = paragraph.add_run("a`")
    second = paragraph.add_run("b")
    for run in (first, second):
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "D9D9D9")
        run._r.get_or_add_rPr().append(shading)

    rendered = render_paragraph_runs(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == "``a`b``"


def test_inline_code_with_boundary_backticks_gets_commonmark_padding():
    doc = Document()
    run = doc.add_paragraph().add_run("`value`")
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), "D9D9D9")
    run._r.get_or_add_rPr().append(shading)

    rendered = render_paragraph_runs(
        doc.paragraphs[0],
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == "`` `value` ``"


def test_append_formatted_run_text_preserves_gray_shading_as_inline_code():
    doc = Document()
    run = doc.add_paragraph().add_run("value")
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), "D9D9D9")
    run._r.get_or_add_rPr().append(shading)
    parts: list[str] = []

    append_formatted_run_text(
        parts,
        run.text,
        run._r,
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert "".join(parts) == "`value`"


def test_render_paragraph_runs_coalesces_adjacent_runs_through_shared_helper():
    doc = Document()
    paragraph = doc.add_paragraph()
    first = paragraph.add_run("Hello ")
    first.bold = True
    second = paragraph.add_run("World")
    second.bold = True

    rendered = render_paragraph_runs(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(bold="underscore"),
    )

    assert rendered == "__Hello World__"


@pytest.mark.parametrize("marker", ["asterisk", "underscore"])
def test_adjacent_different_run_formats_use_unambiguous_boundaries(marker):
    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run("Bold").bold = True
    paragraph.add_run("Italic").italic = True
    syntax = DocxMarkdownSyntaxConfig(bold=marker, italic=marker)
    actual = render_paragraph_runs(paragraph, syntax_config=syntax)
    assert actual == "<strong>Bold</strong><em>Italic</em>"


def test_literal_code_like_text_does_not_coalesce_with_a_styled_code_run():
    doc = Document()
    paragraph = doc.add_paragraph("`literal`")
    run = paragraph.add_run("actual code")
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), "D9D9D9")
    run._r.get_or_add_rPr().append(shading)
    assert render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig()) == r"\`literal\``actual code`"


def test_page_walker_keeps_text_and_soft_breaks_around_multiple_breaks_in_one_run():
    doc = Document()
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("first")
    run.bold = True
    run.add_break()
    run.add_text("second")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("third")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("fourth")

    assert render_paragraph_runs_split_on_page_breaks(paragraph, syntax_config=DocxMarkdownSyntaxConfig()) == [
        "**first**\n**second**",
        "**third**",
        "**fourth**",
    ]


def test_page_break_segments_preserve_formatting_inside_the_same_run():
    doc = Document()
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Before")
    run.bold = True
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")

    rendered = render_paragraph_runs_split_on_page_breaks(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == ["**Before**", "**After**"]


def test_page_break_segments_preserve_formula_callback_order():
    doc = Document()
    paragraph = doc.add_paragraph("Before")
    formula = OxmlElement("m:oMath")
    formula_text = OxmlElement("m:t")
    formula_text.text = "x"
    formula.append(formula_text)
    paragraph._p.append(formula)
    break_run = paragraph.add_run()
    break_run.add_break(WD_BREAK.PAGE)
    paragraph.add_run("After")

    rendered = render_paragraph_runs_split_on_page_breaks(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
        math_renderer=lambda _element: "$x$",
    )

    assert rendered == ["Before$x$", "After"]


@pytest.mark.parametrize("wrapper_tag", ("w:ins", "w:moveTo"))
def test_page_break_segments_include_accepted_revision_text(wrapper_tag: str):
    doc = Document()
    paragraph = doc.add_paragraph()
    wrapper = OxmlElement(wrapper_tag)
    run = OxmlElement("w:r")
    before = OxmlElement("w:t")
    before.text = "Before"
    run.append(before)
    page_break = OxmlElement("w:br")
    page_break.set(qn("w:type"), "page")
    run.append(page_break)
    after = OxmlElement("w:t")
    after.text = "After"
    run.append(after)
    wrapper.append(run)
    paragraph._p.append(wrapper)

    rendered = render_paragraph_runs_split_on_page_breaks(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == ["Before", "After"]


def test_page_break_segments_retain_empty_segment_between_adjacent_breaks():
    doc = Document()
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")

    rendered = render_paragraph_runs_split_on_page_breaks(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == ["Before", "", "After"]


def test_page_break_segments_preserve_note_references():
    class NoteExtractor:
        def get_reference_text(self, note_type, note_id):
            return f"[^{note_type}-{note_id}]"

    doc = Document()
    paragraph = doc.add_paragraph("Before")
    reference_run = paragraph.add_run()
    reference = OxmlElement("w:footnoteReference")
    reference.set(qn("w:id"), "7")
    reference_run._r.append(reference)
    reference_run.add_break(WD_BREAK.PAGE)
    paragraph.add_run("After")

    rendered = render_paragraph_runs_split_on_page_breaks(
        paragraph,
        note_extractor=NoteExtractor(),
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == ["Before[^footnote-7]", "After"]


def test_page_break_segments_wrap_each_hyperlink_side_independently():
    doc = Document()
    paragraph = doc.add_paragraph()
    rel_id = paragraph.part.relate_to(
        "https://example.com/page-aware",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    run_properties.append(OxmlElement("w:b"))
    run.append(run_properties)
    before = OxmlElement("w:t")
    before.text = "Before"
    run.append(before)
    page_break = OxmlElement("w:br")
    page_break.set(qn("w:type"), "page")
    run.append(page_break)
    after = OxmlElement("w:t")
    after.text = "After"
    run.append(after)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)

    rendered = render_paragraph_runs_split_on_page_breaks(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == [
        "[**Before**](https://example.com/page-aware)",
        "[**After**](https://example.com/page-aware)",
    ]


def test_regular_paragraph_renderer_keeps_page_break_as_line_break():
    doc = Document()
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("Before")
    run.add_break(WD_BREAK.PAGE)
    run.add_text("After")

    rendered = render_paragraph_runs(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == "Before\nAfter"


@pytest.mark.parametrize(
    ("style_name", "config_field", "expected"),
    [("HTML Code", "code", "`value`"), ("Quote Char", "quote", "value")],
)
def test_render_paragraph_runs_consumes_character_style_aliases(style_name, config_field, expected):
    doc = Document()
    try:
        doc.styles[style_name]
    except KeyError:
        doc.styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("value")
    run.style = style_name
    config = StyleDetectorConfig(
        code_character_style_names=frozenset({style_name}) if config_field == "code" else frozenset(),
        quote_character_style_names=frozenset({style_name}) if config_field == "quote" else frozenset(),
    )

    rendered = render_paragraph_runs(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
        style_detector_config=config,
    )

    assert rendered == expected


def test_render_paragraph_runs_recognizes_associated_character_style():
    doc = Document()
    style_name = "Runtime Code Alias Char"
    doc.styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("value")
    run.style = style_name

    rendered = render_paragraph_runs(
        paragraph,
        syntax_config=DocxMarkdownSyntaxConfig(),
        style_detector_config=StyleDetectorConfig(code_block_style_fragments=("Runtime Code Alias",)),
    )

    assert rendered == "`value`"


@pytest.mark.parametrize(
    ("tag", "value"),
    [("w:b", "0"), ("w:i", "false"), ("w:strike", "off"), ("w:u", "none")],
)
def test_explicitly_disabled_run_properties_do_not_emit_markdown(tag, value):
    doc = Document()
    run = doc.add_paragraph().add_run("value")
    property_element = OxmlElement(tag)
    property_element.set(qn("w:val"), value)
    run._r.get_or_add_rPr().append(property_element)

    rendered = render_paragraph_runs(
        doc.paragraphs[0],
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert rendered == "value"


@pytest.mark.parametrize(
    ("fill", "value", "config"),
    [
        ("D9D9D9", "clear", StyleDetectorConfig(wps_shading_enabled=False)),
        ("FFFFFF", "pct15", StyleDetectorConfig(word_shading_enabled=False)),
    ],
)
def test_run_shading_switches_can_disable_each_producer(fill, value, config):
    doc = Document()
    run = doc.add_paragraph().add_run("value")
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    shading.set(qn("w:val"), value)
    run._r.get_or_add_rPr().append(shading)

    rendered = render_paragraph_runs(
        doc.paragraphs[0],
        syntax_config=DocxMarkdownSyntaxConfig(),
        style_detector_config=config,
    )

    assert rendered == "value"


def _static_hyperlink_field(paragraph, instruction, *, split_instruction=False):
    """Build a minimal static field; no Word/WPS-saved fixture is required."""
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    paragraph.add_run()._r.append(begin)
    words = (instruction[:8], instruction[8:]) if split_instruction else (instruction,)
    for word in words:
        code = OxmlElement("w:instrText")
        code.set(qn("xml:space"), "preserve")
        code.text = word
        paragraph.add_run()._r.append(code)
    separator = OxmlElement("w:fldChar")
    separator.set(qn("w:fldCharType"), "separate")
    paragraph.add_run()._r.append(separator)
    result = paragraph.add_run("label")
    end_run = paragraph.add_run()
    end_run.style = "Hyperlink" if "Hyperlink" in paragraph.part.document.styles else "Default Paragraph Font"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)
    return begin, separator, result, end


@pytest.mark.parametrize("split_instruction", [False, True])
def test_complex_hyperlink_preserves_target_formatted_result_and_surrounding_text(split_instruction):
    document = Document()
    document.styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER)
    paragraph = document.add_paragraph("Before ")
    _, _, result, _ = _static_hyperlink_field(
        paragraph, r' HYPERLINK "https://example.org/qa" \h ', split_instruction=split_instruction
    )
    result.bold = True
    paragraph.add_run(" after")
    original = paragraph._p.xml

    assert render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig()) == (
        "Before [**label**](https://example.org/qa) after"
    )
    assert render_paragraph_runs(paragraph, preserve_formatting=False, syntax_config=DocxMarkdownSyntaxConfig()) == (
        "Before [label](https://example.org/qa) after"
    )
    assert paragraph._p.xml == original


@pytest.mark.parametrize(
    ("instruction", "destination"),
    [
        (r' hyperlink "https://example.org/qa" \H ', "https://example.org/qa"),
        (r" HYPERLINK https://example.org/qa ", "https://example.org/qa"),
        (r' HYPERLINK "https://example.org/qa" \* MERGEFORMAT ', "https://example.org/qa"),
        (r' HYPERLINK "https://example.org/qa" \o "A tip" \t "_blank" \n ', "https://example.org/qa"),
        (r' HYPERLINK "https://example.org/qa" \l "part" ', "https://example.org/qa#part"),
    ],
)
def test_complex_hyperlink_reads_static_destination_switches(instruction, destination):
    paragraph = Document().add_paragraph()
    _static_hyperlink_field(paragraph, instruction)
    assert render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig()) == f"[label]({destination})"


def test_complex_hyperlink_escapes_destination_without_changing_its_commonmark_target():
    from markdown_it import MarkdownIt

    paragraph = Document().add_paragraph()
    _static_hyperlink_field(paragraph, r' HYPERLINK "https://example.org/a (b)?v=1#part" \h ')
    rendered = render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig())
    children = MarkdownIt("commonmark").parseInline(rendered)[0].children
    assert children is not None
    assert [(token.type, token.attrGet("href")) for token in children if token.type == "link_open"] == [
        ("link_open", "https://example.org/a%20(b)?v=1#part")
    ]


@pytest.mark.parametrize(
    "instruction",
    [
        r' HYPERLINK "https://example.org/qa" \h \h ',
        r' HYPERLINK "https://example.org/qa" \p ',
        r' HYPERLINK "https://example.org/qa" injected ',
        r' HYPERLINK "https://example.org/qa" \* ROMAN ',
        r' HYPERLINK "https://example.org/qa" \o ',
        r' HYPERLINK "https://example.org/qa" \o \h ',
        r' "HYPERLINK" "https://example.org/qa" ',
        r' HYPERLINK "https://example.org/qa" "\h" ',
        r' HYPERLINK "https://example.org/qa ',
        r'HYPERLINK"https://example.org/qa"',
        r' REF "https://example.org/qa" \h ',
        r' INCLUDETEXT "https://example.org/qa" ',
    ],
)
def test_unsupported_complex_field_semantics_keep_visible_text_without_inventing_a_link(instruction):
    paragraph = Document().add_paragraph("Before ")
    _static_hyperlink_field(paragraph, instruction)
    paragraph.add_run(" after")
    assert render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig()) == "Before label after"


@pytest.mark.parametrize("damage", ["missing-separator", "missing-end", "mixed-end", "nested", "hidden-result"])
def test_complex_hyperlink_does_not_own_malformed_or_hidden_field_content(damage):
    paragraph = Document().add_paragraph("Before ")
    _, separate, result, end = _static_hyperlink_field(paragraph, r' HYPERLINK "https://example.org/qa" \h ')
    separate_run, end_run = separate.getparent(), end.getparent()
    assert separate_run is not None and end_run is not None
    if damage == "missing-separator":
        separate_run.remove(separate)
    elif damage == "missing-end":
        end_run.remove(end)
    elif damage == "mixed-end":
        text = OxmlElement("w:t")
        text.text = " boundary"
        end_run.append(text)
    elif damage == "nested":
        outer = OxmlElement("w:r")
        marker = OxmlElement("w:fldChar")
        marker.set(qn("w:fldCharType"), "begin")
        outer.append(marker)
        paragraph._p.insert(1, outer)
    else:
        result.font.hidden = True
    paragraph.add_run(" after")
    rendered = render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig())
    assert "https://" not in rendered and "Before " in rendered and rendered.endswith(" after")
    assert ("label" in rendered) is (damage != "hidden-result")
    assert (" boundary" in rendered) is (damage == "mixed-end")


def test_complex_hyperlink_never_consumes_authenticated_reference_nodes():
    paragraph = Document().add_paragraph()
    _, _, result, _ = _static_hyperlink_field(paragraph, r' HYPERLINK "https://example.org/qa" \h ')
    [text] = result._r.findall(qn("w:t"))
    assert (
        render_paragraph_runs(
            paragraph,
            syntax_config=DocxMarkdownSyntaxConfig(),
            reference_spans={text: (text, "@[[#^owned|Alias]]")},
        )
        == "@[[#^owned|Alias]]"
    )


def test_complex_hyperlink_keeps_page_breaks_and_literal_mode():
    paragraph = Document().add_paragraph()
    _, _, result, _ = _static_hyperlink_field(paragraph, r' HYPERLINK "https://example.org/qa" \h ')
    result.add_break(WD_BREAK.PAGE)
    result.add_text("after")
    assert render_paragraph_runs_split_on_page_breaks(paragraph, syntax_config=DocxMarkdownSyntaxConfig()) == [
        "[label](https://example.org/qa)",
        "[after](https://example.org/qa)",
    ]
    assert (
        render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig(), literal_text=True) == "label\nafter"
    )


@pytest.mark.parametrize("carrier", ["relationship", "complex-field"])
@pytest.mark.parametrize(
    ("target", "expected_href"),
    [
        (r"\\server\docx-image_1_rId9.png", "%5C%5Cserver%5Cdocx-image_1_rId9.png"),
        (r"\\server\share\file.png", "%5C%5Cserver%5Cshare%5Cfile.png"),
        (r"\\server\share\a (b).png", "%5C%5Cserver%5Cshare%5Ca%20(b).png"),
        (r"C:\folder\[item].png", "C:%5Cfolder%5C%5Bitem%5D.png"),
        (r"https://example.org/a\#part", "https://example.org/a%5C#part"),
        ("https://example.org/a?q=&copy;#part", "https://example.org/a?q=&copy;#part"),
        ("https://example.org/a?a=1&b=2#part", "https://example.org/a?a=1&b=2#part"),
        ("https://example.org/a?q=%2520&b=%20#part", "https://example.org/a?q=%2520&b=%20#part"),
    ],
)
def test_native_hyperlink_literal_target_keeps_parsed_uri_identity(carrier, target, expected_href):
    from pathlib import PureWindowsPath
    from urllib.parse import unquote, urlsplit

    import mistune
    from markdown_it import MarkdownIt

    paragraph = Document().add_paragraph()
    if carrier == "relationship":
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(
            qn("r:id"),
            paragraph.part.relate_to(
                target,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                is_external=True,
            ),
        )
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "label"
        run.append(text)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)
    else:
        operand = target.replace("\\", "\\\\")
        _static_hyperlink_field(paragraph, f' HYPERLINK "{operand}" \\h ')
    original = paragraph._p.xml
    rendered = render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig())
    commonmark = MarkdownIt("commonmark").parseInline(rendered)[0].children
    assert commonmark is not None
    assert [token.attrGet("href") for token in commonmark if token.type == "link_open"] == [expected_href]
    parsed = mistune.create_markdown(renderer="ast")(rendered)
    assert isinstance(parsed, list)
    [block] = parsed
    assert isinstance(block, dict)
    [mistune_link] = block["children"]
    assert mistune_link["attrs"]["url"] == expected_href
    assert unquote(expected_href) == unquote(target)
    if target.startswith("\\\\"):
        assert PureWindowsPath(unquote(expected_href)).drive == PureWindowsPath(target).drive
        assert PureWindowsPath(unquote(expected_href)).parts == PureWindowsPath(target).parts
    if target.startswith("https:"):
        assert urlsplit(expected_href).query == urlsplit(target).query
        assert urlsplit(expected_href).fragment == urlsplit(target).fragment
    assert paragraph._p.xml == original


@pytest.mark.parametrize(("control", "encoded"), [("\n", "%0A"), ("\r", "%0D"), ("\t", "%09")])
def test_native_hyperlink_relationship_control_characters_survive_package_and_markdown(control, encoded):
    from io import BytesIO
    from urllib.parse import unquote

    import mistune
    from markdown_it import MarkdownIt

    document = Document()
    paragraph = document.add_paragraph()
    target = f"https://example.org/a{control}b?q=1"
    hyperlink = OxmlElement("w:hyperlink")
    rel_id = paragraph.part.relate_to(
        target, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True
    )
    hyperlink.set(qn("r:id"), rel_id)
    run, text = OxmlElement("w:r"), OxmlElement("w:t")
    text.text = "label"
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    package = BytesIO()
    document.save(package)
    package.seek(0)
    reopened = Document(package)
    assert reopened.part.rels[rel_id].target_ref == target
    rendered = render_paragraph_runs(reopened.paragraphs[0], syntax_config=DocxMarkdownSyntaxConfig())
    expected_href = f"https://example.org/a{encoded}b?q=1"
    tokens = MarkdownIt("commonmark").parseInline(rendered)[0].children
    assert tokens is not None
    assert [token.attrGet("href") for token in tokens if token.type == "link_open"] == [expected_href]
    parsed = mistune.create_markdown(renderer="ast")(rendered)
    assert isinstance(parsed, list)
    [block] = parsed
    assert isinstance(block, dict)
    [link] = block["children"]
    assert link["attrs"]["url"] == expected_href
    assert unquote(expected_href) == target
