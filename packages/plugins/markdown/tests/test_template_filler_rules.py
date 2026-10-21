"""Original-template substitution and safe built-in Gongwen layout contracts."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.styles.style import CharacterStyle, ParagraphStyle

from docwen_plugin_markdown.gongwen_fields import GongwenFieldPlacementError
from docwen_plugin_markdown.template_filler import fill_template, fill_yaml_placeholders
from docwen_plugin_markdown.template_utils import scan_placeholders

pytestmark = pytest.mark.contract


def _fill(document, values: dict[str, Any], *, enabled: bool = True) -> None:
    fill_template(document, values, [], None, scan_placeholders(document), gongwen_fields_enabled=enabled)


def test_all_fields_are_replaced_once_from_original_spans_across_runs() -> None:
    document = Document()
    paragraph = document.add_paragraph()
    first = paragraph.add_run("prefix {{fir")
    first.bold = True
    paragraph.add_run("st}} / {{second}} / {{third}} / {{first}}")
    values = {"first": "{{second}}", "second": None, "third": "{{附件说明}}"}
    _fill(document, values)
    assert paragraph.text == "prefix {{second}} /  / {{附件说明}} / {{second}}"
    assert first.bold is True


def test_missing_non_gongwen_fields_remove_only_markers() -> None:
    document = Document()
    paragraph = document.add_paragraph("prefix {{missing}} suffix")
    _fill(document, {})
    assert paragraph.text == "prefix  suffix"


def test_generic_lists_keep_exact_separator_and_zero_false_values() -> None:
    document = Document()
    paragraph = document.add_paragraph("{{普通列表}}")
    fill_template(
        document,
        {"普通列表": ["甲", ["乙", None], "", "null", "None", 0, False]},
        [],
        None,
        scan_placeholders(document),
        list_separator=" / ",
    )
    assert paragraph.text == "甲 / 乙 / 0 / False"


def test_rendered_markdown_and_inserted_attachment_names_are_literal() -> None:
    document = Document()
    body = document.add_paragraph("{{body}}")
    document.add_paragraph("{{附件说明}}")
    document.add_paragraph("{{名字}}")
    placeholders = scan_placeholders(document)
    rendered = document.add_paragraph("literal {{附件说明}} / {{主送机关}}")
    fill_template(
        document,
        {"附件说明": ["{{名字}}"], "名字": "值"},
        [rendered],
        body,
        placeholders,
    )
    assert [p.text for p in document.paragraphs] == [
        "literal {{附件说明}} / {{主送机关}}",
        "附件：{{名字}}",
        "值",
    ]


@pytest.mark.parametrize("count", [1, 2, 12, 100])
def test_every_body_cell_and_nested_cell_anchor_expands_in_place(count: int) -> None:
    document = Document()
    document.add_paragraph("before")
    document.add_paragraph("{{ 附件说明 }}")
    document.add_paragraph("after")
    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "{{附件说明}}"
    nested = table.cell(0, 1).add_table(rows=1, cols=1)
    nested.cell(0, 0).text = "{{附件说明}}"
    source = {"附件说明": [f"{i}. 材料{i}" for i in range(1, count + 1)]}
    original = deepcopy(source)

    _fill(document, source)

    expected = (
        ["附件：材料1"]
        if count == 1
        else [
            f"{'附件：' if i == 1 else '　' * ((7 - len(str(i))) // 2) + ' ' * ((7 - len(str(i))) % 2)}{i}. 材料{i}"
            for i in range(1, count + 1)
        ]
    )
    assert source == original
    assert [p.text for p in document.paragraphs] == ["before", *expected, "after"]
    assert [p.text for p in table.cell(0, 0).paragraphs] == expected
    assert [p.text for p in nested.cell(0, 0).paragraphs] == expected
    for paragraph in [*document.paragraphs[1:-1], *table.cell(0, 0).paragraphs, *nested.cell(0, 0).paragraphs]:
        left_indent = paragraph.paragraph_format.left_indent
        first_line_indent = paragraph.paragraph_format.first_line_indent
        assert left_indent is not None and left_indent > 0
        assert first_line_indent is not None and first_line_indent < 0
        assert "\t" not in paragraph.text


def test_attachment_expansion_preserves_paragraph_and_opening_run_properties() -> None:
    document = Document()
    style = document.styles.add_style("Attachment", WD_STYLE_TYPE.PARAGRAPH)
    assert isinstance(style, ParagraphStyle)
    style.font.size = Pt(14)
    char_style = document.styles.add_style("AttachmentName", WD_STYLE_TYPE.CHARACTER)
    assert isinstance(char_style, CharacterStyle)
    char_style.font.italic = True
    paragraph = document.add_paragraph(style=style)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(5)
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.right_indent = Pt(8)
    paragraph.paragraph_format.left_indent = Pt(28)
    opening = paragraph.add_run("{{附")
    opening.style = char_style
    opening.bold = True
    opening.font.size = Pt(14)
    paragraph.add_run("件说明}}")

    _fill(document, {"附件说明": ["甲", "乙"]})

    for rendered in document.paragraphs:
        paragraph_style = rendered.style
        run_style = rendered.runs[0].style
        assert paragraph_style is not None and paragraph_style.style_id == style.style_id
        assert run_style is not None and run_style.style_id == char_style.style_id
        assert rendered.runs[0].bold is True
        assert rendered.runs[0].font.size == Pt(14)
        assert rendered.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
        assert rendered.paragraph_format.space_before == Pt(3)
        assert rendered.paragraph_format.space_after == Pt(5)
        assert rendered.paragraph_format.line_spacing == 1.5
        assert rendered.paragraph_format.keep_with_next is True
        assert rendered.paragraph_format.right_indent == Pt(8)
        assert rendered.paragraph_format.left_indent == Pt(84)
        assert rendered.paragraph_format.first_line_indent == Pt(-56)


def test_merged_cell_anchor_is_expanded_only_once() -> None:
    document = Document()
    table = document.add_table(rows=1, cols=2)
    merged = table.cell(0, 0).merge(table.cell(0, 1))
    merged.text = "{{附件说明}}"
    _fill(document, {"附件说明": ["甲", "乙"]})
    assert [p.text for p in merged.paragraphs] == ["附件：1. 甲", "\u3000\u3000\u30002. 乙"]


def test_empty_attachment_and_paragraph_cleanup_preserve_valid_cell_structure() -> None:
    document = Document()
    document.add_paragraph("{{附件说明}}")
    table = document.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.text = "{{附件说明}}"
    nested = cell.add_table(rows=1, cols=1)
    nested.cell(0, 0).text = "unrelated"
    cell.paragraphs[-1].text = "{{主送机关}}："

    _fill(document, {"附件说明": [None, {"nested": " "}]})

    assert document.paragraphs == []
    assert nested.cell(0, 0).text == "unrelated"
    assert cell._tc[-1].tag == qn("w:p")
    assert cell.paragraphs[-1].text == ""


def test_only_pure_empty_field_rows_and_paragraphs_are_deleted() -> None:
    document = Document()
    document.add_paragraph("主送机关：{{主送机关}}")
    document.add_paragraph("正文不得删除 {{主送机关}}")
    document.add_paragraph("{{发文字号}} {{标题}}")
    table = document.add_table(rows=4, cols=2)
    table.rows[0].cells[0].text = "抄送机关"
    table.rows[0].cells[1].text = "{{抄送机关}}"
    table.rows[1].cells[0].text = "版记"
    table.rows[1].cells[1].text = "{{印发机关}} / {{印发日期}}"
    table.rows[2].cells[0].text = "{{抄送机关}}"
    table.rows[2].cells[1].text = "{{负责人}}"
    table.rows[3].cells[0].text = "{{抄送机关}}"
    table.rows[3].cells[1].text = "这部分正文必须保留"

    _fill(document, {"标题": "标题值", "负责人": "负责人值"})

    assert [p.text for p in document.paragraphs] == ["正文不得删除 ", " 标题值"]
    assert [[cell.text for cell in row.cells] for row in table.rows] == [
        ["", "负责人值"],
        ["", "这部分正文必须保留"],
    ]


def test_nonempty_fields_anywhere_in_a_row_prevent_row_deletion() -> None:
    document = Document()
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "{{抄送机关}}"
    table.cell(0, 1).text = "{{印发机关}}"
    table.cell(1, 0).text = "{{印发机关}}"
    table.cell(1, 1).text = "{{印发日期}}"
    _fill(document, {"印发机关": "办公室", "抄送机关": []})
    assert len(table.rows) == 2
    assert [[cell.text for cell in row.cells] for row in table.rows] == [["", "办公室"], ["办公室", ""]]


def test_row_group_must_be_complete_in_one_row_and_nested_rows_are_independent() -> None:
    document = Document()
    outer = document.add_table(rows=3, cols=1)
    outer.cell(0, 0).text = "{{印发机关}}"
    outer.cell(1, 0).text = "{{印发日期}}"
    nested = outer.cell(2, 0).add_table(rows=1, cols=2)
    nested.cell(0, 0).text = "{{印发机关}}"
    nested.cell(0, 1).text = "{{印发日期}}"
    _fill(document, {})
    assert len(outer.rows) == 3
    assert outer.cell(2, 0).tables == []
    assert outer.cell(2, 0)._tc[-1].tag == qn("w:p")


def test_last_empty_row_removes_table_but_not_other_content() -> None:
    document = Document()
    document.add_paragraph("正文")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "抄送：{{抄送机关}}"
    _fill(document, {})
    assert document.tables == []
    assert [p.text for p in document.paragraphs] == ["正文"]


def test_empty_rule_preserves_drawing_in_same_paragraph() -> None:
    document = Document()
    paragraph = document.add_paragraph("{{主送机关}}")
    paragraph.add_run()._r.append(OxmlElement("w:drawing"))
    _fill(document, {})
    assert len(document.paragraphs) == 1
    assert paragraph._p.find(f"{qn('w:r')}/{qn('w:drawing')}") is not None


@pytest.mark.parametrize(
    "location", ["inline", "duplicate", "header", "footer", "textbox", "break", "key_break", "hyperlink"]
)
def test_unsupported_attachment_locations_fail_when_enabled(location: str) -> None:
    document = Document()
    if location == "header":
        paragraph = document.sections[0].header.paragraphs[0]
    elif location == "footer":
        paragraph = document.sections[0].footer.paragraphs[0]
    else:
        paragraph = document.add_paragraph()
    if location == "inline":
        paragraph.text = "do not lose this {{附件说明}} suffix"
    elif location == "duplicate":
        paragraph.text = "{{附件说明}}{{附件说明}}"
    elif location == "break":
        paragraph.add_run("{{附件").add_break()
        paragraph.add_run("说明}}")
    elif location == "key_break":
        paragraph.add_run("{{").add_break()
        paragraph.add_run("附件说明}}")
    elif location in {"textbox", "hyperlink"}:
        wrapper = OxmlElement("w:txbxContent" if location == "textbox" else "w:hyperlink")
        owner = OxmlElement("w:p") if location == "textbox" else wrapper
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "{{附件说明}}"
        run.append(text)
        owner.append(run)
        if location == "textbox":
            wrapper.append(owner)
            paragraph.add_run()._r.append(wrapper)
        else:
            paragraph._p.append(wrapper)
    else:
        paragraph.text = "{{附件说明}}"
    snapshot = document.element.xml
    with pytest.raises(GongwenFieldPlacementError, match="附件说明占位符必须独占"):
        _fill(document, {"附件说明": ["甲", "乙"]})
    assert document.element.xml == snapshot


def test_disabling_fields_keeps_generic_inline_attachment_fill_and_empty_containers() -> None:
    document = Document()
    inline = document.add_paragraph("prefix {{附件说明}} suffix")
    empty = document.add_paragraph("主送：{{主送机关}}")
    header = document.sections[0].header.paragraphs[0]
    header.text = "{{附件说明}}"
    _fill(document, {"附件说明": ["甲", "乙"]}, enabled=False)
    assert inline.text == "prefix 甲、乙 suffix"
    assert empty.text == "主送："
    assert header.text == "甲、乙"


def test_hyperlink_placeholder_and_surrounding_run_properties_are_preserved() -> None:
    document = Document()
    paragraph = document.add_paragraph("prefix ")
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), "rId-test")
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    properties.append(OxmlElement("w:b"))
    run.append(properties)
    text = OxmlElement("w:t")
    text.text = "{{first}} / {{second}}"
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    paragraph.add_run(" suffix")
    _fill(document, {"first": "{{second}}", "second": ""})
    assert paragraph._p.find(qn("w:hyperlink")) is hyperlink
    assert run.find(qn("w:rPr")) is properties
    assert paragraph.text == "prefix {{second}} /  suffix"


def test_any_yaml_key_keeps_each_placeholder_location_format() -> None:
    document = Document()
    first = document.add_paragraph()
    first.alignment = WD_ALIGN_PARAGRAPH.CENTER
    first_run = first.add_run("{{任意")
    first_run.font.size = Pt(16)
    first.add_run("键}}")
    second = document.add_paragraph()
    second.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    second_run = second.add_run("{{任意键}}")
    second_run.font.size = Pt(10.5)
    fill_yaml_placeholders(document, {"任意键": "真实内容"}, scan_placeholders(document))
    assert first.text == second.text == "真实内容"
    assert first.runs[0].font.size == Pt(16)
    assert second.runs[0].font.size == Pt(10.5)
    assert first.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert second.alignment == WD_ALIGN_PARAGRAPH.RIGHT
