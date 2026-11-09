"""Source parsing must reach real run formatting, including nested HTML."""

import pytest
from docx import Document
from docx.oxml.ns import qn

from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer import MdToDocxRenderer
from docwen_plugin_markdown.to_docx.notes import process_md_body_with_notes

pytestmark = pytest.mark.contract


def _render(source):
    document = Document()
    MdToDocxRenderer(document).render(parse_markdown_text(source))
    return document


@pytest.mark.parametrize(
    ("source", "text", "properties"),
    [
        ("这是<u>下划线</u>文本。", "下划线", {"u": "single"}),
        ("这是**<u>粗体下划线</u>**文本。", "粗体下划线", {"u": "single", "b": None}),
        ("这是*<u>斜体下划线</u>*文本。", "斜体下划线", {"u": "single", "i": None}),
        ("<u>外层**内部粗体***内部斜体*</u>", "内部粗体", {"u": "single", "b": None}),
        ("H<sub>2</sub>O", "2", {"vertAlign": "subscript"}),
        ("x<sup>2</sup>", "2", {"vertAlign": "superscript"}),
        ("<del>删除</del>", "删除", {"strike": None}),
        ("<mark>高亮</mark>", "高亮", {"highlight": "yellow"}),
        ("<strong>粗体</strong>", "粗体", {"b": None}),
        ("<em>斜体</em>", "斜体", {"i": None}),
        ("<U>大写标签</U>", "大写标签", {"u": "single"}),
        ("<u>外<u>内</u>外</u>", "内", {"u": "single"}),
        ("<sup>外<sub>内</sub></sup>", "内", {"vertAlign": "subscript"}),
        ("<u>`</u>`尾</u>", "</u>", {"u": "single"}),
    ],
)
def test_supported_html_formats_render_native_properties(source, text, properties):
    doc = _render(source)
    runs = [run for paragraph in doc.paragraphs for run in paragraph.runs if run.text == text]
    assert len(runs) == 1
    for name, value in properties.items():
        prop = runs[0]._r.find(qn("w:rPr") + "/" + qn(f"w:{name}"))
        assert prop is not None
        assert prop.get(qn("w:val")) == value


@pytest.mark.parametrize(
    ("source", "visible"),
    [
        ("`<u>代码</u>`", "<u>代码</u>"),
        (r"\<u>转义\</u>", "<u>转义</u>"),
        ('Text <u class="example">属性</u>', 'Text <u class="example">属性</u>'),
        ("Text <span>其他标签</span>", "Text <span>其他标签</span>"),
        ("````\n<u>围栏</u>\n````", "<u>围栏</u>"),
    ],
)
def test_literal_examples_do_not_acquire_underline(source, visible):
    doc = _render(source)
    assert "".join(paragraph.text for paragraph in doc.paragraphs) == visible
    assert all(not run.underline for paragraph in doc.paragraphs for run in paragraph.runs)


def test_notes_keep_nested_formatting_on_code_and_script_override():
    nodes, context = process_md_body_with_notes(
        "Body[^endnote:x].\n\n[^endnote:x]: <u>**`code`**</u> <sup>outer<sub>inner</sub></sup> <span>literal</span>\n"
    )
    doc = Document()
    MdToDocxRenderer(doc, note_ctx=context).render(nodes)
    note = context.endnote_elements[0]
    runs = {"".join(node.text or "" for node in run.iter(qn("w:t"))): run for run in note.iter(qn("w:r"))}
    code = runs["code"].find(qn("w:rPr"))
    assert code is not None
    assert all(code.find(qn(f"w:{name}")) is not None for name in ("b", "u", "shd"))
    scripts = runs["inner"].findall(qn("w:rPr") + "/" + qn("w:vertAlign"))
    assert len(scripts) == 1
    assert scripts[0].get(qn("w:val")) == "subscript"
    assert "<span>" in runs and "</span>" in runs


@pytest.mark.parametrize(
    ("source", "child_type"),
    [
        (r"<u>$\text{</u>}$ and **bold**</u>", "inline_math"),
        ("<u>[text](https://example.test/</u>) and **bold**</u>", "link"),
        ("<u>`</u>` and **bold**</u>", "codespan"),
    ],
)
def test_format_boundaries_ignore_tags_owned_by_other_inline_syntax(source, child_type):
    paragraph = parse_markdown_text(source)[0]
    assert len(paragraph["children"]) == 1
    underline = paragraph["children"][0]
    assert underline["type"] == "underline"
    assert underline["children"][0]["type"] == child_type
    assert underline["children"][-1]["type"] == "strong"
