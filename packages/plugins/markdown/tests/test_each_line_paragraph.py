"""Request-owned ordinary Markdown prose paragraph boundaries."""

from docx import Document

from docwen_plugin_markdown.ast_transforms import split_plain_paragraph_lines
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer import MdToDocxRenderer


def test_off_by_default_keeps_one_paragraph_with_soft_breaks() -> None:
    nodes = parse_markdown_text("First\nSecond\nThird")
    doc = Document()
    MdToDocxRenderer(doc).render(nodes)
    assert [paragraph.text for paragraph in doc.paragraphs] == ["First\nSecond\nThird"]


def test_each_line_creates_distinct_word_paragraphs_without_losing_styles() -> None:
    nodes = parse_markdown_text("First **bold**\nSecond `code`\nThird")
    doc = Document()
    MdToDocxRenderer(doc).render(split_plain_paragraph_lines(nodes))
    assert [paragraph.text for paragraph in doc.paragraphs] == ["First bold", "Second code", "Third"]
    assert len(doc.paragraphs) == 3


def test_explicit_break_and_containers_remain_intact() -> None:
    text = (
        "First  \\nSecond\\n\\n> Quote one\\n> Quote two\\n\\n"
        "- Item one\\n- Item two\\n\\n```\\nA\\nB\\n```"
    )
    original = parse_markdown_text(text)
    transformed = split_plain_paragraph_lines(original)
    assert transformed == original


def test_semantic_owner_does_not_split() -> None:
    source = {
        "type": "paragraph",
        "children": [
            {"type": "text", "raw": "First"},
            {"type": "softbreak"},
            {"type": "text", "raw": "Second"},
        ],
        "_docwen_v3_ordinary_anchor": {"id": "p"},
    }
    assert split_plain_paragraph_lines([source]) == [source]


def test_only_top_level_prose_is_split() -> None:
    nodes = parse_markdown_text("First\nSecond\n\n- Item\n  continuation\n\n# Heading")
    transformed = split_plain_paragraph_lines(nodes)
    assert [node["type"] for node in transformed][:2] == ["paragraph", "paragraph"]
    assert any(node["type"] == "list" for node in transformed)
