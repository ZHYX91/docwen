"""Current DOCX bytes retain inline formats, note identity and native breaks."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, cast
from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.style import CharacterStyle
from lxml import etree

from tests.integration._round_trip_helper import docx_to_md, md_to_docx

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@pytest.fixture
def formatting_runtime(tmp_path: Path, request: Any) -> Any:
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_runtime.config import ConfigLoader

    loader = ConfigLoader(user_dir=tmp_path / "configs")
    for key, value in {
        "template_fill.gongwen_numbering_text": False,
        "conversion.docx_to_md.preserve_formatting": True,
        "conversion.docx_to_md.preserve_heading_formatting": True,
        "conversion.docx_to_md.preserve_table_header_formatting": True,
        "conversion.md_to_docx.table_header_formatting_mode": "apply",
        "conversion.md_to_docx.heading_merge_mode": "never",
        "conversion.syntax.bold": getattr(request, "param", "asterisk"),
        "conversion.syntax.italic": getattr(request, "param", "asterisk"),
    }.items():
        assert loader.set_value(key, value)
    return create_runtime_port(config_loader=loader)


def _parts(path: Path) -> dict[str, Any]:
    with ZipFile(path) as package:
        return {
            name: etree.fromstring(package.read(name))
            for name in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")
            if name in package.namelist()
        }


def _text(element: Any) -> str:
    return "".join(text.text or "" for text in element.iter(_W + "t"))


def _assert_format(element: Any, text: str, **properties: str | bool) -> None:
    matching = [run for run in element.iter(_W + "r") if text in _text(run)]
    assert len(matching) == 1, (text, [_text(run) for run in matching])
    for name, expected in properties.items():
        prop = matching[0].find(f"{_W}rPr/{_W}{name}")
        if expected is True:
            assert prop is not None, (text, name)
            assert prop.get(_W + "val", "1") not in {"0", "false", "none"}
        elif expected is False:
            assert prop is None or prop.get(_W + "val") in {"0", "false", "none"}, (text, name)
        else:
            assert prop is not None and prop.get(_W + "val") == expected, (text, name)


def _assert_note_graph(parts: dict[str, Any]) -> None:
    document = parts["word/document.xml"]
    bookmarks: dict[str, tuple[str, str]] = {}
    active: dict[str, str] = {}
    references: list[tuple[str, str]] = []
    for node in document.iter():
        if node.tag == _W + "bookmarkStart":
            active[node.get(_W + "id")] = node.get(_W + "name")
        elif node.tag == _W + "bookmarkEnd":
            active.pop(node.get(_W + "id"), None)
        elif node.tag in {_W + "footnoteReference", _W + "endnoteReference"}:
            target = (node.tag.rsplit("}", 1)[-1], node.get(_W + "id"))
            references.append(target)
            for name in active.values():
                bookmarks[name] = target
        elif node.tag == _W + "instrText":
            match = re.search(r"\bNOTEREF\s+(\S+)", node.text or "")
            if match:
                assert match[1] in bookmarks
                references.append(bookmarks[match[1]])
    foot_refs = [value for kind, value in references if kind == "footnoteReference"]
    end_refs = [value for kind, value in references if kind == "endnoteReference"]
    assert len(foot_refs) == 3 and foot_refs[0] == foot_refs[1] and foot_refs[0] != foot_refs[2]
    assert len(end_refs) == 2 and end_refs[0] == end_refs[1]
    for part_name, tag, referenced in (
        ("word/footnotes.xml", "footnote", foot_refs),
        ("word/endnotes.xml", "endnote", end_refs),
    ):
        definitions = [note for note in parts[part_name].findall(_W + tag) if note.get(_W + "type") is None]
        assert {note.get(_W + "id") for note in definitions} == set(referenced)
        assert len(definitions) == len(set(referenced))


def _edit_package(source: Path, destination: Path, replacements: dict[str, str]) -> None:
    replaced: set[str] = set()
    with ZipFile(source) as package, ZipFile(destination, "w") as output:
        for info in package.infolist():
            data = package.read(info.filename)
            if info.filename in {"word/document.xml", "word/footnotes.xml", "word/endnotes.xml"}:
                root = etree.fromstring(data)
                for text in root.iter(_W + "t"):
                    if text.text in replacements:
                        replaced.add(text.text)
                        text.text = replacements[text.text]
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            output.writestr(info, data)
    assert replaced == set(replacements)


def test_formats_and_typed_notes_survive_two_cycles_and_native_edits(formatting_runtime: Any, tmp_path: Path):
    source = tmp_path / "source.md"
    source.write_text(
        "# 标题基线\n\n"
        "这是<u>正文下划线</u>。这是**<u>粗体下划线</u>**。这是*<u>斜体下划线</u>*。\n\n"
        "这是<u>外层下划线中有**内部粗体**和*内部斜体*</u>。\n\n"
        "H<sub>下标</sub>O、x<sup>上标</sup>、<del>删除</del>、<mark>高亮</mark>。\n\n"
        "| <u>表头下划线</u> | 普通表头 |\n| --- | --- |\n| <u>单元格下划线</u> | **表格粗体** |\n\n"
        "- <u>列表下划线</u>\n\n"
        "> <u>引用下划线</u>和**引用粗体**\n\n"
        "正文脚注[^f]，重复[^f]。尾注[^endnote:t]，重复[^endnote:t]。普通标签[^endnote-topic]。\n\n"
        "`<u>字面代码</u>`，转义：\\<u>字面转义\\</u>。\n\n"
        "```text\n<u>围栏字面</u>\n[^endnote:literal]\n___\n```\n\n"
        "[^f]: **<u>注释粗线</u>**和<u>**`注释代码`**</u>、*注释斜体*。\n\n"
        "    *注释第二段*\n\n"
        "[^endnote:t]: <mark>尾注高亮</mark>与H<sub>注释下标</sub>O。\n"
        "[^endnote-topic]: 普通脚注正文。\n",
        encoding="utf-8",
    )
    current = md_to_docx(formatting_runtime, source, tmp_path / "first", options={"remove_numbering": False})
    source.unlink()  # subsequent legs have only the current DOCX and its assets

    for generation in range(3):
        parts = _parts(current)
        body = parts["word/document.xml"]
        _assert_note_graph(parts)
        for text in (
            "正文下划线" if generation == 0 else "改后正文下划线",
            "表头下划线",
            "单元格下划线",
            "列表下划线",
            "引用下划线",
        ):
            _assert_format(body, text, u=True)
        _assert_format(body, "粗体下划线", u=True, b=True)
        _assert_format(body, "斜体下划线", u=True, i=True)
        _assert_format(body, "内部粗体", u=True, b=True)
        _assert_format(body, "内部斜体", u=True, i=True)
        _assert_format(body, "下标", vertAlign="subscript")
        _assert_format(body, "上标", vertAlign="superscript")
        _assert_format(body, "删除", strike=True)
        _assert_format(body, "高亮", highlight="yellow")
        _assert_format(body, "字面代码", u=False)
        _assert_format(body, "字面转义", u=False)
        _assert_format(body, "围栏字面", u=False)
        _assert_format(parts["word/footnotes.xml"], "注释粗线", u=True, b=True)
        _assert_format(parts["word/footnotes.xml"], "注释代码", u=True, b=True)
        _assert_format(parts["word/footnotes.xml"], "注释斜体", i=True)
        _assert_format(parts["word/footnotes.xml"], "注释第二段", i=True)
        rich_note = next(note for note in parts["word/footnotes.xml"] if "注释粗线" in _text(note))
        assert len(rich_note.findall(_W + "p")) == 2
        _assert_format(parts["word/endnotes.xml"], "尾注高亮", highlight="yellow")
        _assert_format(parts["word/endnotes.xml"], "注释下标", vertAlign="subscript")
        if generation == 2:
            break
        if generation == 0:
            isolated = tmp_path / "edited.docx"
            _edit_package(current, isolated, {"正文下划线": "改后正文下划线", "注释粗线": "改后注释粗线"})
            current = isolated
        markdown = docx_to_md(formatting_runtime, current, tmp_path / f"reverse-{generation}")
        assert "改后正文下划线" in markdown and "改后注释粗线" in markdown
        assert "[^endnote:" in markdown
        assert "<u>围栏字面</u>" in markdown and "[^endnote:literal]" in markdown
        assert "**标题基线**" not in markdown
        next_source = tmp_path / f"generation-{generation}.md"
        next_source.write_text(markdown, encoding="utf-8")
        current = md_to_docx(
            formatting_runtime, next_source, tmp_path / f"forward-{generation}", options={"remove_numbering": False}
        )
        next_source.unlink()


def _break_signature(markdown: str) -> list[str]:
    expected = {"分页前", "分页后", "分节后", "表格后", "---", "***", "___"}
    lines = markdown.splitlines()
    if lines and lines[0] == "---":
        lines = lines[lines.index("---", 1) + 1 :]
    return [
        "表格" if line.startswith("| 表格内容") else line
        for line in lines
        if line in expected or line.startswith("| 表格内容")
    ]


def test_style_based_rules_keep_native_positions_before_tables_and_at_end(formatting_runtime: Any, tmp_path: Path):
    source = tmp_path / "breaks.md"
    source.write_text(
        "分页前\n\n---\n\n分页后\n\n***\n\n分节后\n\n___\n\n| 表头 |\n| --- |\n| 表格内容 |\n\n表格后\n\n___\n",
        encoding="utf-8",
    )
    current = md_to_docx(formatting_runtime, source, tmp_path / "first-breaks")
    source.unlink()
    expected = ["分页前", "---", "分页后", "***", "分节后", "___", "表格", "表格后", "___"]
    for generation in range(2):
        root = _parts(current)["word/document.xml"]
        assert len([node for node in root.iter(_W + "br") if node.get(_W + "type") == "page"]) == 1
        assert len(root.findall(f".//{_W}pPr/{_W}sectPr")) == 1
        markdown = docx_to_md(formatting_runtime, current, tmp_path / f"break-back-{generation}")
        assert _break_signature(markdown) == expected
        next_source = tmp_path / f"break-{generation}.md"
        next_source.write_text(markdown, encoding="utf-8")
        current = md_to_docx(formatting_runtime, next_source, tmp_path / f"break-forward-{generation}")
        next_source.unlink()

    doc = Document(str(current))
    rule_paragraphs = [p for p in doc.paragraphs if p.style is not None and "HorizontalRule" in p.style.style_id]
    assert len(rule_paragraphs) == 2
    # A current direct reset overrides the style, so the removed rule must not
    # be reconstructed from original source metadata.
    border = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "nil")
    border.append(bottom)
    rule_paragraphs[0]._p.get_or_add_pPr().append(border)
    edited = tmp_path / "breaks-edited.docx"
    doc.save(str(edited))
    markdown = docx_to_md(formatting_runtime, edited, tmp_path / "break-edited-back")
    expected.remove("___")
    assert _break_signature(markdown) == expected


@pytest.mark.parametrize("formatting_runtime", ["asterisk", "underscore"], indirect=True)
def test_native_character_styles_and_resets_share_body_table_list_quote_note_rules(
    formatting_runtime: Any, tmp_path: Path
):
    source = tmp_path / "native-styles.md"
    source.write_text(
        "# 标题基线\n\n正文样式\n\n"
        "| 表头 |\n| --- |\n| 表格样式 |\n\n"
        "- 列表样式\n\n> 引用样式\n\n脚注[^f]。\n\n[^f]: 注释样式\n",
        encoding="utf-8",
    )
    initial = md_to_docx(formatting_runtime, source, tmp_path / "native-first")
    source.unlink()
    document = Document(str(initial))
    base = cast(CharacterStyle, document.styles.add_style("Authored Base", WD_STYLE_TYPE.CHARACTER))
    base.font.bold = True
    derived = cast(CharacterStyle, document.styles.add_style("Authored Details", WD_STYLE_TYPE.CHARACTER))
    derived.base_style = base
    derived.font.italic = True
    derived.font.underline = True
    labels = {"正文样式", "表格样式", "列表样式", "引用样式", "注释样式"}

    def style_runs(root: Any) -> None:
        for run in root.iter(_W + "r"):
            if _text(run) not in labels:
                continue
            properties = run.find(_W + "rPr")
            if properties is None:
                properties = etree.Element(_W + "rPr")
                run.insert(0, properties)
            for prop in list(properties):
                if prop.tag in {_W + "rStyle", _W + "b", _W + "i", _W + "u"}:
                    properties.remove(prop)
            style = etree.Element(_W + "rStyle", {_W + "val": derived.style_id})
            properties.insert(0, style)

    style_runs(document.element)
    run = document.add_paragraph().add_run("显式关闭")
    run.style = derived
    run.bold = run.italic = run.underline = False
    boundary = document.add_paragraph("前")
    boundary.add_run("Bold").bold = True
    boundary.add_run("Italic").italic = True
    boundary.add_run("后")
    document.add_paragraph().add_run(" spaced ").bold = True
    saved = tmp_path / "native-styled.docx"
    document.save(str(saved))
    current = tmp_path / "native-note-styled.docx"
    with ZipFile(saved) as package, ZipFile(current, "w") as output:
        for info in package.infolist():
            data = package.read(info.filename)
            if info.filename == "word/footnotes.xml":
                root = etree.fromstring(data)
                style_runs(root)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            output.writestr(info, data)

    for generation in range(2):
        markdown = docx_to_md(formatting_runtime, current, tmp_path / f"native-back-{generation}")
        assert "**标题基线**" not in markdown and "__标题基线__" not in markdown
        next_source = tmp_path / f"native-reimport-{generation}.md"
        next_source.write_text(markdown, encoding="utf-8")
        current = md_to_docx(formatting_runtime, next_source, tmp_path / f"native-forward-{generation}")
        next_source.unlink()
        parts = _parts(current)
        for label in labels:
            root = parts["word/footnotes.xml" if label == "注释样式" else "word/document.xml"]
            _assert_format(root, label, b=True, i=True, u=True)
        _assert_format(parts["word/document.xml"], "显式关闭", b=False, i=False, u=False)
        _assert_format(parts["word/document.xml"], "Bold", b=True, i=False)
        _assert_format(parts["word/document.xml"], "Italic", b=False, i=True)
        _assert_format(parts["word/document.xml"], "spaced", b=True)
