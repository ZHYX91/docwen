"""Source block ownership agrees with final native headings and content."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from docwen_bundle.runtime_factory import create_runtime_port
from docwen_runtime.config.loader import ConfigLoader
from tests.integration._round_trip_helper import docx_to_md
from tests.integration.test_markdown_visible_content import _NS, _convert, _text, _xml

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def _runtime(tmp_path: Path, *, remove_links: bool = False) -> Any:
    policy = "remove" if remove_links else "hyperlink"
    return create_runtime_port(
        config_loader=ConfigLoader(
            base_dir=Path(__file__).resolve().parents[2] / "configs",
            user_dir=tmp_path / "user-config",
            runtime_overrides={
                "link": {
                    "non_embed_links": {"markdown_mode": policy, "wiki_mode": policy, "auto_link_bare_url": False},
                    "embed_links": {"md_file_mode": "embed"},
                }
            },
        )
    )


def _write(
    runtime: Any, tmp_path: Path, source: str, *, numbering: bool = False, extended: bool | None = False
) -> Path:
    path = tmp_path / "source.md"
    path.write_bytes(source.encode("utf-8"))
    dialect = {"captions_references": True}
    if extended is not None:
        dialect["extended_headings"] = extended
    return _convert(
        runtime,
        path,
        tmp_path / "forward",
        options={
            "add_numbering": numbering,
            "remove_numbering": False,
            "numbering_scheme": "hierarchical_standard",
            "heading_numbering_render_mode": "text",
            "markdown_extensions": {"input": dialect},
        },
    )


def _paragraph_text(paragraph: ET.Element) -> str:
    pieces = []
    for node in paragraph.iter():
        if node.tag == f"{{{_NS['w']}}}t":
            pieces.append(node.text or "")
        elif node.tag == f"{{{_NS['w']}}}tab":
            pieces.append("\t")
        elif node.tag == f"{{{_NS['w']}}}br" and node.get(f"{{{_NS['w']}}}type", "textWrapping") == "textWrapping":
            pieces.append("\n")
    return "".join(pieces)


def _heading_texts(path: Path) -> list[tuple[int, str]]:
    """Read native outline roles, including paragraphs inside semantic SDTs."""
    with ZipFile(path) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        styles = ET.fromstring(package.read("word/styles.xml"))
    by_id = {style.get(f"{{{_NS['w']}}}styleId"): style for style in styles.findall("w:style", _NS)}
    result = []
    for paragraph in document.findall(".//w:p", _NS):
        properties = paragraph.find("w:pPr", _NS)
        visited = set()
        style = None
        while properties is not None:
            outline = properties.find("w:outlineLvl", _NS)
            if outline is not None:
                level = int(outline.get(f"{{{_NS['w']}}}val", "9"))
                if level < 9:
                    result.append((level + 1, _paragraph_text(paragraph).strip()))
                break
            style_id = properties.find("w:pStyle", _NS) if style is None else style.find("w:basedOn", _NS)
            if style_id is None:
                break
            key = style_id.get(f"{{{_NS['w']}}}val")
            if key in visited or key not in by_id:
                break
            visited.add(key)
            style = by_id[key]
            properties = style.find("w:pPr", _NS)
            if properties is None:
                break
    return result


def _field_results(root: ET.Element, command: str) -> list[str]:
    """Inspect the cached text between native separate/end field markers."""
    results = []
    stack: list[tuple[list[str], list[str], bool]] = []
    for node in root.iter():
        if node.tag == f"{{{_NS['w']}}}fldChar":
            kind = node.get(f"{{{_NS['w']}}}fldCharType")
            if kind == "begin":
                stack.append(([], [], False))
            elif kind == "separate" and stack:
                instructions, text, _separated = stack[-1]
                stack[-1] = (instructions, text, True)
            elif kind == "end" and stack:
                instructions, text, _separated = stack.pop()
                if "".join(instructions).strip().split()[0] == command:
                    results.append("".join(text))
        elif stack:
            instructions, text, separated = stack[-1]
            if node.tag == f"{{{_NS['w']}}}instrText":
                instructions.append(node.text or "")
            elif separated and node.tag == f"{{{_NS['w']}}}t":
                text.append(node.text or "")
    assert stack == []
    return results


def _read(runtime: Any, output: Path, tmp_path: Path) -> str:
    return docx_to_md(
        runtime,
        output,
        tmp_path / "reverse",
        options={"markdown_extensions": {"output": {"extended_headings": True, "captions_references": True}}},
    )


@pytest.mark.parametrize("opener", ["```", "<pre>"], ids=["fence", "html-pre"])
def test_lazy_list_fence_does_not_hide_later_active_links(tmp_path: Path, opener: str) -> None:
    runtime = _runtime(tmp_path, remove_links=True)
    source = f"- text\nlazy\n  {opener}\n  hidden\n\n# Real\nREAL ^target\n\n[out](https://example.test) [[Outside]]\n"
    output = _write(runtime, tmp_path, source)
    root = _xml(output)
    assert _heading_texts(output) == [(1, "Real")]
    assert [_paragraph_text(p).strip() for p in root.findall(".//w:p", _NS) if _paragraph_text(p).strip()] == [
        "text\nlazy",
        "hidden",
        "Real",
        "REAL",
    ]
    assert "out" not in _text(root) and "Outside" not in _text(root)
    assert root.find(".//w:hyperlink", _NS) is None
    returned = _read(runtime, output, tmp_path)
    assert "# Real" in returned and "REAL ^target" in returned
    assert "[out]" not in returned and "Outside" not in returned


def test_noninterrupting_ordered_marker_does_not_number_a_fake_heading(tmp_path: Path) -> None:
    runtime = _runtime(tmp_path)
    output = _write(runtime, tmp_path, "intro\n2. # Fake\n\n# Real\nBody\n", numbering=True)
    root = _xml(output)
    assert _heading_texts(output) == [(1, "1 Real")]
    assert "2. # Fake" in _text(root)
    assert "# 1 Fake" not in _text(root)
    returned = _read(runtime, output, tmp_path)
    assert "# 1 Real" in returned and "Fake" in returned and "# 1 Fake" not in returned


@pytest.mark.parametrize("marker", ["0.", "2.", "9)", "10."])
@pytest.mark.parametrize("literal", ["[literal](https://example.test)", "# Fake"])
def test_noninterrupting_ordered_markers_keep_the_complete_code_span(tmp_path: Path, marker: str, literal: str) -> None:
    runtime = _runtime(tmp_path, remove_links=True)
    output = _write(runtime, tmp_path, f"start `open\n{marker} {literal}\nclose`\n\n# Real\n", numbering=True)
    root = _xml(output)
    assert _heading_texts(output) == [(1, "1 Real")]
    assert [_paragraph_text(p).strip() for p in root.findall(".//w:p", _NS)] == [
        f"start open {marker} {literal} close",
        "1 Real",
    ]
    assert root.find(".//w:hyperlink", _NS) is None
    returned = _read(runtime, output, tmp_path)
    assert literal in returned and "# 1 Real" in returned


def test_ordered_start_one_can_interrupt_and_number_its_real_child_heading(tmp_path: Path) -> None:
    runtime = _runtime(tmp_path)
    output = _write(runtime, tmp_path, "intro\n1. # Actual\n\n# Real\nBody\n", numbering=True)
    assert _heading_texts(output) == [(1, "1 Actual"), (1, "2 Real")]
    returned = _read(runtime, output, tmp_path)
    assert "1 Actual" in returned and "2 Real" in returned and "intro" in returned


def test_setext_end_ends_the_inline_owner_before_a_later_link(tmp_path: Path) -> None:
    runtime = _runtime(tmp_path, remove_links=True)
    output = _write(runtime, tmp_path, "start `open\n-\n[literal](https://example.test) close`\n")
    root = _xml(output)
    assert _heading_texts(output) == [(2, "start `open")]
    assert "literal" not in _text(root)
    assert root.find(".//w:hyperlink", _NS) is None
    assert "close`" in _text(root)
    returned = _read(runtime, output, tmp_path)
    assert "literal" not in returned and "close`" in returned


def test_multiline_setext_keeps_its_valid_internal_code_owner(tmp_path: Path) -> None:
    runtime = _runtime(tmp_path, remove_links=True)
    output = _write(runtime, tmp_path, "First `open\n[literal](https://example.test) close`\n===\n\nAfter\n")
    root = _xml(output)
    assert _heading_texts(output) == [(1, "First open [literal](https://example.test) close")]
    assert root.find(".//w:hyperlink", _NS) is None
    returned = _read(runtime, output, tmp_path)
    assert "literal" in returned and "https://example.test" in returned and "After" in returned


def test_thematic_break_is_not_a_setext_numbering_candidate(tmp_path: Path) -> None:
    runtime = _runtime(tmp_path)
    output = _write(runtime, tmp_path, "***\n===\n\n# Real\nBody\n", numbering=True)
    root = _xml(output)
    assert _heading_texts(output) == [(1, "1 Real")]
    assert "===" in _text(root) and "1 ***" not in _text(root)
    returned = _read(runtime, output, tmp_path)
    assert "# 1 Real" in returned and "1 ***" not in returned


@pytest.mark.parametrize("prefix", ["Setext\n===\n", "> Setext\n> ===\n", "- Setext\n  ===\n"])
def test_setext_numbering_and_heading_reference_cache_share_the_inventory(tmp_path: Path, prefix: str) -> None:
    runtime = _runtime(tmp_path)
    output = _write(runtime, tmp_path, prefix + "\n# Real ^real\n\nSee @[[#^real]].\n", numbering=True)
    root = _xml(output)
    assert _heading_texts(output) == [(1, "1 Setext"), (1, "2 Real")]
    assert _field_results(root, "REF") == ["2"]
    assert "See 2." in _text(root)
    returned = _read(runtime, output, tmp_path)
    assert "Real ^real" in returned and "@[[#^real]]" in returned


def test_noninterrupting_list_spelling_participates_in_a_multiline_setext_heading(tmp_path: Path) -> None:
    runtime = _runtime(tmp_path)
    output = _write(runtime, tmp_path, "foo\n2. bar\n===\n\n# Real\n", numbering=True)
    assert _heading_texts(output) == [(1, "1 foo\n2. bar"), (1, "2 Real")]
    returned = _read(runtime, output, tmp_path)
    assert "foo<br>2. bar" in returned and "# 2 Real" in returned


def test_empty_atx_numbering_keeps_native_heading_structure(tmp_path: Path) -> None:
    runtime = _runtime(tmp_path)
    output = _write(runtime, tmp_path, "#\n\n# Real\nBody\n", numbering=True)
    root = _xml(output)
    assert _heading_texts(output) == [(1, "1"), (1, "2 Real")]
    assert "#1" not in _text(root)
    returned = _read(runtime, output, tmp_path)
    assert "# 1" in returned and "# 2 Real" in returned


@pytest.mark.parametrize("level", range(2, 10))
def test_empty_atx_levels_keep_their_native_roles_and_counter(tmp_path: Path, level: int) -> None:
    runtime = _runtime(tmp_path)
    ancestors = "".join(f"{'#' * parent} Parent {parent}\n\n" for parent in range(1, level))
    source = ancestors + f"{'#' * level}\n\n{'#' * level} Real ^real\n\nSee @[[#^real]].\n"
    output = _write(runtime, tmp_path, source, numbering=True, extended=level > 6)
    expected = [(parent, f"{'.'.join(['1'] * parent)} Parent {parent}") for parent in range(1, level)]
    prefix = ".".join(["1"] * (level - 1)) + "."
    assert _heading_texts(output) == [*expected, (level, prefix + "1"), (level, prefix + "2 Real")]
    assert _field_results(_xml(output), "REF") == [prefix + "2"]
    returned = _read(runtime, output, tmp_path)
    assert "Real ^real" in returned and "@[[#^real]]" in returned


@pytest.mark.parametrize("prefix", ["", "> ", "- "], ids=["top", "quote", "list"])
def test_empty_heading_closing_hashes_keep_bom_and_crlf_input_unchanged(tmp_path: Path, prefix: str) -> None:
    runtime = _runtime(tmp_path)
    output = _write(
        runtime, tmp_path, f"\ufeff{prefix}# ###\r\n\r\n# Real ^real\r\n\r\nSee @[[#^real]].\r\n", numbering=True
    )
    assert _heading_texts(output) == [(1, "1"), (1, "2 Real")]
    assert _field_results(_xml(output), "REF") == ["2"]
    returned = _read(runtime, output, tmp_path)
    assert "Real ^real" in returned and "@[[#^real]]" in returned


@pytest.mark.parametrize(
    ("title", "query", "native_heading", "expected", "excluded"),
    [
        ("A <https://example.test>", "A", "A", "SECOND", "FIRST"),
        ("A <https://example.test>", "A https://example.test", "A https://example.test", "FIRST", "SECOND"),
        ("A <user@example.test>", "A", "A", "SECOND", "FIRST"),
        ("A <user@example.test>", "A user@example.test", "A user@example.test", "FIRST", "SECOND"),
        (r"A \<https://example.test\>", "A <https://example.test>", "A <https://example.test>", "FIRST", "SECOND"),
        ("A `<literal>`", "A <literal>", "A <literal>", "FIRST", "SECOND"),
        ("A `<literal>`", "A", "A", "SECOND", "FIRST"),
        # Unsupported inline HTML remains visible in DOCX; its tag spelling is
        # nevertheless not part of the Markdown section's navigation title.
        ("A <span>visible</span>", "A visible", "A <span>visible</span>", "FIRST", "SECOND"),
        ("A <span></span>", "A", "A <span></span>", "FIRST", "SECOND"),
    ],
)
def test_embedded_heading_match_keeps_autolink_visible_text(
    tmp_path: Path, title: str, query: str, native_heading: str, expected: str, excluded: str
) -> None:
    runtime = _runtime(tmp_path)
    child = tmp_path / "child.md"
    raw = f"# {title}\nFIRST\n\n# A\nSECOND\n".encode()
    child.write_bytes(raw)
    output = _write(runtime, tmp_path, f"![[child.md#{query}]]\n")
    root = _xml(output)
    assert _heading_texts(output) == [(1, native_heading)]
    assert expected in _text(root) and excluded not in _text(root)
    with ZipFile(output) as package:
        relationship_root = ET.fromstring(package.read("word/_rels/document.xml.rels"))
    relationships = {
        item.get("Id"): (item.get("Target"), item.get("TargetMode"))
        for item in relationship_root.findall("rel:Relationship", _NS)
    }
    actual_links = [
        (_text(link), relationships.get(link.get(f"{{{_NS['r']}}}id"))) for link in root.findall(".//w:hyperlink", _NS)
    ]
    expected_links = {
        ("A <https://example.test>", "FIRST"): [("https://example.test", ("https://example.test", "External"))],
        ("A <user@example.test>", "FIRST"): [("user@example.test", ("mailto:user@example.test", "External"))],
    }.get((title, expected), [])
    assert actual_links == expected_links
    assert child.read_bytes() == raw
    returned = _read(runtime, output, tmp_path)
    assert expected in returned and excluded not in returned


@pytest.mark.parametrize("level", [1, 6])
@pytest.mark.parametrize("marker", ["- ", "2. "], ids=["unordered", "ordered"])
def test_standard_list_heading_with_source_id_keeps_its_native_list_and_recovery(
    tmp_path: Path, level: int, marker: str
) -> None:
    runtime = _runtime(tmp_path)
    source = f"{marker}{'#' * level} Standard ^id\n"
    output = _write(runtime, tmp_path, source)
    root = _xml(output)
    assert _heading_texts(output) == [(level, "Standard")]
    num_id = root.find(".//w:p/w:pPr/w:numPr/w:numId", _NS)
    assert num_id is not None
    value_attribute = f"{{{_NS['w']}}}val"
    with ZipFile(output) as package:
        numbering = ET.fromstring(package.read("word/numbering.xml"))
    native_num = numbering.find(f"w:num[@w:numId='{num_id.get(value_attribute)}']", _NS)
    assert native_num is not None
    abstract_id = native_num.find("w:abstractNumId", _NS)
    assert abstract_id is not None
    native_level = numbering.find(
        f"w:abstractNum[@w:abstractNumId='{abstract_id.get(value_attribute)}']/w:lvl[@w:ilvl='0']", _NS
    )
    assert native_level is not None
    num_format = native_level.find("w:numFmt", _NS)
    start = native_level.find("w:start", _NS)
    assert num_format is not None and start is not None
    assert num_format.get(value_attribute) == ("bullet" if marker == "- " else "decimal")
    assert start.get(value_attribute) == ("1" if marker == "- " else "2")
    assert len(root.findall(".//w:bookmarkStart", _NS)) == 1
    returned = _read(runtime, output, tmp_path)
    assert f"{'#' * level} Standard ^id" in returned
    assert returned.count("^id") == 1 and "DOCWENV3" not in returned


@pytest.mark.parametrize("container", ["> ", "- "], ids=["quote", "list"])
@pytest.mark.parametrize("level", [7, 8, 9])
@pytest.mark.parametrize("identified", [False, True], ids=["idless", "identified"])
def test_extended_container_heading_keeps_native_level_and_optional_source_id(
    tmp_path: Path, container: str, level: int, identified: bool
) -> None:
    runtime = _runtime(tmp_path)
    suffix = " ^id" if identified else ""
    output = _write(runtime, tmp_path, f"{container}{'#' * level} Seven{suffix}\n", extended=True)
    root = _xml(output)
    assert _heading_texts(output) == [(level, "Seven")]
    if container == "- ":
        assert root.find(".//w:p/w:pPr/w:numPr/w:numId", _NS) is not None
    else:
        assert root.find(".//w:p/w:pPr/w:pBdr/w:left", _NS) is not None
    assert "DOCWENV3" not in _text(root)
    assert len(root.findall(".//w:bookmarkStart", _NS)) == int(identified)
    returned = _read(runtime, output, tmp_path)
    assert f"{'#' * level} Seven" in returned
    assert returned.count("^id") == int(identified)


@pytest.mark.parametrize("extended", [None, False], ids=["default", "disabled"])
@pytest.mark.parametrize("container", ["> ", "- "], ids=["quote", "list"])
@pytest.mark.parametrize("level", [7, 8, 9])
def test_disabled_extended_container_headings_stay_visible_prose(
    tmp_path: Path, container: str, level: int, extended: bool | None
) -> None:
    runtime = _runtime(tmp_path)
    title = f"{'#' * level} Seven"
    output = _write(runtime, tmp_path, container + title + "\n", extended=extended)
    root = _xml(output)
    assert _heading_texts(output) == []
    assert _text(root) == title
    assert root.find(".//w:bookmarkStart", _NS) is None
    returned = _read(runtime, output, tmp_path)
    assert "Seven" in returned and "DOCWENV3" not in returned


@pytest.mark.parametrize("container", ["quote", "list"])
@pytest.mark.parametrize("formula", ["$$x+7=y$$", "$$\nx+7=y\n$$"], ids=["single-line", "multiline"])
def test_container_display_math_retains_native_formula_and_source_anchor(
    tmp_path: Path, container: str, formula: str
) -> None:
    runtime = _runtime(tmp_path)
    lines = [*formula.splitlines(), "", "^formula"]
    if container == "quote":
        nested = "\n".join("> " + line for line in lines)
    else:
        nested = "- " + lines[0] + "\n" + "\n".join("  " + line for line in lines[1:])
    source = "Before.\n\n" + nested + "\n\nAfter.\n\n# Real ^real\n\nSee @[[#^real]].\n"
    output = _write(runtime, tmp_path, source, numbering=True)
    root = _xml(output)

    equations = root.findall(".//m:oMath", _NS)
    assert len(equations) == 1
    assert "".join(node.text or "" for node in equations[0].findall(".//m:t", _NS)) == "x+7=y"
    assert _field_results(root, "SEQ") == []
    assert _field_results(root, "REF") == ["1"]
    assert _heading_texts(output) == [(1, "1 Real")]
    assert [_paragraph_text(p).strip() for p in root.findall(".//w:p", _NS) if _paragraph_text(p).strip()] == [
        "Before.",
        "After.",
        "1 Real",
        "See 1.",
    ]
    assert "DOCWENV3" not in _text(root) and "^formula" not in _text(root)

    digest = hashlib.sha256(b"docwen-target-map-v1\0heading\0real").hexdigest()
    bookmark_name = "DW_T_" + digest[:35]
    assert [node.get(f"{{{_NS['w']}}}name") for node in root.findall(".//w:bookmarkStart", _NS)] == [bookmark_name]
    instructions = [(node.text or "").split() for node in root.findall(".//w:instrText", _NS)]
    assert [instruction[:2] for instruction in instructions] == [["REF", bookmark_name]]

    target_namespace = "https://docwen.dev/schema/document-target-map/v1"
    with ZipFile(output) as package:
        maps = [
            ET.fromstring(package.read(name))
            for name in package.namelist()
            if name.startswith("customXml/item") and name.endswith(".xml")
        ]
    targets = [
        record
        for item in maps
        for record in item.findall(f"{{{target_namespace}}}targets/{{{target_namespace}}}target")
    ]
    assert [(item.get("kind"), item.get("source_id"), item.get("bookmark_name")) for item in targets] == [
        ("heading", "real", bookmark_name)
    ]
    anchors = [
        record
        for item in maps
        for record in item.findall(f"{{{target_namespace}}}anchors/{{{target_namespace}}}anchor")
    ]
    assert [(item.get("block_kind"), item.get("source_id")) for item in anchors] == [("equation", "formula")]
    value_attribute = f"{{{_NS['w']}}}val"
    anchor_wrappers = [
        wrapper
        for wrapper in root.findall(".//w:sdt", _NS)
        if (tag := wrapper.find("w:sdtPr/w:tag", _NS)) is not None and tag.get(value_attribute) == anchors[0].get("tag")
    ]
    assert len(anchor_wrappers) == 1
    anchor_content = anchor_wrappers[0].find("w:sdtContent", _NS)
    assert anchor_content is not None
    assert [block.tag for block in anchor_content] == [f"{{{_NS['w']}}}p"]
    assert anchor_content.find(".//m:oMath", _NS) is not None
    assert anchor_content.find(".//w:bookmarkStart", _NS) is None

    returned = _read(runtime, output, tmp_path)
    assert returned.count("x+7=y") == 1 and returned.count("^formula") == 1
    assert returned.count("# 1 Real ^real") == 1 and returned.count("@[[#^real]]") == 1
    assert (
        returned.index("Before.")
        < returned.index("x+7=y")
        < returned.index("^formula")
        < returned.index("After.")
        < returned.index("# 1 Real ^real")
        < returned.index("See @[[#^real]].")
    )


@pytest.mark.parametrize("signature", ["", "\ufeff"], ids=["without-signature", "with-signature"])
def test_yaml_body_bom_remains_literal_after_the_file_signature_is_consumed(tmp_path: Path, signature: str) -> None:
    runtime = _runtime(tmp_path)
    source = signature + "---\ncustom: metadata\n---\n\ufeff# Literal\n\n# Real\nBody.\n"
    expected_digest = hashlib.sha256(source.encode("utf-8")).hexdigest()
    output = _write(runtime, tmp_path, source)
    root = _xml(output)
    assert _heading_texts(output) == [(1, "Real")]
    assert [_paragraph_text(p) for p in root.findall(".//w:p", _NS) if _paragraph_text(p)] == [
        "\ufeff# Literal",
        "Real",
        "Body.",
    ]
    assert _text(root).count("\ufeff") == 1
    assert hashlib.sha256((tmp_path / "source.md").read_bytes()).hexdigest() == expected_digest
    returned = _read(runtime, output, tmp_path)
    assert returned.count("\ufeff") == 1 and returned.count("Literal") == 1
    assert "# Real" in returned and "Body." in returned


@pytest.mark.parametrize(
    ("leading_kind", "leading_source", "leading_texts"),
    [
        (
            "table",
            "- | Unique header |\n  | --- |\n  | Unique value |",
            ("Unique header", "Unique value"),
        ),
        ("fence", "- ```text\n  Code before continuation.\n  ```", ("Code before continuation.",)),
        ("math", "- $$x+7=y$$", ("x+7=y",)),
        ("heading", "- # Heading before continuation.", ("Heading before continuation.",)),
        ("paragraph", "- Paragraph before continuation.", ("Paragraph before continuation.",)),
    ],
    ids=["after-table", "after-fence", "after-math", "after-heading", "after-paragraph"],
)
def test_list_continuation_id_owns_only_its_paragraph(
    tmp_path: Path, leading_kind: str, leading_source: str, leading_texts: tuple[str, ...]
) -> None:
    runtime = _runtime(tmp_path)
    continuation = f"Continuation after {leading_kind}."
    outside = f"Outside after {leading_kind}."
    source = f"{leading_source}\n\n  {continuation} ^item-id\n\n{outside}\n"
    output = _write(runtime, tmp_path, source)
    root = _xml(output)

    native_prefix = [] if leading_kind == "math" else list(leading_texts)
    assert [_paragraph_text(p).strip() for p in root.findall(".//w:p", _NS) if _paragraph_text(p).strip()] == [
        *native_prefix,
        continuation,
        outside,
    ]
    assert _heading_texts(output) == ([(1, "Heading before continuation.")] if leading_kind == "heading" else [])
    tables = root.findall(".//w:tbl", _NS)
    assert len(tables) == int(leading_kind == "table")
    if tables:
        assert [[_text(cell) for cell in row.findall("w:tc", _NS)] for row in tables[0].findall("w:tr", _NS)] == [
            ["Unique header"],
            ["Unique value"],
        ]
    equations = root.findall(".//m:oMath", _NS)
    assert len(equations) == int(leading_kind == "math")
    if equations:
        assert "".join(node.text or "" for node in equations[0].findall(".//m:t", _NS)) == "x+7=y"

    value_attribute = f"{{{_NS['w']}}}val"
    tagged_wrappers = [
        (wrapper, tag.get(value_attribute, ""))
        for wrapper in root.findall(".//w:sdt", _NS)
        if (tag := wrapper.find("w:sdtPr/w:tag", _NS)) is not None
    ]
    fences = [wrapper for wrapper, tag in tagged_wrappers if tag.startswith("docwen-fenced-source-v1:")]
    assert len(fences) == int(leading_kind == "fence")
    if fences:
        assert _paragraph_text(fences[0]) == "Code before continuation.\n"
    assert root.find(".//w:bookmarkStart", _NS) is None
    assert _field_results(root, "SEQ") == [] and _field_results(root, "REF") == []
    assert "^item-id" not in _text(root) and "DOCWENV3" not in _text(root)

    digest = hashlib.sha256(b"docwen-anchor-map-v1\0anchor\0item-id").hexdigest()
    expected_tag = "docwen-anchor-v1:" + digest[:32]
    anchor_wrappers = [(wrapper, tag) for wrapper, tag in tagged_wrappers if tag.startswith("docwen-anchor-v1:")]
    assert len(anchor_wrappers) == 1 and anchor_wrappers[0][1] == expected_tag
    anchor_wrapper = anchor_wrappers[0][0]
    anchor_content = anchor_wrapper.find("w:sdtContent", _NS)
    assert anchor_content is not None
    assert [block.tag for block in anchor_content] == [f"{{{_NS['w']}}}p"]
    assert _paragraph_text(anchor_content[0]) == continuation
    assert anchor_content.find(".//w:sdt", _NS) is None
    assert anchor_content.find(".//m:oMath", _NS) is None
    if equations:
        native_order = list(root.iter())
        assert native_order.index(equations[0]) < native_order.index(anchor_wrapper)

    target_namespace = "https://docwen.dev/schema/document-target-map/v1"
    with ZipFile(output) as package:
        maps = [
            ET.fromstring(package.read(name))
            for name in package.namelist()
            if name.startswith("customXml/item") and name.endswith(".xml")
        ]
    assert [
        record
        for item in maps
        for record in item.findall(f"{{{target_namespace}}}targets/{{{target_namespace}}}target")
    ] == []
    assert [
        record.attrib
        for item in maps
        for record in item.findall(f"{{{target_namespace}}}anchors/{{{target_namespace}}}anchor")
    ] == [{"block_kind": "paragraph", "source_id": "item-id", "tag": expected_tag, "sha256": digest}]

    returned = _read(runtime, output, tmp_path)
    expected_texts = [*leading_texts, continuation, "^item-id", outside]
    assert all(returned.count(text) == 1 for text in expected_texts)
    assert [returned.index(text) for text in expected_texts] == sorted(returned.index(text) for text in expected_texts)
    assert f"{continuation} ^item-id" in returned


@pytest.mark.parametrize("lead", ["- Lead ^item\n", "-\n  Lead ^item\n"], ids=["normal-marker", "empty-marker"])
def test_list_item_and_continuation_ids_keep_their_distinct_native_owners(tmp_path: Path, lead: str) -> None:
    runtime = _runtime(tmp_path)
    source = lead + "\n  ```text\n  Literal body.\n  ```\n\n  Continuation. ^child\n\nOutside.\n"
    output = _write(runtime, tmp_path, source)
    root = _xml(output)
    expected_texts = ["Lead", "Literal body.", "Continuation.", "Outside."]
    paragraphs = root.findall(".//w:p", _NS)
    assert [_paragraph_text(p).strip() for p in paragraphs if _paragraph_text(p).strip()] == expected_texts
    assert _heading_texts(output) == []
    assert root.find(".//w:bookmarkStart", _NS) is None
    assert _field_results(root, "SEQ") == [] and _field_results(root, "REF") == []
    assert "^item" not in _text(root) and "^child" not in _text(root) and "DOCWENV3" not in _text(root)

    expected_records = []
    tags = {}
    for source_id, block_kind in [("item", "list_item"), ("child", "paragraph")]:
        digest = hashlib.sha256(f"docwen-anchor-map-v1\0anchor\0{source_id}".encode()).hexdigest()
        tags[source_id] = "docwen-anchor-v1:" + digest[:32]
        expected_records.append(
            {"block_kind": block_kind, "source_id": source_id, "tag": tags[source_id], "sha256": digest}
        )
    target_namespace = "https://docwen.dev/schema/document-target-map/v1"
    with ZipFile(output) as package:
        maps = [
            ET.fromstring(package.read(name))
            for name in package.namelist()
            if name.startswith("customXml/item") and name.endswith(".xml")
        ]
    records = [
        record.attrib
        for item in maps
        for record in item.findall(f"{{{target_namespace}}}anchors/{{{target_namespace}}}anchor")
    ]
    assert sorted(records, key=lambda item: item["source_id"]) == sorted(
        expected_records, key=lambda item: item["source_id"]
    )
    assert [
        record
        for item in maps
        for record in item.findall(f"{{{target_namespace}}}targets/{{{target_namespace}}}target")
    ] == []
    topology_namespace = "https://docwen.dev/schema/document-anchor-topology-map/v1"
    topology_maps = [item for item in maps if item.tag == f"{{{topology_namespace}}}documentAnchorTopologyMap"]
    assert len(topology_maps) == 1 and topology_maps[0].attrib == {"version": "1"}
    edge_digest = hashlib.sha256(
        f"docwen-anchor-topology-edge-v1\0{tags['child']}\0{tags['item']}".encode()
    ).hexdigest()
    assert [edge.attrib for edge in topology_maps[0].findall(f"{{{topology_namespace}}}edge")] == [
        {"child_tag": tags["child"], "parent_tag": tags["item"], "sha256": edge_digest}
    ]

    value_attribute = f"{{{_NS['w']}}}val"
    tagged_wrappers = [
        (wrapper, tag.get(value_attribute, ""))
        for wrapper in root.findall(".//w:sdt", _NS)
        if (tag := wrapper.find("w:sdtPr/w:tag", _NS)) is not None
    ]
    anchors = [(wrapper, tag) for wrapper, tag in tagged_wrappers if tag.startswith("docwen-anchor-v1:")]
    assert len(anchors) == 2 and {tag for _wrapper, tag in anchors} == set(tags.values())
    by_tag = {tag: wrapper for wrapper, tag in anchors}
    parent_content = by_tag[tags["item"]].find("w:sdtContent", _NS)
    child_content = by_tag[tags["child"]].find("w:sdtContent", _NS)
    assert parent_content is not None and child_content is not None
    assert [block.tag for block in parent_content] == [f"{{{_NS['w']}}}p", f"{{{_NS['w']}}}p", f"{{{_NS['w']}}}sdt"]
    assert _paragraph_text(parent_content[0]) == "Lead"
    assert parent_content[0].find("w:pPr/w:numPr/w:numId", _NS) is not None
    assert _paragraph_text(parent_content[1]) == "Literal body.\n"
    assert parent_content[2] is by_tag[tags["child"]]
    assert [block.tag for block in child_content] == [f"{{{_NS['w']}}}p"]
    assert _paragraph_text(child_content[0]) == "Continuation."
    assert child_content.find(".//w:sdt", _NS) is None
    assert [_paragraph_text(p) for p in parent_content.findall(".//w:p", _NS)] == [
        "Lead",
        "Literal body.\n",
        "Continuation.",
    ]
    fences = [wrapper for wrapper, tag in tagged_wrappers if tag.startswith("docwen-fenced-source-v1:")]
    assert len(fences) == 1 and _paragraph_text(fences[0]) == "Literal body.\n"
    assert fences[0] in list(parent_content[1].iter())
    outside = [p for p in paragraphs if _paragraph_text(p) == "Outside."]
    assert len(outside) == 1 and outside[0] not in list(parent_content.iter())

    returned = _read(runtime, output, tmp_path)
    assert all(returned.count(text) == 1 for text in expected_texts)
    assert [returned.index(text) for text in expected_texts] == sorted(returned.index(text) for text in expected_texts)
    assert returned.count("^item") == 1 and returned.count("^child") == 1
    assert returned.count("```text") == 1 and returned.count("```") == 2
    assert "Continuation. ^child" in returned
    assert returned.index("Lead") < returned.index("^item") < returned.index("Outside.")
