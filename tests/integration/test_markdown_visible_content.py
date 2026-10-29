"""Final DOCX packages retain supported ordinary and rich Markdown content."""

from __future__ import annotations

import hashlib
import posixpath
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from PIL import Image

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from tests.integration._round_trip_helper import docx_to_md

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _convert(
    runtime: Any,
    source: Path,
    output: Path,
    *,
    options: dict[str, Any] | None = None,
    resources: list[tuple[str, Path, str]] | None = None,
) -> Path:
    authored = source.read_bytes()
    refs = [FileRef(path=str(source), format="markdown", category="document")]
    request_options = {"remove_numbering": False, **(options or {})}
    if resources is not None:
        refs[0] = FileRef(
            path=str(source),
            format="markdown",
            category="document",
            input_kind="document",
            input_role="source",
            logical_path="Notes/source.md",
            media_type="text/markdown",
        )
        images = []
        for token, path, logical_path in resources:
            if not any(ref.path == str(path) for ref in refs):
                refs.append(
                    FileRef(
                        path=str(path),
                        format="png",
                        category="image",
                        input_kind="resource",
                        input_role="linked_resource",
                        logical_path=logical_path,
                        media_type="image/png",
                    )
                )
            images.append({"authored_token": token, "logical_path": logical_path})
        request_options["markdown_resource_bindings"] = {
            "authored_sha256": hashlib.sha256(authored).hexdigest(),
            "images": images,
        }
    result = runtime.execute(
        ConversionRequest(
            request_id="visible-content",
            input_refs=refs,
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(output)),
            options=request_options,
        )
    )
    assert result.success, result.error
    assert source.read_bytes() == authored
    return Path(next(artifact.staging_path for artifact in result.artifacts if artifact.kind == "primary"))


def _xml(path: Path) -> ET.Element:
    with ZipFile(path) as package:
        return ET.fromstring(package.read("word/document.xml"))


def _text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:t", _NS))


def _table_matrix(table: ET.Element) -> list[list[str]]:
    return [[_text(cell) for cell in row.findall("w:tc", _NS)] for row in table.findall("w:tr", _NS)]


def _container(table: str, kind: str) -> str:
    if kind == "quote":
        return "\n".join("> " + line for line in table.splitlines())
    if kind == "list":
        return "- Parent.\n\n" + "\n".join("  " + line for line in table.splitlines())
    return table


def _extensions(enabled: bool | None) -> dict[str, Any]:
    return {} if enabled is None else {"markdown_extensions": {"input": {"structural_tables": enabled}}}


@pytest.mark.parametrize("enabled", [None, False, True], ids=["default", "disabled", "enabled"])
@pytest.mark.parametrize("kind", ["top", "quote", "list"])
def test_ordinary_pipe_tables_are_independent_of_structural_opt_in(
    round_trip_runtime: Any,
    tmp_path: Path,
    enabled: bool | None,
    kind: str,
) -> None:
    table = "| Name | Count |\n| --- | --- |\n| Alpha | 7 |"
    source = tmp_path / "ordinary.md"
    source.write_text("Before.\n\n" + _container(table, kind) + "\n\nAfter.\n", encoding="utf-8")
    output = _convert(round_trip_runtime, source, tmp_path / "forward", options=_extensions(enabled))
    root = _xml(output)
    tables = root.findall(".//w:tbl", _NS)
    assert len(tables) == 1
    assert _table_matrix(tables[0]) == [["Name", "Count"], ["Alpha", "7"]]
    expected = ["Before."] + (["Parent."] if kind == "list" else []) + ["Name", "Count", "Alpha", "7", "After."]
    text = _text(root)
    assert all(text.count(value) == 1 for value in expected)
    assert [text.index(value) for value in expected] == sorted(text.index(value) for value in expected)
    returned = docx_to_md(round_trip_runtime, output, tmp_path / "reverse", options={"markdown_extensions": {}})
    assert all(value in returned for value in ["Name", "Count", "Alpha", "7", "Before.", "After."])


@pytest.mark.parametrize("enabled", [None, False, True], ids=["default", "disabled", "enabled"])
@pytest.mark.parametrize("kind", ["quote", "list"])
def test_nested_ordinary_tables_do_not_enable_merge_markers(
    round_trip_runtime: Any,
    tmp_path: Path,
    enabled: bool | None,
    kind: str,
) -> None:
    source = tmp_path / "markers.md"
    source.write_text(_container("| Group | < |\n| --- | --- |\n| Data | Value |", kind), encoding="utf-8")
    output = _convert(round_trip_runtime, source, tmp_path / "forward", options=_extensions(enabled))
    tables = _xml(output).findall(".//w:tbl", _NS)
    assert len(tables) == 1
    if enabled:
        assert _table_matrix(tables[0]) == [["Group"], ["Data", "Value"]]
        assert tables[0].find("w:tr/w:tc/w:tcPr/w:gridSpan", _NS).get(f"{{{_NS['w']}}}val") == "2"
    else:
        assert _table_matrix(tables[0]) == [["Group", "<"], ["Data", "Value"]]
        assert tables[0].find(".//w:gridSpan", _NS) is None


@pytest.mark.parametrize("enabled", [False, True])
def test_quoted_fenced_table_stays_literal(round_trip_runtime: Any, tmp_path: Path, enabled: bool) -> None:
    source = tmp_path / "literal.md"
    literal = "| Name | Count |\n| --- | --- |\n| Alpha | 7 |"
    source.write_text(_container("```markdown\n" + literal + "\n```", "quote"), encoding="utf-8")
    output = _convert(round_trip_runtime, source, tmp_path / "forward", options=_extensions(enabled))
    root = _xml(output)
    assert root.find(".//w:tbl", _NS) is None
    assert all(line in _text(root) for line in literal.splitlines())


def _drawings(path: Path) -> list[tuple[str, str]]:
    """Match each drawing's own relationship bytes to its own alt description."""
    with ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        relationships = ET.fromstring(package.read("word/_rels/document.xml.rels"))
        targets = {item.get("Id"): item.get("Target", "") for item in relationships}
        result = []
        for drawing in root.findall(".//w:drawing", _NS):
            blip = drawing.find(".//a:blip", _NS)
            props = drawing.find(".//wp:docPr", _NS)
            assert blip is not None and props is not None
            target = targets[blip.get(f"{{{_NS['r']}}}embed")]
            media_path = posixpath.normpath(posixpath.join("word", target))
            result.append((hashlib.sha256(package.read(media_path)).hexdigest(), props.get("descr", "")))
        return result


def test_highlight_keeps_formula_image_link_note_and_nested_emphasis(
    round_trip_runtime: Any,
    tmp_path: Path,
) -> None:
    image = tmp_path / "blue.png"
    Image.new("RGB", (8, 6), "blue").save(image)
    source = tmp_path / "highlight.md"
    source.write_text(
        "Before ==**Strong** and *emphasis*, $x+1$, [Link](https://example.com), "
        "![Original blue alt](blue.png), note[^one]== After.\n\n[^one]: Note body.\n",
        encoding="utf-8",
    )
    output = _convert(round_trip_runtime, source, tmp_path / "forward")
    root = _xml(output)
    text = _text(root)
    assert text == "Before Strong and emphasis, , Link, , note After."
    assert "".join(node.text or "" for node in root.findall(".//m:t", _NS)) == "x+1"
    assert len(root.findall(".//m:oMath", _NS)) == 1
    assert _drawings(output) == [(hashlib.sha256(image.read_bytes()).hexdigest(), "Original blue alt")]
    assert root.find(".//w:footnoteReference", _NS) is not None
    for label, property_name in [("Strong", "b"), ("emphasis", "i")]:
        run = next(run for run in root.findall(".//w:r", _NS) if _text(run) == label)
        assert run.find(f"w:rPr/w:{property_name}", _NS) is not None
        assert run.find("w:rPr/w:highlight", _NS).get(f"{{{_NS['w']}}}val") == "yellow"
    link = root.find(".//w:hyperlink", _NS)
    assert link is not None and _text(link) == "Link"
    assert link.find("w:r/w:rPr/w:highlight", _NS).get(f"{{{_NS['w']}}}val") == "yellow"
    with ZipFile(output) as package:
        assert "Note body." in _text(ET.fromstring(package.read("word/footnotes.xml")))
        rels = ET.fromstring(package.read("word/_rels/document.xml.rels"))
        assert any(item.get("Target") == "https://example.com" for item in rels)
    returned = docx_to_md(round_trip_runtime, output, tmp_path / "reverse", options={"markdown_extensions": {}})
    assert all(value in returned for value in ["Strong", "emphasis", "x+1", "Link", "Note body."])


def test_task_items_keep_states_children_and_order(round_trip_runtime: Any, tmp_path: Path) -> None:
    image = tmp_path / "blue.png"
    Image.new("RGB", (8, 6), "blue").save(image)
    source = tmp_path / "tasks.md"
    source.write_text(
        "Before.\n\n- [ ] TODO_ALPHA\n\n  Continuation.\n\n  ![Task image](blue.png)\n\n"
        "  | Name | Count |\n  | --- | --- |\n  | Task data | 7 |\n\n"
        "  - [x] NESTED_DONE\n\n- [x] DONE_BETA\n- Ordinary item.\n- [?] Literal marker.\n\nAfter.\n",
        encoding="utf-8",
    )
    output = _convert(round_trip_runtime, source, tmp_path / "forward")
    root = _xml(output)
    paragraphs = [_text(item) for item in root.findall(".//w:body/w:p", _NS)]
    assert "[ ] TODO_ALPHA" in paragraphs
    assert "[x] NESTED_DONE" in paragraphs
    assert "[x] DONE_BETA" in paragraphs
    assert "[?] Literal marker." in paragraphs
    expected = [
        "Before.",
        "TODO_ALPHA",
        "Continuation.",
        "Name",
        "Count",
        "Task data",
        "7",
        "NESTED_DONE",
        "DONE_BETA",
        "Ordinary item.",
        "After.",
    ]
    text = _text(root)
    assert all(text.count(value) == 1 for value in expected)
    assert [text.index(value) for value in expected] == sorted(text.index(value) for value in expected)
    assert _table_matrix(root.find(".//w:tbl", _NS)) == [["Name", "Count"], ["Task data", "7"]]
    assert _drawings(output) == [(hashlib.sha256(image.read_bytes()).hexdigest(), "Task image")]
    for paragraph in root.findall(".//w:body/w:p", _NS):
        if _text(paragraph).startswith(("[ ]", "[x]")):
            assert paragraph.find("w:pPr/w:numPr", _NS) is not None
    returned = docx_to_md(round_trip_runtime, output, tmp_path / "reverse", options={"markdown_extensions": {}})
    assert all(
        value in returned
        for value in ["[ ] TODO_ALPHA", "[x] NESTED_DONE", "[x] DONE_BETA", "Continuation.", "Task data"]
    )


def test_removing_heading_formatting_keeps_highlight_semantic_children(
    round_trip_runtime: Any,
    tmp_path: Path,
) -> None:
    image = tmp_path / "blue.png"
    Image.new("RGB", (8, 6), "blue").save(image)
    source = tmp_path / "heading.md"
    source.write_text("# ==Strong $x+1$ ![Heading image](blue.png)==\n", encoding="utf-8")
    output = _convert(
        round_trip_runtime,
        source,
        tmp_path / "forward",
        options={"heading_formatting_mode": "remove"},
    )
    root = _xml(output)
    assert _text(root) == "Strong  "
    assert "".join(node.text or "" for node in root.findall(".//m:t", _NS)) == "x+1"
    assert _drawings(output) == [(hashlib.sha256(image.read_bytes()).hexdigest(), "Heading image")]
    assert root.find(".//w:highlight", _NS) is None


@pytest.mark.parametrize("marker", ["^^", "^", "~"], ids=["underline", "superscript", "subscript"])
def test_nested_highlight_style_containers_keep_formula_and_image(
    round_trip_runtime: Any,
    tmp_path: Path,
    marker: str,
) -> None:
    image = tmp_path / "blue.png"
    Image.new("RGB", (8, 6), "blue").save(image)
    source = tmp_path / "nested-highlight.md"
    source.write_text(
        f"Before =={marker}$x${marker} {marker}![Blue](blue.png){marker}== After.\n",
        encoding="utf-8",
    )
    output = _convert(round_trip_runtime, source, tmp_path / "forward")
    root = _xml(output)
    assert _text(root) == "Before   After."
    assert "".join(node.text or "" for node in root.findall(".//m:t", _NS)) == "x"
    assert _drawings(output) == [(hashlib.sha256(image.read_bytes()).hexdigest(), "Blue")]


@pytest.mark.parametrize("wiki_first", [True, False])
def test_image_alts_belong_to_each_source_occurrence(round_trip_runtime: Any, tmp_path: Path, wiki_first: bool) -> None:
    red = tmp_path / "private-0000.png"
    blue = tmp_path / "private-0001.png"
    Image.new("RGB", (8, 6), "red").save(red)
    Image.new("RGB", (8, 6), "blue").save(blue)
    wiki = ("![[red.png]]", red, "red.png", "red.png")
    markdown = ("![Original blue alt](blue.png)", blue, "blue.png", "Original blue alt")
    occurrences = ([wiki, markdown] if wiki_first else [markdown, wiki]) + [
        (r"![**Blue** \[bracket\] \| label](blue.png)", blue, "blue.png", "Blue [bracket] | label"),
        ("![[red.png|Wiki red label]]", red, "red.png", "Wiki red label"),
        ("![](blue.png)", blue, "blue.png", ""),
        ("![  Blue alt  ](blue.png)", blue, "blue.png", "  Blue alt  "),
    ]
    source = tmp_path / "images.md"
    source.write_text(
        "\n\n".join(item[0] for item in occurrences) + "\n\n`![not an image](missing.png)`\n", encoding="utf-8"
    )
    output = _convert(
        round_trip_runtime,
        source,
        tmp_path / "forward",
        resources=[(token, path, logical) for token, path, logical, _alt in occurrences],
    )
    assert _drawings(output) == [
        (hashlib.sha256(path.read_bytes()).hexdigest(), alt) for _token, path, _logical, alt in occurrences
    ]
    assert "![not an image](missing.png)" in _text(_xml(output))
    returned = docx_to_md(
        round_trip_runtime,
        output,
        tmp_path / "reverse",
        options={"markdown_extensions": {}, "image_link_style": "markdown_embed"},
    )
    assert "Original blue alt" in returned and "Wiki red label" in returned


def test_table_image_occurrences_keep_alt_dimensions_and_column_boundaries(
    round_trip_runtime: Any,
    tmp_path: Path,
) -> None:
    image = tmp_path / "private.png"
    Image.new("RGB", (8, 6), "blue").save(image)
    wiki_alt = "Alias [link](https://invalid.example) {literal}"
    wiki = f"![[blue.png|{wiki_alt}|20x10]]"
    markdown = r"![**Blue** \[bracket\] \| label](blue.png =30x15)"
    source = tmp_path / "image-table.md"
    source.write_text(
        f"| Image | Note |\n| --- | --- |\n| {wiki} | Wiki sentinel |\n| {markdown} | Markdown sentinel |\n",
        encoding="utf-8",
    )
    output = _convert(
        round_trip_runtime,
        source,
        tmp_path / "forward",
        resources=[(wiki, image, "blue.png"), (markdown, image, "blue.png")],
    )
    root = _xml(output)
    table = root.find(".//w:tbl", _NS)
    assert table is not None
    assert _table_matrix(table) == [["Image", "Note"], ["", "Wiki sentinel"], ["", "Markdown sentinel"]]
    digest = hashlib.sha256(image.read_bytes()).hexdigest()
    assert _drawings(output) == [(digest, wiki_alt), (digest, "Blue [bracket] | label")]
    extents = root.findall(".//wp:extent", _NS)
    assert [(int(item.get("cx", "0")) // 9525, int(item.get("cy", "0")) // 9525) for item in extents] == [
        (20, 10),
        (30, 15),
    ]
    assert root.find(".//w:hyperlink", _NS) is None


@pytest.mark.parametrize(
    ("configured", "requested", "structural"),
    [(None, None, False), (None, False, False), (None, True, True), (True, None, True), (True, False, False)],
    ids=["default", "request-disabled", "request-enabled", "config-enabled", "request-overrides-config"],
)
def test_link_removal_uses_the_same_table_dialect_as_the_writer(
    tmp_path: Path, configured: bool | None, requested: bool | None, structural: bool
) -> None:
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_runtime.config.loader import ConfigLoader

    overrides: dict[str, Any] = {"link": {"non_embed_links": {"markdown_mode": "remove"}}}
    if configured is not None:
        overrides["conversion"] = {"markdown_extensions": {"input": {"structural_tables": configured}}}
    runtime = create_runtime_port(
        config_loader=ConfigLoader(
            base_dir=Path(__file__).resolve().parents[2] / "configs",
            user_dir=tmp_path / "user-config",
            runtime_overrides=overrides,
        )
    )
    source = tmp_path / "table-scope.md"
    source.write_text(
        "| A | B |\n| --- | --- |\n| ` | [Visible](https://example.com) ` |\n",
        encoding="utf-8",
    )
    output = _convert(runtime, source, tmp_path / "forward", options=_extensions(requested))
    root = _xml(output)
    text = _text(root)
    tables = root.findall(".//w:tbl", _NS)
    if structural:
        # The advanced grammar treats the pipe inside the matched code span
        # as literal: its inconsistent width remains a visible paragraph.
        assert tables == []
        assert "| [Visible](https://example.com)" in text
    else:
        # Ordinary GFM cells have independent inline scopes. Neither tick in
        # this row is paired in its own cell, so the active link is removed.
        assert len(tables) == 1
        assert _table_matrix(tables[0]) == [["A", "B"], ["`", "`"]]
        assert "Visible" not in text and "https://example.com" not in text
    assert root.find(".//w:hyperlink", _NS) is None
    returned = docx_to_md(runtime, output, tmp_path / "reverse", options={"markdown_extensions": {}})
    assert ("[Visible](https://example.com)" in returned) is structural


@pytest.mark.parametrize(
    ("configured", "requested", "extended"),
    [(None, None, False), (None, False, False), (None, True, True), (True, None, True), (True, False, False)],
    ids=["default", "request-disabled", "request-enabled", "config-enabled", "request-overrides-config"],
)
def test_link_removal_uses_the_same_heading_dialect_as_the_writer(
    tmp_path: Path, configured: bool | None, requested: bool | None, extended: bool
) -> None:
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_runtime.config.loader import ConfigLoader

    overrides: dict[str, Any] = {"link": {"non_embed_links": {"markdown_mode": "remove"}}}
    if configured is not None:
        overrides["conversion"] = {"markdown_extensions": {"input": {"extended_headings": configured}}}
    runtime = create_runtime_port(
        config_loader=ConfigLoader(
            base_dir=Path(__file__).resolve().parents[2] / "configs",
            user_dir=tmp_path / "user-config",
            runtime_overrides=overrides,
        )
    )
    source = tmp_path / "heading-scope.md"
    source.write_text("`start\n####### １２３ [link](https://example.com)\nend`\n", encoding="utf-8")
    options = {} if requested is None else {"markdown_extensions": {"input": {"extended_headings": requested}}}
    output = _convert(runtime, source, tmp_path / "forward", options=options)
    root = _xml(output)
    text = _text(root)
    paragraphs = [_text(item) for item in root.findall(".//w:body/w:p", _NS)]
    if extended:
        # The opted-in H7 interrupts the paragraph. Its active link is
        # removed, and the unmatched ticks stay literal in their own blocks.
        assert paragraphs == ["`start", "１２３", "end`"]
        assert "[link]" not in text and "https://example.com" not in text
    else:
        # H7 is ordinary paragraph text in the default grammar: one code
        # span owns all three lines, including the link-shaped literal.
        assert paragraphs == ["start ####### １２３ [link](https://example.com) end"]
    assert root.find(".//w:hyperlink", _NS) is None
    returned = docx_to_md(runtime, output, tmp_path / "reverse", options={"markdown_extensions": {}})
    assert ("[link](https://example.com)" in returned) is not extended


@pytest.mark.parametrize("enabled", [False, True], ids=["caption-disabled", "caption-enabled"])
@pytest.mark.parametrize(("ticks", "caption"), [("`", "Figure"), ("``", "Table")])
def test_caption_shaped_lines_inside_code_keep_the_complete_literal(
    round_trip_runtime: Any, tmp_path: Path, enabled: bool, ticks: str, caption: str
) -> None:
    source = tmp_path / "caption-code.md"
    source.write_text(f"{ticks}start\n{caption}: Literal title\nend{ticks}\n", encoding="utf-8")
    output = _convert(
        round_trip_runtime,
        source,
        tmp_path / "forward",
        options={"markdown_extensions": {"input": {"captions_references": enabled}}},
    )
    root = _xml(output)
    expected = f"start {caption}: Literal title end"
    assert [_text(item) for item in root.findall(".//w:body/w:p", _NS)] == [expected]
    assert "DOCWENV3" not in _text(root)
    assert root.find(".//w:fldSimple", _NS) is None
    assert root.find(".//w:fldChar", _NS) is None
    assert root.find(".//w:bookmarkStart", _NS) is None
    returned = docx_to_md(round_trip_runtime, output, tmp_path / "reverse", options={"markdown_extensions": {}})
    # The existing reader normalizes a whole code-only paragraph to a fenced
    # block. Its complete code text must survive that presentation choice.
    assert f"\n```\n{expected}\n```\n" in returned
    assert "DOCWENV3" not in returned
