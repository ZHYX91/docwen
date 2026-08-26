"""Production composition gates for exact v4 fenced-source round trips."""

from __future__ import annotations

import re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import lxml.etree as etree
import pytest
from docx.oxml.ns import qn

from docwen_core.docx_semantics_v3 import (
    ANCHOR_TAG_PREFIX,
    ANCHOR_TOPOLOGY_MAP_NAMESPACE,
    FENCED_SOURCE_MAP_NAMESPACE,
    derive_anchor_identity_v3,
    derive_anchor_topology_edge_v3,
    derive_target_identity_v3,
)
from docwen_plugin_markdown.document_semantics_v3 import analyze_markdown_semantics_v3
from tests.integration._round_trip_helper import _primary_path, _run, docx_to_md

pytestmark = pytest.mark.integration


def _yaml(title: str) -> str:
    return f"---\naliases:\n  - {title}\ntitle: {title}\nsubtitle: \n---\n\n"


@pytest.mark.parametrize(
    ("case_id", "body"),
    [
        ("rust-mixed-eol", "```rust  exact\r\nalpha\r\nbeta\n```  \n"),
        ("mermaid-quote", "> ```mermaid\r\n> graph TD\r\n> ```\r\n"),
        ("query-list", "- ```query\n  tag:#x\n  ```\n"),
        ("view-omitted-eof", "```view\nbody"),
    ],
)
def test_fenced_source_exact_round_trip_matrix(
    tmp_path: Path,
    round_trip_runtime,
    case_id: str,
    body: str,
) -> None:
    source = _yaml(case_id) + body
    markdown_path = tmp_path / f"{case_id}.md"
    markdown_path.write_bytes(source.encode("utf-8"))

    forward = _run(
        round_trip_runtime,
        f"fenced-{case_id}-forward",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / f"{case_id}-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
        package_text = b"\n".join(package.read(name) for name in package.namelist()).decode(errors="ignore")

    assert FENCED_SOURCE_MAP_NAMESPACE in package_text
    assert document_xml.count("docwen-fenced-source-v1:") == 1
    assert "bookmarkStart" not in document_xml
    assert " SEQ " not in document_xml
    assert " REF " not in document_xml

    reverse = _run(
        round_trip_runtime,
        f"fenced-{case_id}-reverse",
        docx_path,
        source_format="docx",
        target_format="md",
        output_dir=tmp_path / f"{case_id}-reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False, "locale": "en"},
    )
    returned = _primary_path(reverse).read_bytes().decode("utf-8")
    assert returned == source
    analysis = analyze_markdown_semantics_v3(returned, input_id=f"{case_id}-returned.md")
    assert len(analysis.projection["fenced_sources"]) == 1


def test_caption_target_raw_anchor_and_fenced_carrier_remain_layered(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source = _yaml("caption-anchor") + "Code: Example ^code-target\n\n```rust\nfn main() {}\n```\n\n^raw-code\n"
    path = tmp_path / "caption-anchor.md"
    path.write_text(source, encoding="utf-8", newline="")
    forward = _run(
        round_trip_runtime,
        "fenced-caption-anchor-forward",
        path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "caption-anchor-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
        package_payload = b"\n".join(package.read(name) for name in package.namelist())
        document_root = etree.fromstring(package.read("word/document.xml"))

    assert document_xml.count(" SEQ Code ") == 1
    assert document_xml.count("bookmarkStart") == 1
    assert document_xml.count("docwen-fenced-source-v1:") == 1
    assert ANCHOR_TOPOLOGY_MAP_NAMESPACE.encode() not in package_payload
    target = derive_target_identity_v3("code_block", "code-target")
    anchor = derive_anchor_identity_v3("code_block", "raw-code")
    tags = _semantic_sdt_tags(document_root, include_targets=True)
    assert tags[:2] == [target.tag, anchor.tag]
    assert tags[2].startswith("docwen-fenced-source-v1:")
    returned = docx_to_md(round_trip_runtime, docx_path, tmp_path / "caption-anchor-reverse")
    analysis = analyze_markdown_semantics_v3(returned, input_id="returned.md")
    assert [(item["id"], item["kind"]) for item in analysis.projection["targets"]] == [("code-target", "code_block")]
    assert [(item["id"], item["block_kind"]) for item in analysis.projection["anchors"]] == [("raw-code", "code_block")]
    assert len(analysis.projection["fenced_sources"]) == 1


def test_multi_paragraph_list_item_and_whole_list_anchors_round_trip_once(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source = _yaml("nested-list") + "- first ^inner-item\n\n  ```rust\n  body\n  ```\n\n- second\n\n^outer-list\n"
    path = tmp_path / "nested-list.md"
    path.write_text(source, encoding="utf-8", newline="")
    forward = _run(
        round_trip_runtime,
        "fenced-nested-list-forward",
        path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "nested-list-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()

    assert document_xml.count("docwen-fenced-source-v1:") == 1
    assert document_xml.count("docwen-anchor-v1:") == 2
    assert "bookmarkStart" not in document_xml
    assert " SEQ " not in document_xml
    assert " REF " not in document_xml
    returned = docx_to_md(
        round_trip_runtime,
        docx_path,
        tmp_path / "nested-list-reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False, "locale": "en"},
    )
    assert returned == source
    analysis = analyze_markdown_semantics_v3(returned, input_id="nested-list-returned.md")
    assert [(item["id"], item["block_kind"]) for item in analysis.projection["anchors"]] == [
        ("inner-item", "list_item"),
        ("outer-list", "list"),
    ]
    assert returned.count("^inner-item") == 1
    assert returned.count("^outer-list") == 1
    assert len(analysis.projection["fenced_sources"]) == 1


@pytest.mark.parametrize(
    ("case_id", "body"),
    [
        (
            "equal-physical-range",
            "> ```mermaid\n> graph TD\n> ```\n>\n> ^inner-fence\n\n^outer-quote\n",
        ),
        (
            "strict-physical-containment",
            "> before\n> ```mermaid\n> graph TD\n> ```\n>\n> ^inner-fence\n\n^outer-quote\n",
        ),
    ],
)
def test_nested_quote_fence_and_whole_quote_anchors_restore_exact_container_marker(
    tmp_path: Path,
    round_trip_runtime,
    case_id: str,
    body: str,
) -> None:
    source = _yaml(case_id) + body
    source_path = tmp_path / f"{case_id}.md"
    source_path.write_text(source, encoding="utf-8", newline="")

    forward = _run(
        round_trip_runtime,
        f"fenced-{case_id}-forward",
        source_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / f"{case_id}-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()

    assert document_xml.count("docwen-fenced-source-v1:") == 1
    assert document_xml.count("docwen-anchor-v1:") == 2
    assert "bookmarkStart" not in document_xml
    assert " SEQ " not in document_xml
    assert " REF " not in document_xml
    returned = docx_to_md(
        round_trip_runtime,
        docx_path,
        tmp_path / f"{case_id}-reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False, "locale": "en"},
    )

    assert returned == source
    assert returned.count("> ^inner-fence") == 1
    assert returned.count("\n^outer-quote") == 1
    analysis = analyze_markdown_semantics_v3(returned, input_id=f"{case_id}-returned.md")
    assert [(item["id"], item["block_kind"]) for item in analysis.projection["anchors"]] == [
        ("inner-fence", "fenced_block"),
        ("outer-quote", "block_quote"),
    ]

    inner = derive_anchor_identity_v3("code_block", "inner-fence")
    outer = derive_anchor_identity_v3("block_quote", "outer-quote")
    edge = derive_anchor_topology_edge_v3(inner.tag, outer.tag)
    with ZipFile(docx_path) as package:
        topology = etree.fromstring(_owned_custom_xml(package, ANCHOR_TOPOLOGY_MAP_NAMESPACE))
        [edge_element] = list(topology)
        assert tuple(edge_element.attrib.items()) == (
            ("child_tag", edge.child_tag),
            ("parent_tag", edge.parent_tag),
            ("sha256", edge.sha256),
        )
        document_root = etree.fromstring(package.read("word/document.xml"))
    tags = _semantic_sdt_tags(document_root)
    assert (
        tags.index(outer.tag)
        < tags.index(inner.tag)
        < next(index for index, tag in enumerate(tags) if tag.startswith("docwen-fenced-source-v1:"))
    )


@pytest.mark.parametrize(
    ("case_id", "body", "inner_id", "inner_kind", "outer_id", "outer_kind"),
    [
        (
            "list-in-quote",
            "> - one\n> - two\n>\n> ^inner-list\n\n^outer-quote\n",
            "inner-list",
            "list",
            "outer-quote",
            "block_quote",
        ),
        (
            "quote-in-list",
            "- > quote\n  > continuation\n\n  ^inner-quote\n\n^outer-list\n",
            "inner-quote",
            "block_quote",
            "outer-list",
            "list",
        ),
    ],
)
def test_container_anchor_topology_round_trips_at_exact_source_path(
    tmp_path: Path,
    round_trip_runtime,
    case_id: str,
    body: str,
    inner_id: str,
    inner_kind: str,
    outer_id: str,
    outer_kind: str,
) -> None:
    source = _yaml(case_id) + body
    source_path = tmp_path / f"{case_id}.md"
    source_path.write_text(source, encoding="utf-8", newline="")
    forward = _run(
        round_trip_runtime,
        f"topology-{case_id}-forward",
        source_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / f"{case_id}-forward",
    )
    docx_path = _primary_path(forward)
    inner = derive_anchor_identity_v3(inner_kind, inner_id)
    outer = derive_anchor_identity_v3(outer_kind, outer_id)
    expected_edge = derive_anchor_topology_edge_v3(inner.tag, outer.tag)
    with ZipFile(docx_path) as package:
        topology = etree.fromstring(_owned_custom_xml(package, ANCHOR_TOPOLOGY_MAP_NAMESPACE))
        [edge_element] = list(topology)
        assert tuple(edge_element.attrib.values()) == (
            expected_edge.child_tag,
            expected_edge.parent_tag,
            expected_edge.sha256,
        )
        document_root = etree.fromstring(package.read("word/document.xml"))
    tags = _semantic_sdt_tags(document_root)
    assert tags.index(outer.tag) < tags.index(inner.tag)
    assert not list(document_root.iter(qn("w:bookmarkStart")))
    assert not list(document_root.iter(qn("w:instrText")))

    returned = docx_to_md(
        round_trip_runtime,
        docx_path,
        tmp_path / f"{case_id}-reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False, "locale": "en"},
    )
    assert returned == source
    analysis = analyze_markdown_semantics_v3(returned, input_id=f"{case_id}-returned.md")
    assert [(item["id"], item["block_kind"]) for item in analysis.projection["anchors"]] == [
        (inner_id, inner_kind),
        (outer_id, outer_kind),
    ]
    assert returned.count(f"^{inner_id}") == 1
    assert returned.count(f"^{outer_id}") == 1


def test_disjoint_top_level_anchors_emit_no_topology_map(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source = _yaml("disjoint-top-level") + "Top A ^top-a\n\nTop B ^top-b\n"
    source_path = tmp_path / "disjoint-top-level.md"
    source_path.write_text(source, encoding="utf-8", newline="")
    forward = _run(
        round_trip_runtime,
        "topology-disjoint-forward",
        source_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "disjoint-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        package_payload = b"\n".join(package.read(name) for name in package.namelist())
    assert ANCHOR_TOPOLOGY_MAP_NAMESPACE.encode() not in package_payload
    assert (
        docx_to_md(
            round_trip_runtime,
            docx_path,
            tmp_path / "disjoint-reverse",
            options={"to_md_keep_images": False, "to_md_enable_ocr": False, "locale": "en"},
        )
        == source
    )


def test_anchor_topology_map_hash_tamper_fails_closed(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source_path = tmp_path / "topology-tamper.md"
    source_path.write_text(
        _yaml("topology-tamper") + "> ```mermaid\n> graph TD\n> ```\n>\n> ^inner-fence\n\n^outer-quote\n",
        encoding="utf-8",
        newline="",
    )
    forward = _run(
        round_trip_runtime,
        "topology-tamper-forward",
        source_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "topology-tamper-forward",
    )
    original = _primary_path(forward)
    tampered = tmp_path / "topology-tampered.docx"
    with ZipFile(original) as source, ZipFile(tampered, "w", ZIP_DEFLATED) as output:
        for item in source.infolist():
            payload = source.read(item.filename)
            if ANCHOR_TOPOLOGY_MAP_NAMESPACE.encode() in payload:
                payload = payload.replace(b'sha256="', b'sha256="0', 1)
            output.writestr(item, payload)

    reverse = _run(
        round_trip_runtime,
        "topology-tamper-reverse",
        tampered,
        source_format="docx",
        target_format="md",
        output_dir=tmp_path / "topology-tamper-reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False},
    )
    assert not reverse.success
    assert reverse.error is not None
    assert "topology" in reverse.error.message.casefold()


def test_fenced_source_package_tamper_fails_closed(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source_path = tmp_path / "tamper.md"
    source_path.write_text(_yaml("tamper") + "```rust\nbody\n```\n", encoding="utf-8", newline="")
    forward = _run(
        round_trip_runtime,
        "fenced-tamper-forward",
        source_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "tamper-forward",
    )
    original = _primary_path(forward)
    tampered = tmp_path / "tampered.docx"
    with ZipFile(original) as source, ZipFile(tampered, "w", ZIP_DEFLATED) as output:
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == "word/document.xml":
                payload = payload.replace(b">body<", b">evil<", 1)
            output.writestr(item, payload)

    reverse = _run(
        round_trip_runtime,
        "fenced-tamper-reverse",
        tampered,
        source_format="docx",
        target_format="md",
        output_dir=tmp_path / "tamper-reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False},
    )
    assert not reverse.success
    assert reverse.error is not None
    assert "fenced" in reverse.error.message.casefold()


def test_invalid_backtick_info_is_not_a_carrier_and_never_leaks_a_core_error(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source = _yaml("invalid-backtick-info") + "```rust`invalid\nbody\n```\n"
    source_path = tmp_path / "invalid-backtick-info.md"
    source_path.write_text(source, encoding="utf-8", newline="")

    forward = _run(
        round_trip_runtime,
        "invalid-backtick-info-forward",
        source_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "invalid-backtick-info-forward",
    )
    assert forward.success
    returned = docx_to_md(
        round_trip_runtime,
        _primary_path(forward),
        tmp_path / "invalid-backtick-info-reverse",
    )

    # CommonMark treats the invalid first line as paragraph text and the later
    # bare backticks as an independent omitted-EOF opener.  Only the latter is
    # carrier-owned; surrounding generic block whitespace is not source proof.
    assert "```rust`invalid\nbody" in returned
    [record] = analyze_markdown_semantics_v3(
        returned,
        input_id="invalid-backtick-info-returned.md",
    ).projection["fenced_sources"]
    assert record["info_b64"] == ""
    assert record["closing_state"] == "omitted_eof"


def _owned_custom_xml(package: ZipFile, namespace: str) -> bytes:
    matches = [
        package.read(name)
        for name in package.namelist()
        if re.fullmatch(r"customXml/item[1-9][0-9]*\.xml", name) and namespace.encode() in package.read(name)
    ]
    assert len(matches) == 1
    return matches[0]


def _semantic_sdt_tags(root: etree._Element, *, include_targets: bool = False) -> list[str]:
    tags: list[str] = []
    for sdt in root.iter(qn("w:sdt")):
        tag = sdt.find(f"./{qn('w:sdtPr')}/{qn('w:tag')}")
        if tag is None:
            continue
        value = tag.get(qn("w:val"))
        prefixes = (ANCHOR_TAG_PREFIX, "docwen-fenced-source-v1:")
        if include_targets:
            prefixes = (*prefixes, "docwen-target-v1:")
        if value is not None and value.startswith(prefixes):
            tags.append(value)
    return tags
