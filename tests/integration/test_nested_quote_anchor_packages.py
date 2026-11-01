"""Default-dialect quote anchors retain native content and authenticated nesting."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from docwen_core.markdown_extensions import MarkdownExtensions
from tests.integration._round_trip_helper import docx_to_md
from tests.integration.test_markdown_visible_content import _NS, _convert, _text, _xml

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

_MAP_NS = "https://docwen.dev/schema/document-target-map/v1"
_TOPOLOGY_NS = "https://docwen.dev/schema/document-anchor-topology-map/v1"


def _write(runtime: Any, tmp_path: Path, source: str) -> Path:
    path = tmp_path / "source.md"
    path.write_bytes(source.encode("utf-8"))
    return _convert(
        runtime,
        path,
        tmp_path / "forward",
        options={"markdown_extensions": {"input": MarkdownExtensions().to_dict()}},
    )


def _read(runtime: Any, path: Path, tmp_path: Path) -> str:
    return docx_to_md(
        runtime,
        path,
        tmp_path / "reverse",
        options={"markdown_extensions": {"output": MarkdownExtensions().to_dict()}},
    )


def _map_roots(path: Path) -> list[ET.Element]:
    with ZipFile(path) as package:
        return [
            ET.fromstring(package.read(name))
            for name in package.namelist()
            if name.startswith("customXml/item") and name.endswith(".xml")
        ]


def _identity(source_id: str, kind: str) -> dict[str, str]:
    digest = hashlib.sha256(f"docwen-anchor-map-v1\0anchor\0{source_id}".encode()).hexdigest()
    return {"block_kind": kind, "source_id": source_id, "tag": "docwen-anchor-v1:" + digest[:32], "sha256": digest}


@pytest.mark.parametrize(
    ("source", "native_texts", "owner_kind"),
    [
        ("> > text ^id\n", ["text"], "paragraph"),
        ("> > > text ^id\n\nOutside.\n", ["text", "Outside."], "paragraph"),
        ("> - text ^id\n\nOutside.\n", ["text", "Outside."], "list_item"),
        ("- > text ^id\n\nOutside.\n", ["text", "Outside."], "paragraph"),
        ("> a\n> > text ^id\n\nOutside.\n", ["a", "text", "Outside."], "paragraph"),
        ("\ufeff> > text ^id\r\n\r\nOutside.\r\n", ["text", "Outside."], "paragraph"),
    ],
    ids=["original", "triple-quote", "quote-list", "list-quote", "short-inner", "bom-crlf"],
)
def test_default_dialect_nested_quote_anchor_has_exact_native_content_before_reading(
    tmp_path: Path, round_trip_runtime: Any, source: str, native_texts: list[str], owner_kind: str
) -> None:
    path = _write(round_trip_runtime, tmp_path, source)
    root = _xml(path)
    assert [_text(p) for p in root.findall(".//w:p", _NS) if _text(p)] == native_texts
    assert root.find(".//w:bookmarkStart", _NS) is None
    assert root.find(".//w:instrText", _NS) is None
    assert "^id" not in _text(root) and "DOCWENV3" not in _text(root)
    expected = _identity("id", owner_kind)
    wrappers = [
        node
        for node in root.findall(".//w:sdt", _NS)
        if (tag := node.find("w:sdtPr/w:tag", _NS)) is not None and tag.get(f"{{{_NS['w']}}}val") == expected["tag"]
    ]
    assert len(wrappers) == 1
    content = wrappers[0].find("w:sdtContent", _NS)
    assert content is not None
    assert [child.tag for child in content] == [f"{{{_NS['w']}}}p"]
    assert _text(content) == "text"
    assert "Outside." not in _text(content) and content.find(".//w:sdt", _NS) is None
    maps = _map_roots(path)
    assert [record.attrib for item in maps for record in item.findall(f"{{{_MAP_NS}}}anchors/{{{_MAP_NS}}}anchor")] == [
        expected
    ]
    assert [record for item in maps for record in item.findall(f"{{{_MAP_NS}}}targets/{{{_MAP_NS}}}target")] == []

    returned = _read(round_trip_runtime, path, tmp_path)
    assert returned.count("text ^id") == 1
    assert returned.count("^id") == 1
    assert returned.count("Outside.") == int("Outside." in native_texts)
    if "Outside." in native_texts:
        assert returned.index("text ^id") < returned.index("Outside.")
    assert (tmp_path / "source.md").read_bytes() == source.encode("utf-8")


@pytest.mark.parametrize(
    ("source", "inner_kind"),
    [
        ("> > text ^inner\n\n^outer\n\nOutside.\n", "paragraph"),
        ("> > text\n>\n> ^inner\n\n^outer\n\nOutside.\n", "block_quote"),
    ],
    ids=["equal-range-ancestors", "nested-container-id"],
)
def test_quote_parent_and_child_keep_their_physical_edge_in_a_fresh_package(
    tmp_path: Path, round_trip_runtime: Any, source: str, inner_kind: str
) -> None:
    path = _write(round_trip_runtime, tmp_path, source)
    root = _xml(path)
    assert [_text(p) for p in root.findall(".//w:p", _NS) if _text(p)] == ["text", "Outside."]
    assert root.find(".//w:bookmarkStart", _NS) is None and root.find(".//w:instrText", _NS) is None
    expected = [_identity("inner", inner_kind), _identity("outer", "block_quote")]
    by_tag = {
        tag.get(f"{{{_NS['w']}}}val"): wrapper
        for wrapper in root.findall(".//w:sdt", _NS)
        if (tag := wrapper.find("w:sdtPr/w:tag", _NS)) is not None
    }
    assert set(by_tag) == {item["tag"] for item in expected}
    inner, outer = (by_tag[item["tag"]] for item in expected)
    inner_content = inner.find("w:sdtContent", _NS)
    outer_content = outer.find("w:sdtContent", _NS)
    assert inner_content is not None and outer_content is not None
    assert list(outer_content) == [inner]
    assert [node.tag for node in inner_content] == [f"{{{_NS['w']}}}p"]
    assert _text(inner_content) == "text" and _text(outer_content) == "text"
    outside = [p for p in root.findall(".//w:p", _NS) if _text(p) == "Outside."]
    assert len(outside) == 1 and outside[0] not in list(outer.iter())
    maps = _map_roots(path)
    records = [record.attrib for item in maps for record in item.findall(f"{{{_MAP_NS}}}anchors/{{{_MAP_NS}}}anchor")]
    assert sorted(records, key=lambda item: item["source_id"]) == expected
    [topology] = [item for item in maps if item.tag == f"{{{_TOPOLOGY_NS}}}documentAnchorTopologyMap"]
    child_tag, parent_tag = (item["tag"] for item in expected)
    digest = hashlib.sha256(f"docwen-anchor-topology-edge-v1\0{child_tag}\0{parent_tag}".encode()).hexdigest()
    assert [edge.attrib for edge in topology.findall(f"{{{_TOPOLOGY_NS}}}edge")] == [
        {"child_tag": child_tag, "parent_tag": parent_tag, "sha256": digest}
    ]

    returned = _read(round_trip_runtime, path, tmp_path)
    assert returned.count("text") == returned.count("Outside.") == 1
    assert returned.count("^inner") == returned.count("^outer") == 1
    assert returned.index("text") < returned.index("^inner") < returned.index("^outer") < returned.index("Outside.")
