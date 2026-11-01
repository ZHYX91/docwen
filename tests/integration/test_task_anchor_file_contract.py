"""Real task-list packages preserve nested ordinary source ownership."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from tests.integration._round_trip_helper import docx_to_md
from tests.integration.test_markdown_block_heading_contract import _paragraph_text
from tests.integration.test_markdown_visible_content import _NS, _convert, _table_matrix, _text, _xml

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

_TARGET_NS = "https://docwen.dev/schema/document-target-map/v1"
_TOPOLOGY_NS = "https://docwen.dev/schema/document-anchor-topology-map/v1"


def _maps(path: Path) -> list[ET.Element]:
    with ZipFile(path) as package:
        return [
            ET.fromstring(package.read(name))
            for name in package.namelist()
            if name.startswith("customXml/item") and name.endswith(".xml")
        ]


def _anchor_records(path: Path) -> dict[str, dict[str, str]]:
    return {
        record.attrib["source_id"]: record.attrib
        for part in _maps(path)
        for record in part.findall(f"{{{_TARGET_NS}}}anchors/{{{_TARGET_NS}}}anchor")
    }


def _anchor_wrappers(root: ET.Element) -> dict[str, ET.Element]:
    result = {}
    for wrapper in root.findall(".//w:sdt", _NS):
        tag = wrapper.find("w:sdtPr/w:tag", _NS)
        if tag is None:
            continue
        value = tag.get(f"{{{_NS['w']}}}val", "")
        if value.startswith("docwen-anchor-v1:"):
            assert value not in result
            result[value] = wrapper
    return result


def _expected_anchor(source_id: str, kind: str) -> dict[str, str]:
    digest = hashlib.sha256(f"docwen-anchor-map-v1\0anchor\0{source_id}".encode()).hexdigest()
    return {"source_id": source_id, "block_kind": kind, "tag": "docwen-anchor-v1:" + digest[:32], "sha256": digest}


@pytest.mark.parametrize("state", [None, "[ ]", "[x]"], ids=["ordinary", "unchecked", "checked"])
@pytest.mark.parametrize("rich", [False, True], ids=["minimal", "complete-children"])
def test_task_parent_and_child_anchors_have_independent_physical_ranges(
    round_trip_runtime: Any, tmp_path: Path, state: str | None, rich: bool
) -> None:
    lead = (state + " " if state is not None else "") + "Lead"
    middle = (
        "  ```text\n  Literal body.\n  ```\n\n  | Key |\n  | --- |\n  | Value |\n\n  - Nested item.\n\n" if rich else ""
    )
    source = tmp_path / "tasks.md"
    authored = (f"- {lead} ^item\n\n" + middle + "  Continuation. ^child\n\nOutside.\n").encode()
    source.write_bytes(authored)
    output = _convert(round_trip_runtime, source, tmp_path / "forward", options={"markdown_extensions": {}})
    assert source.read_bytes() == authored

    root = _xml(output)
    records = _anchor_records(output)
    expected = {"item": _expected_anchor("item", "list_item"), "child": _expected_anchor("child", "paragraph")}
    assert records == expected
    maps = _maps(output)
    assert not [node for part in maps for node in part.findall(f"{{{_TARGET_NS}}}targets/{{{_TARGET_NS}}}target")]
    topology = [part for part in maps if part.tag == f"{{{_TOPOLOGY_NS}}}documentAnchorTopologyMap"]
    assert len(topology) == 1 and topology[0].attrib == {"version": "1"}
    child_tag = expected["child"]["tag"]
    parent_tag = expected["item"]["tag"]
    digest = hashlib.sha256(f"docwen-anchor-topology-edge-v1\0{child_tag}\0{parent_tag}".encode()).hexdigest()
    assert [edge.attrib for edge in topology[0].findall(f"{{{_TOPOLOGY_NS}}}edge")] == [
        {"child_tag": child_tag, "parent_tag": parent_tag, "sha256": digest}
    ]

    wrappers = _anchor_wrappers(root)
    assert set(wrappers) == {parent_tag, child_tag}
    parent_content = wrappers[parent_tag].find("w:sdtContent", _NS)
    child_content = wrappers[child_tag].find("w:sdtContent", _NS)
    assert parent_content is not None and child_content is not None
    assert [block.tag for block in child_content] == [f"{{{_NS['w']}}}p"]
    assert _paragraph_text(child_content[0]) == "Continuation."
    assert child_content.find(".//w:sdt", _NS) is None
    assert wrappers[child_tag] is parent_content[-1]
    expected_tags = [f"{{{_NS['w']}}}p"]
    expected_texts = [lead]
    if rich:
        expected_tags.extend([f"{{{_NS['w']}}}p", f"{{{_NS['w']}}}tbl", f"{{{_NS['w']}}}p"])
        expected_texts.extend(["Literal body.\n", "Key", "Value", "Nested item."])
        assert _table_matrix(parent_content[2]) == [["Key"], ["Value"]]
    expected_tags.append(f"{{{_NS['w']}}}sdt")
    expected_texts.append("Continuation.")
    assert [block.tag for block in parent_content] == expected_tags
    assert [_paragraph_text(p) for p in parent_content.findall(".//w:p", _NS)] == expected_texts
    assert parent_content[0].find("w:pPr/w:numPr/w:numId", _NS) is not None
    assert root.find(".//w:bookmarkStart", _NS) is None
    assert root.find(".//w:instrText", _NS) is None
    assert "^item" not in _text(root) and "^child" not in _text(root) and "DOCWENV3" not in _text(root)
    outside = [p for p in root.findall(".//w:p", _NS) if _paragraph_text(p) == "Outside."]
    assert len(outside) == 1 and outside[0] not in list(parent_content.iter())
    body = root.find("w:body", _NS)
    assert body is not None
    assert list(body).index(wrappers[parent_tag]) < list(body).index(outside[0])

    returned = docx_to_md(round_trip_runtime, output, tmp_path / "reverse", options={"markdown_extensions": {}})
    visible = [value.removesuffix("\n") for value in expected_texts] + ["Outside."]
    assert all(returned.count(value) == 1 for value in visible)
    assert [returned.index(value) for value in visible] == sorted(returned.index(value) for value in visible)
    assert returned.count("^item") == 1 and returned.count("^child") == 1
    assert "Continuation. ^child" in returned
    assert returned.index(lead) < returned.index("^item") < returned.index("Outside.")


@pytest.mark.parametrize("state", ["[ ]", "[x]"], ids=["unchecked", "checked"])
def test_task_continuation_id_does_not_claim_the_unanchored_lead(
    round_trip_runtime: Any, tmp_path: Path, state: str
) -> None:
    source = tmp_path / "child-only.md"
    source.write_text(f"- {state} Lead\n\n  Continuation. ^child\n\nOutside.\n", encoding="utf-8")
    output = _convert(round_trip_runtime, source, tmp_path / "forward", options={"markdown_extensions": {}})
    root = _xml(output)
    expected = _expected_anchor("child", "paragraph")
    assert _anchor_records(output) == {"child": expected}
    wrappers = _anchor_wrappers(root)
    assert set(wrappers) == {expected["tag"]}
    content = wrappers[expected["tag"]].find("w:sdtContent", _NS)
    assert content is not None
    assert [child.tag for child in content] == [f"{{{_NS['w']}}}p"]
    assert _paragraph_text(content[0]) == "Continuation."
    assert [_paragraph_text(p) for p in root.findall(".//w:p", _NS) if _paragraph_text(p)] == [
        f"{state} Lead",
        "Continuation.",
        "Outside.",
    ]
    assert not [part for part in _maps(output) if part.tag == f"{{{_TOPOLOGY_NS}}}documentAnchorTopologyMap"]
    assert root.find(".//w:bookmarkStart", _NS) is None and root.find(".//w:instrText", _NS) is None
    returned = docx_to_md(round_trip_runtime, output, tmp_path / "reverse", options={"markdown_extensions": {}})
    assert f"{state} Lead" in returned and "Continuation. ^child" in returned
    assert returned.count("^child") == 1 and "^item" not in returned
    assert returned.index("Lead") < returned.index("Continuation.") < returned.index("Outside.")
