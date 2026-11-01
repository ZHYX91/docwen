"""Native trailing separators preserve the complete authenticated parent range."""

from __future__ import annotations

import hashlib
from io import BytesIO
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from markdown_it import MarkdownIt

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_core.docx_semantics_v3 import DocxSemanticsV3Recovery, DocxSemanticsV3Session
from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer import MdToDocxRenderer
from docwen_plugin_markdown.runtime_semantics_v3 import apply_runtime_semantics_v3, prepare_runtime_semantics_v3
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
from tests.integration._round_trip_helper import _run
from tests.integration.test_nested_quote_anchor_packages import _identity
from tests.integration.test_semantic_quote_break_recovery import _variant

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


@pytest.mark.parametrize("captions", [False, True], ids=["captions-off", "captions-on"])
@pytest.mark.parametrize("ending", ["none", "page", "section"])
@pytest.mark.parametrize("container", ["list-closed", "quote-remains-open"])
def test_nested_quote_separator_keeps_complete_parent_and_child_ranges(
    tmp_path: Path, round_trip_runtime: Any, captions: bool, ending: str, container: str
) -> None:
    if container == "list-closed":
        source = "# Target\n\n- > Before @[[#Target]]<br>After ^quote-id\n\n  ^inner\n\n^outer\n"
        outer_kind = "list"
        expected_outer_text = "Before 1After"
        expected_groups = [("quote-id", "paragraph", 0, 1), ("inner", "block_quote", 0, 1), ("outer", "list", 0, 1)]
    else:
        source = (
            "# Target\n\n> Prelude\n>\n> > Before @[[#Target]]<br>After ^quote-id\n>\n> ^inner\n>\n> Tail\n\n^outer\n"
        )
        outer_kind = "block_quote"
        expected_outer_text = "PreludeBefore 1AfterTail"
        expected_groups = [
            ("quote-id", "paragraph", 0, 1),
            ("inner", "block_quote", 0, 1),
            ("outer", "block_quote", 1, 3),
        ]
    dialect = MarkdownExtensions(captions_references=True)
    plan = prepare_runtime_semantics_v3(
        source, input_id="source.md", extensions=dialect, consumer_profile="number_suite_direct"
    )
    assert not plan.analysis.has_errors
    source_file = tmp_path / "source.md"
    source_file.write_bytes(source.encode())
    document = Document()
    session = DocxSemanticsV3Session(document, source_sha256=plan.source_sha256, standalone_caption_authority=True)
    parsed = parse_markdown_text(plan.shielded_source, extensions=dialect, auto_link_bare_url=False)
    assert isinstance(parsed, list)
    MdToDocxRenderer(doc=document, semantic_v3_session=session).render(apply_runtime_semantics_v3(parsed, plan))
    session.finalize_document()
    stream = BytesIO()
    document.save(stream)
    original = tmp_path / "writer.docx"
    original.write_bytes(stream.getvalue())
    session.write_package(original)
    original_bytes = original.read_bytes()
    native = tmp_path / "native.docx"
    _variant(original, native, break_tag="br", ending=ending)
    assert original.read_bytes() == original_bytes
    assert source_file.read_bytes() == source.encode()

    with ZipFile(native) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    tags = {
        tag.get(qn("w:val")): owner
        for owner in root.iter(qn("w:sdt"))
        if (tag := owner.find(f"{qn('w:sdtPr')}/{qn('w:tag')}")) is not None
    }
    child = tags[_identity("quote-id", "paragraph")["tag"]]
    inner = tags[_identity("inner", "block_quote")["tag"]]
    outer = tags[_identity("outer", outer_kind)["tag"]]
    assert child in inner.iter() and inner in outer.iter()
    assert "".join(node.text or "" for node in child.iter(qn("w:t"))) == "Before 1After"
    assert "".join(node.text or "" for node in outer.iter(qn("w:t"))) == expected_outer_text
    assert not list(root.iter(qn("w:bookmarkStart")))
    assert not list(root.iter(qn("w:instrText")))
    assert len([node for node in child.iter(qn("w:br")) if node.get(qn("w:type")) == "page"]) == int(ending == "page")

    fresh = Document(BytesIO(native.read_bytes()))
    authenticated = DocxSemanticsV3Recovery.load(native, fresh)
    [paragraph] = [
        item
        for item in fresh.element.body.iter(qn("w:p"))
        if "Before " in "".join(node.text or "" for node in item.iter(qn("w:t")))
    ]
    assert [
        (group.anchor.source_id, group.anchor.block_kind, group.index, len(group.elements))
        for group in authenticated.ordinary_anchor_groups(paragraph)
    ] == expected_groups
    result = _run(
        round_trip_runtime,
        "nested-native-break",
        native,
        source_format="docx",
        target_format="md",
        output_dir=tmp_path / "reverse",
        options={
            "remove_numbering": False,
            "markdown_extensions": {"output": MarkdownExtensions(captions_references=captions).to_dict()},
            "page_break_separator": "---",
            "section_break_separator": "***",
        },
    )
    assert result.success, result.error
    draft = build_bundle_draft(
        profile="document_with_resources", output_media_type="text/markdown", artifacts=result.artifacts
    )
    bundle = ArtifactBundleCommitter().commit(
        task_id="nested-native-break", staging_root=str(tmp_path / "reverse"), draft=draft
    )
    [preferred] = [entry.artifact_id for entry in bundle.entries if entry.preferred]
    [artifact] = [item for item in bundle.artifacts if item.artifact_id == preferred]
    data = (tmp_path / "reverse" / artifact.locator).read_bytes()
    assert artifact.sha256 == hashlib.sha256(data).hexdigest()
    returned = data.decode("utf-8-sig")
    body = returned.split("\n---\n", 1)[1] if returned.startswith("---\n") else returned
    reference = "@[[#Target]]" if captions else "1"
    assert f"Before {reference}<br>After ^quote-id" in body
    assert (
        body.count("Before")
        == body.count("After")
        == body.count("^quote-id")
        == body.count("^inner")
        == body.count("^outer")
        == 1
    )
    assert body.count("@[[#Target]]") == int(captions)
    tokens = MarkdownIt("commonmark").parse(body)
    assert not [token for token in tokens if token.type in {"code_block", "fence"}]
    separators = [token for token in tokens if token.type == "hr"]
    assert len(separators) == int(ending != "none")
    if separators:
        assert separators[0].level == (0 if container == "list-closed" else 1)
    if container == "list-closed":
        assert [
            (token.type, token.level) for token in tokens if token.type in {"bullet_list_open", "blockquote_open"}
        ] == [("bullet_list_open", 0), ("blockquote_open", 2)]
    else:
        assert [(token.type, token.level) for token in tokens if token.type == "blockquote_open"] == [
            ("blockquote_open", 0),
            ("blockquote_open", 1),
        ]

    recovered = prepare_runtime_semantics_v3(
        returned, input_id="recovered.md", extensions=dialect, consumer_profile="number_suite_direct"
    )
    assert not recovered.analysis.has_errors
    anchors = {item["id"]: item for item in recovered.analysis.projection["anchors"]}
    assert {key: item["block_kind"] for key, item in anchors.items()} == {
        "quote-id": "paragraph",
        "inner": "block_quote",
        "outer": outer_kind,
    }
    ranges = {key: item["block_range"] for key, item in anchors.items()}
    for key in ("quote-id", "inner"):
        assert ranges["outer"]["start"] <= ranges[key]["start"] < ranges[key]["end"] <= ranges["outer"]["end"]
        owned = returned[ranges[key]["start"] : ranges[key]["end"]]
        assert "Before" in owned and "After" in owned
        assert "Prelude" not in owned and "Tail" not in owned
    outer_source = returned[ranges["outer"]["start"] : ranges["outer"]["end"]]
    assert "Before" in outer_source and "After" in outer_source
    if container == "quote-remains-open":
        assert "Prelude" in outer_source and "Tail" in outer_source
        assert (
            outer_source.index("Prelude")
            < outer_source.index("Before")
            < outer_source.index("After")
            < outer_source.index("Tail")
        )
