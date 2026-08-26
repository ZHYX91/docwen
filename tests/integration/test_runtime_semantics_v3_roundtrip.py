"""Production-composition round trips for Markdown semantics v3."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from PIL import Image

from docwen_core.docx_semantics_v3 import (
    ANCHOR_TAG_PREFIX,
    REFERENCE_OCCURRENCE_MAP_NAMESPACE,
    SOFT_REFERENCE_MAP_NAMESPACE,
    TARGET_MAP_NAMESPACE,
)
from docwen_plugin_markdown.document_semantics_v3 import analyze_markdown_semantics_v3
from tests.integration._round_trip_helper import _primary_path, _run, docx_to_md

pytestmark = pytest.mark.integration


def test_production_heading_anchor_reference_citation_round_trip(
    tmp_path: Path,
) -> None:
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_runtime.config.loader import ConfigLoader

    config_loader = ConfigLoader(
        base_dir=Path(__file__).resolve().parents[2] / "configs",
        user_dir=tmp_path / "lossless-link-config",
    )
    assert config_loader.set_values(
        {
            "link.non_embed_links.wiki_mode": "keep",
            "link.embed_links.wiki_image_mode": "keep",
            "link.embed_links.md_file_mode": "keep",
            "link.error_handling.file_not_found": "keep",
        }
    )
    round_trip_runtime = create_runtime_port(config_loader=config_loader)
    source = """# Intro ^intro

# Other

Plain block ^raw

See [[Page#^raw]], ![[Page#^raw]], @[[#^intro|Intro]], @[[#Other|Other]], and @fig-legacy.
"""
    markdown_path = tmp_path / "source.md"
    markdown_path.write_text(source, encoding="utf-8")
    forward = _run(
        round_trip_runtime,
        "v3-production-forward",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
        package_text = b"\n".join(package.read(name) for name in package.namelist()).decode(errors="ignore")
    assert "SEQ " not in document_xml
    assert " REF DW_T_" in document_xml
    assert ANCHOR_TAG_PREFIX in document_xml
    assert TARGET_MAP_NAMESPACE in package_text
    assert SOFT_REFERENCE_MAP_NAMESPACE in package_text
    assert REFERENCE_OCCURRENCE_MAP_NAMESPACE in package_text

    returned = docx_to_md(
        round_trip_runtime,
        docx_path,
        tmp_path / "reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False},
    )
    assert "# Intro ^intro" in returned
    assert "Plain block ^raw" in returned
    assert ("See [[Page#^raw]], ![[Page#^raw]], @[[#^intro|Intro]], @[[#Other|Other]], and @fig-legacy.") in returned


def test_production_invalid_anchor_keeps_source_evidence(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    markdown_path = tmp_path / "invalid.md"
    markdown_path.write_text("Plain ^bad_id\n", encoding="utf-8")
    result = _run(
        round_trip_runtime,
        "v3-invalid",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "invalid-output",
    )
    assert not result.success
    [diagnostic] = result.diagnostics
    payload = diagnostic.to_dict()
    assert payload["code"] == "docwen.markdown.anchor.invalid_id"
    assert payload["source"]["coordinate_system"] == "unicode_code_point"
    assert payload["range"] == {"start": 6, "end": 13}


@pytest.mark.parametrize(
    ("kind", "declaration", "object_markdown", "expected_object_token"),
    [
        ("Figure", "Figure: Image caption", "![pixel](pixel.png)", "![pixel](pixel.png)"),
        ("Table", "Table: Results", "| A | B |\n|---|---|\n| 1 | 2 |", "| A | B |"),
        ("Equation", "Equation: Energy", "$$\nE = mc^2\n$$", "E=mc^{2}"),
        ("Code", "Code: Example", "```rust\nfn main() {}\n```", "fn main() {}"),
    ],
)
@pytest.mark.parametrize("target_id", ["target-id", None])
def test_production_caption_matrix_round_trip(
    tmp_path: Path,
    round_trip_runtime,
    kind: str,
    declaration: str,
    object_markdown: str,
    expected_object_token: str,
    target_id: str | None,
) -> None:
    source_declaration = declaration + (f" ^{target_id}" if target_id else "")
    source = f"{source_declaration}\n\n{object_markdown}\n"
    markdown_path = tmp_path / f"{kind.lower()}-{target_id or 'idless'}.md"
    markdown_path.write_text(source, encoding="utf-8")
    if kind == "Figure":
        with Image.new("RGB", (2, 2), (32, 96, 160)) as image:
            image.save(tmp_path / "pixel.png", format="PNG")
    forward = _run(
        round_trip_runtime,
        f"v3-{kind}-{target_id or 'idless'}-forward",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / f"{kind}-{target_id or 'idless'}-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
        package_text = b"\n".join(package.read(name) for name in package.namelist()).decode(errors="ignore")
    assert f" SEQ {kind} " in document_xml
    if target_id is None:
        assert "DW_T_" not in document_xml
        assert f'source_id="{target_id}"' not in package_text
    else:
        assert "DW_T_" in document_xml
        assert f'source_id="{target_id}"' in package_text

    returned = docx_to_md(
        round_trip_runtime,
        docx_path,
        tmp_path / f"{kind}-{target_id or 'idless'}-reverse",
        options={"to_md_keep_images": True, "to_md_enable_ocr": False},
    )
    assert source_declaration in returned
    assert expected_object_token in returned


def test_production_whole_list_anchor_round_trip(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source = "- one\n- two\n\n^whole-list\n"
    markdown_path = tmp_path / "whole-list.md"
    markdown_path.write_text(source, encoding="utf-8")
    forward = _run(
        round_trip_runtime,
        "v3-list-forward",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "whole-list-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
    assert ANCHOR_TAG_PREFIX in document_xml
    assert "bookmarkStart" not in document_xml
    assert " SEQ " not in document_xml
    assert " REF " not in document_xml

    returned = docx_to_md(round_trip_runtime, docx_path, tmp_path / "whole-list-reverse")
    assert "- one" in returned
    assert "- two" in returned
    assert "^whole-list" in returned


def test_production_inline_image_anchor_round_trip(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source = "![pixel](pixel.png) ^image-id\n"
    markdown_path = tmp_path / "image-anchor.md"
    markdown_path.write_text(source, encoding="utf-8")
    with Image.new("RGB", (2, 2), (32, 96, 160)) as image:
        image.save(tmp_path / "pixel.png", format="PNG")

    forward = _run(
        round_trip_runtime,
        "v3-image-anchor-forward",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "image-anchor-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
    assert ANCHOR_TAG_PREFIX in document_xml
    assert "bookmarkStart" not in document_xml
    assert " SEQ " not in document_xml
    assert " REF " not in document_xml

    returned = docx_to_md(
        round_trip_runtime,
        docx_path,
        tmp_path / "image-anchor-reverse",
        options={"to_md_keep_images": True, "to_md_enable_ocr": False},
    )
    analysis = analyze_markdown_semantics_v3(returned, input_id="returned.md")
    assert not analysis.has_errors
    assert [(item["id"], item["block_kind"], item["placement"]) for item in analysis.projection["anchors"]] == [
        ("image-id", "image", "inline")
    ]


def test_production_each_list_item_anchor_round_trip(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source = "- first ^first-item\n- middle ^middle-item\n- last ^last-item\n"
    markdown_path = tmp_path / "list-item-anchors.md"
    markdown_path.write_text(source, encoding="utf-8")

    forward = _run(
        round_trip_runtime,
        "v3-list-item-anchors-forward",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "list-item-anchors-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
    assert document_xml.count(ANCHOR_TAG_PREFIX) == 3
    assert "bookmarkStart" not in document_xml
    assert " SEQ " not in document_xml
    assert " REF " not in document_xml

    returned = docx_to_md(round_trip_runtime, docx_path, tmp_path / "list-item-anchors-reverse")
    analysis = analyze_markdown_semantics_v3(returned, input_id="returned.md")
    assert not analysis.has_errors
    assert [(item["id"], item["block_kind"], item["placement"]) for item in analysis.projection["anchors"]] == [
        ("first-item", "list_item", "inline"),
        ("middle-item", "list_item", "inline"),
        ("last-item", "list_item", "inline"),
    ]


@pytest.mark.parametrize(
    ("case_id", "source", "anchor_id", "block_kind"),
    [
        ("quote", "> first\n> second\n\n^quote-id\n", "quote-id", "block_quote"),
        ("callout", "> [!NOTE] Title\n> body\n\n^callout-id\n", "callout-id", "callout"),
        ("table", "| A | B |\n|---|---|\n| 1 | 2 |\n\n^table-id\n", "table-id", "table"),
        ("equation", "$$\nE = mc^2\n$$\n\n^equation-id\n", "equation-id", "equation"),
    ],
)
def test_production_structured_ordinary_anchor_matrix_round_trip(
    tmp_path: Path,
    round_trip_runtime,
    case_id: str,
    source: str,
    anchor_id: str,
    block_kind: str,
) -> None:
    markdown_path = tmp_path / f"{case_id}.md"
    markdown_path.write_text(source, encoding="utf-8")
    forward = _run(
        round_trip_runtime,
        f"v3-{case_id}-forward",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / f"{case_id}-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
    assert ANCHOR_TAG_PREFIX in document_xml
    assert "bookmarkStart" not in document_xml
    assert " SEQ " not in document_xml
    assert " REF " not in document_xml

    returned = docx_to_md(round_trip_runtime, docx_path, tmp_path / f"{case_id}-reverse")
    analysis = analyze_markdown_semantics_v3(returned, input_id="returned.md")
    assert not analysis.has_errors
    assert [(item["id"], item["block_kind"], item["placement"]) for item in analysis.projection["anchors"]] == [
        (anchor_id, block_kind, "post_block")
    ]


def test_production_caption_target_and_raw_table_anchor_remain_distinct(
    tmp_path: Path,
    round_trip_runtime,
) -> None:
    source = "Table: Results ^table-target\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n^raw-table\n"
    markdown_path = tmp_path / "caption-and-raw-anchor.md"
    markdown_path.write_text(source, encoding="utf-8")
    forward = _run(
        round_trip_runtime,
        "v3-caption-and-raw-anchor-forward",
        markdown_path,
        source_format="markdown",
        target_format="docx",
        output_dir=tmp_path / "caption-and-raw-anchor-forward",
    )
    docx_path = _primary_path(forward)
    with ZipFile(docx_path) as package:
        document_xml = package.read("word/document.xml").decode()
        package_text = b"\n".join(package.read(name) for name in package.namelist()).decode(errors="ignore")
    assert document_xml.count(" SEQ Table ") == 1
    assert document_xml.count("bookmarkStart") == 1
    assert document_xml.count(ANCHOR_TAG_PREFIX) == 1
    assert 'source_id="table-target"' in package_text
    assert 'source_id="raw-table"' in package_text

    returned = docx_to_md(round_trip_runtime, docx_path, tmp_path / "caption-and-raw-anchor-reverse")
    analysis = analyze_markdown_semantics_v3(returned, input_id="returned.md")
    assert not analysis.has_errors
    assert [(item["id"], item["kind"]) for item in analysis.projection["targets"]] == [("table-target", "table")]
    assert [(item["id"], item["block_kind"]) for item in analysis.projection["anchors"]] == [("raw-table", "table")]
