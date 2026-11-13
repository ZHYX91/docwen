"""Ordinary heading navigation stays separate from semantic block anchors."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn

from ._link_processing_routes_support import _convert_docx, _link_config
from .conftest import make_context

pytestmark = pytest.mark.contract


def _anchors(xml: str) -> list[str]:
    root = ET.fromstring(xml)
    return [node.attrib[qn("w:anchor")] for node in root.iter(qn("w:hyperlink")) if qn("w:anchor") in node.attrib]


def _navigation_bookmarks(xml: str) -> dict[str, str]:
    root = ET.fromstring(xml)
    return {
        node.attrib[qn("w:name")]: node.attrib[qn("w:id")]
        for node in root.iter(qn("w:bookmarkStart"))
        if node.attrib.get(qn("w:name"), "").startswith("DWNav_")
    }


@pytest.mark.parametrize(
    ("heading", "links"),
    [
        ("Overview", "[Markdown](#overview) and [[#Overview|Wiki]]"),
        ("项目计划", "[Markdown](#%E9%A1%B9%E7%9B%AE%E8%AE%A1%E5%88%92) and [[#项目计划|Wiki]]"),
        ("Project Plan", "[Markdown](#project-plan) and [[#Project Plan|Wiki]]"),
    ],
)
def test_markdown_and_wiki_heading_links_share_real_bookmark(tmp_path: Path, heading: str, links: str) -> None:
    source = tmp_path / "heading.md"
    original = f"{links}\n\n# {heading}\n\nBody.\n".encode()
    source.write_bytes(original)
    observation = _convert_docx(source, _link_config())

    anchors = _anchors(observation.document_xml)
    bookmarks = _navigation_bookmarks(observation.document_xml)
    assert len(anchors) == 2
    assert len(set(anchors)) == 1
    assert set(anchors) == set(bookmarks)
    assert len(bookmarks) == 1
    assert not observation.hyperlink_targets
    assert "docwen-navigation:" not in observation.document_xml
    assert "Markdown" in observation.text and "Wiki" in observation.text
    assert source.read_bytes() == original


@pytest.mark.parametrize("heading", ["# Repeated\n\n# Repeated", "# Another"])
def test_duplicate_and_unknown_heading_links_extract_text_by_default(tmp_path: Path, heading: str) -> None:
    source = tmp_path / "unknown.md"
    links = '[**Markdown**](#repeated "title") and [[#Repeated|Wiki]]'
    source.write_text(f"{links}\n\n{heading}\n", encoding="utf-8")
    observation = _convert_docx(source, _link_config())

    assert "Markdown and Wiki" in observation.text
    assert not _anchors(observation.document_xml)
    assert not _navigation_bookmarks(observation.document_xml)
    assert not observation.hyperlink_targets
    assert "docwen-navigation:" not in observation.document_xml


@pytest.mark.parametrize("fallback", ["keep", "extract_text", "extract_address"])
def test_unknown_heading_fallback_is_syntax_specific(tmp_path: Path, fallback: str) -> None:
    source = tmp_path / "fallback.md"
    source.write_text("[Markdown](#missing) / [[#Absent|Wiki]]\n", encoding="utf-8")
    config = _link_config()
    config["link"]["non_embed_links"].update(
        markdown_hyperlink_fallback=fallback,
        wiki_hyperlink_fallback="extract_address" if fallback != "extract_address" else "extract_text",
    )
    observation = _convert_docx(source, config)

    expected_markdown = {"keep": "[Markdown](#missing)", "extract_text": "Markdown", "extract_address": "#missing"}[
        fallback
    ]
    expected_wiki = "#Absent" if fallback != "extract_address" else "Wiki"
    assert f"{expected_markdown} / {expected_wiki}" in observation.text
    assert not _anchors(observation.document_xml)
    assert not observation.hyperlink_targets


def test_ordinary_block_id_link_does_not_bind_frozen_semantic_anchor(tmp_path: Path) -> None:
    source = tmp_path / "block.md"
    source.write_text("[[#^intro|Ordinary]]\n\n# Intro ^intro\n", encoding="utf-8")
    observation = _convert_docx(source, _link_config())

    assert "Ordinary" in observation.text
    assert "[[#^intro|Ordinary]]" not in observation.text
    assert not _anchors(observation.document_xml)
    assert not _navigation_bookmarks(observation.document_xml)
    assert not observation.hyperlink_targets


def test_heading_bookmark_and_internal_hyperlink_survive_docx_save_reload(tmp_path: Path) -> None:
    from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

    source = tmp_path / "saved.md"
    source.write_text("[Go](#target)\n\n# Target\n", encoding="utf-8")
    context, _workspace = make_context(str(source), target_format="docx", config_values=_link_config())
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    document = Document(result.artifacts[0].staging_path)
    original_xml = document.element.xml
    stream = BytesIO()
    document.save(stream)
    stream.seek(0)
    reopened = Document(stream)

    assert _anchors(reopened.element.xml) == _anchors(original_xml)
    assert _navigation_bookmarks(reopened.element.xml) == _navigation_bookmarks(original_xml)
    assert set(_anchors(reopened.element.xml)) == set(_navigation_bookmarks(reopened.element.xml))
    assert all(not relationship.reltype.endswith("/hyperlink") for relationship in reopened.part.rels.values())


@pytest.mark.parametrize("kind", ["markdown", "wiki"])
@pytest.mark.parametrize("matching", [True, False])
def test_yaml_heading_links_use_rendered_body_targets_or_default_text_fallback(
    tmp_path: Path, kind: str, matching: bool
) -> None:
    source = tmp_path / "yaml.md"
    link = "[YAML jump](#overview)" if kind == "markdown" else "[[#Overview|YAML jump]]"
    heading = "Overview" if matching else "Other"
    source.write_text(f"---\nsite: '{link}'\n---\n\n# {heading}\n", encoding="utf-8")
    template = Document()
    paragraph = template.add_paragraph("Site: ")
    paragraph.add_run("{{si").bold = True
    paragraph.add_run("te}} end").bold = True
    template.add_paragraph("{{正文}}")
    template_path = tmp_path / "template.docx"
    template.save(str(template_path))
    observation = _convert_docx(source, _link_config(), options={"template_name": str(template_path)})

    assert "Site: YAML jump end" in observation.text
    if matching:
        assert len(_anchors(observation.document_xml)) == 1
        assert set(_anchors(observation.document_xml)) == set(_navigation_bookmarks(observation.document_xml))
    else:
        assert not _anchors(observation.document_xml)
    assert not observation.hyperlink_targets
    assert "docwen-navigation:" not in observation.document_xml
    assert "DOCWENYAMLLINK" not in observation.document_xml


@pytest.mark.parametrize(
    ("source_link", "expected"),
    [("[**Markdown**](#missing)", "Markdown"), ("[[#Missing|*Wiki*]]", "*Wiki*")],
)
def test_unresolved_extract_text_uses_visible_text_not_generated_markdown(
    tmp_path: Path, source_link: str, expected: str
) -> None:
    source = tmp_path / "formatted-fallback.md"
    source.write_text(source_link, encoding="utf-8")
    config = _link_config()
    config["link"]["non_embed_links"].update(
        markdown_hyperlink_fallback="extract_text", wiki_hyperlink_fallback="extract_text"
    )
    observation = _convert_docx(source, config)
    assert observation.text == expected
    assert not _anchors(observation.document_xml)
    assert not observation.hyperlink_targets


def test_user_private_scheme_lookalike_is_never_a_navigation_capability(tmp_path: Path) -> None:
    source = tmp_path / "lookalike.md"
    source_link = "[Fake](docwen-navigation:guessed0)"
    source.write_text(source_link + "\n\n# guessed0\n", encoding="utf-8")
    observation = _convert_docx(source, _link_config())
    assert "Fake" in observation.text
    assert source_link not in observation.text
    assert not _anchors(observation.document_xml)
    assert not _navigation_bookmarks(observation.document_xml)
    assert not observation.hyperlink_targets


def test_embedded_document_cannot_supply_a_root_fragment_target(tmp_path: Path) -> None:
    from docwen_core.docx_semantics_v3 import ANCHOR_TAG_PREFIX

    child = tmp_path / "child.md"
    child_bytes = b"# Embedded\n\nChild body.\n"
    child.write_bytes(child_bytes)
    source = tmp_path / "root.md"
    source_bytes = b"[Child](#embedded) / [[#Embedded|Wiki child]] / [Root](#root)\n\n# Root ^root\n\nPlain root ^plain-root\n\n![[child.md]]\n"
    source.write_bytes(source_bytes)
    observation = _convert_docx(source, _link_config())

    assert "Child body." in observation.text
    assert "Child / Wiki child / Root" in observation.text
    assert not _anchors(observation.document_xml)
    assert not _navigation_bookmarks(observation.document_xml)
    assert ANCHOR_TAG_PREFIX in observation.document_xml
    assert source.read_bytes() == source_bytes
    assert child.read_bytes() == child_bytes


@pytest.mark.parametrize("label", ["Energy $mc^2$ result", "$mc^2$"])
def test_unresolved_heading_link_preserves_formula_in_text_fallback(tmp_path: Path, label: str) -> None:
    source = tmp_path / "formula-label.md"
    source.write_text(f"[{label}](#missing)", encoding="utf-8")
    config = _link_config()
    config["link"]["non_embed_links"]["markdown_hyperlink_fallback"] = "extract_text"
    observation = _convert_docx(source, config)
    assert observation.text == label
    assert not _anchors(observation.document_xml)
    assert not observation.hyperlink_targets
