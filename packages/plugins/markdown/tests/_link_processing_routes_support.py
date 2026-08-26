"""Route-level contracts for request-scoped Markdown link processing."""

from __future__ import annotations

import base64
import csv
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest
from docx import Document
from openpyxl import Workbook, load_workbook

pytestmark = pytest.mark.contract

from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_spreadsheet.converter import (
    MdToCsvConverter,
    MdToXlsxConverter,
)

from .conftest import make_context

_TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)

_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


@dataclass(frozen=True)
class _DocxObservation:
    text: str
    document_xml: str
    hyperlink_targets: frozenset[str]
    media_names: tuple[str, ...]


def _link_config(
    *,
    wiki_mode: str = "hyperlink",
    markdown_mode: str = "hyperlink",
    auto_link_bare_url: bool = False,
    wiki_image_mode: str = "embed",
    markdown_image_mode: str = "embed",
    md_file_mode: str = "embed",
    max_depth: int = 3,
    file_not_found_mode: str = "placeholder",
) -> dict[str, Any]:
    return {
        "link": {
            "non_embed_links": {
                "wiki_mode": wiki_mode,
                "markdown_mode": markdown_mode,
                "auto_link_bare_url": auto_link_bare_url,
            },
            "embed_links": {
                "wiki_image_mode": wiki_image_mode,
                "markdown_image_mode": markdown_image_mode,
                "md_file_mode": md_file_mode,
            },
            "embedding": {"max_depth": max_depth},
            "path_resolution": {
                "search_dirs": [".", "assets", "images", "attachments"],
            },
            "error_handling": {
                "file_not_found": file_not_found_mode,
                "detect_circular": True,
                "circular_reference": "placeholder",
                "max_depth_reached": "placeholder",
            },
        }
    }


def _convert_docx(
    source: Path,
    config_values: dict[str, Any],
    *,
    options: dict[str, Any] | None = None,
) -> _DocxObservation:
    context, _workspace = make_context(
        str(source),
        target_format="docx",
        options=options,
        config_values=config_values,
    )
    result = MdToDocxConverter().convert(context)
    assert result.success is True, result.error

    output = Path(result.artifacts[0].staging_path)
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
        rels_root = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        media_names = tuple(sorted(name for name in archive.namelist() if name.startswith("word/media/")))

    document_root = ET.fromstring(document_xml)
    text = _visible_docx_text(document_root)
    hyperlink_targets = frozenset(
        relationship.attrib["Target"]
        for relationship in rels_root.iter(f"{{{_REL_NS}}}Relationship")
        if relationship.attrib.get("Type", "").endswith("/hyperlink")
    )
    return _DocxObservation(
        text=text,
        document_xml=document_xml,
        hyperlink_targets=hyperlink_targets,
        media_names=media_names,
    )


def _convert_declared_docx(
    source: Path,
    *,
    source_logical_path: str,
    resource: Path,
    resource_logical_path: str,
) -> _DocxObservation:
    from docwen_core.models.file_ref import FileRef

    context, workspace = make_context(
        str(source),
        target_format="docx",
        config_values=_link_config(markdown_image_mode="embed"),
    )
    source_ref = FileRef(
        path=str(source),
        format="markdown",
        category="document",
        input_kind="document",
        input_role="source",
        logical_path=source_logical_path,
    )
    resource_ref = FileRef(
        path=str(resource),
        format="png",
        category="image",
        input_kind="resource",
        input_role="linked_resource",
        logical_path=resource_logical_path,
    )
    context.request.input_refs = [source_ref, resource_ref]
    workspace._input_refs = (source_ref, resource_ref)
    result = MdToDocxConverter().convert(context)
    assert result.success is True, result.error
    output = Path(result.artifacts[0].staging_path)
    with zipfile.ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
        rels_root = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        media_names = tuple(sorted(name for name in archive.namelist() if name.startswith("word/media/")))
    document_root = ET.fromstring(document_xml)
    text = _visible_docx_text(document_root)
    return _DocxObservation(
        text=text,
        document_xml=document_xml,
        hyperlink_targets=frozenset(
            relationship.attrib["Target"]
            for relationship in rels_root.iter(f"{{{_REL_NS}}}Relationship")
            if relationship.attrib.get("Type", "").endswith("/hyperlink")
        ),
        media_names=media_names,
    )


def _visible_docx_text(document_root: ET.Element) -> str:
    """Project Word text while preserving explicit line-break separation."""

    text_tag = f"{{{_WORD_NS}}}t"
    break_tag = f"{{{_WORD_NS}}}br"
    return "".join(
        " " if node.tag == break_tag else node.text or ""
        for node in document_root.iter()
        if node.tag in {text_tag, break_tag}
    )


def _assert_target(observation: _DocxObservation, expected: str) -> None:
    normalized_expected = expected.replace("\\", "/")
    assert any(
        target.replace("\\", "/") == normalized_expected
        or target.replace("\\", "/").endswith(f"/{normalized_expected}")
        for target in observation.hyperlink_targets
    ), observation.hyperlink_targets


__all__ = (
    "ET",
    "_TINY_PNG",
    "_WP_NS",
    "Any",
    "Document",
    "MdToCsvConverter",
    "MdToDocxConverter",
    "MdToXlsxConverter",
    "Path",
    "Workbook",
    "_assert_target",
    "_convert_declared_docx",
    "_convert_docx",
    "_link_config",
    "csv",
    "load_workbook",
    "make_context",
    "pytest",
    "pytestmark",
)
