"""Full-input YAML and Machine diagnostic evidence for Markdown -> DOCX v4."""

from __future__ import annotations

import hashlib
from pathlib import Path
from zipfile import ZipFile

import pytest
from tests.support.numbering import repository_numbering_registry

from docwen_core.docx_semantics_v3 import (
    REFERENCE_OCCURRENCE_MAP_NAMESPACE,
    TARGET_MAP_NAMESPACE,
)
from docwen_core.models.file_ref import FileRef
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from .conftest import make_context

pytestmark = pytest.mark.contract


def _bind_machine_source(workspace: object, source: Path, *, input_id: str) -> None:
    """Mirror the admitted Machine source metadata used by ConversionService."""

    source_ref = FileRef(
        path=str(source),
        format="markdown",
        category="document",
        size_bytes=source.stat().st_size,
        input_kind="document",
        input_role="source",
        logical_path="source.md",
        media_type="text/markdown",
        metadata={
            "machine_input_id": input_id,
            "machine_input_size_bytes": source.stat().st_size,
            "machine_input_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        },
    )
    workspace._input_refs = (source_ref,)  # type: ignore[attr-defined]


def test_closed_yaml_is_masked_and_invalid_body_id_has_full_input_wire_evidence(tmp_path: Path) -> None:
    source_text = '\ufeff---\ntitle: "@yaml-citation [[Page#Heading]] ^yaml_bad"\n---\n\n前缀😀正文 ^bad_id\n'
    source = tmp_path / "yaml-invalid.md"
    source.write_bytes(source_text.encode("utf-8"))
    context, workspace = make_context(str(source), target_format="docx")
    _bind_machine_source(workspace, source, input_id="machine-input-yaml")

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.artifacts == []
    assert workspace.registered_artifacts == []
    [diagnostic] = result.diagnostics
    payload = diagnostic.to_dict()
    token_start = source_text.rindex("^bad_id")
    assert payload["code"] == "docwen.markdown.anchor.invalid_id"
    assert payload["evidence_schema"] == "docwen.machine.diagnostic_evidence.v1"
    assert payload["source"] == {
        "input_id": "machine-input-yaml",
        "sha256": hashlib.sha256(source_text.encode("utf-8")).hexdigest(),
        "encoding": "utf-8",
        "coordinate_system": "unicode_code_point",
        "offset_base": 0,
        "range_end": "exclusive",
    }
    assert payload["range"] == {
        "start": token_start,
        "end": token_start + len("^bad_id"),
    }
    assert source_text[payload["range"]["start"] : payload["range"]["end"]] == "^bad_id"
    assert all("yaml_bad" not in str(item) for item in result.diagnostics)
    assert all("yaml-citation" not in str(item) for item in result.diagnostics)


def test_yaml_body_marker_survives_text_numbering_without_reusing_old_body_offset(tmp_path: Path) -> None:
    source_text = (
        "---\n"
        'title: "@yaml-citation [[Page#Heading]] ^yaml_bad"\n'
        "---\n"
        "\n"
        "# 正文😀 ^stable-heading\n"
        "\n"
        "See @[[#^stable-heading]].\n"
    )
    source = tmp_path / "yaml-numbered.md"
    source.write_bytes(source_text.encode("utf-8"))
    context, workspace = make_context(
        str(source),
        target_format="docx",
        options={
            "add_numbering": True,
            "numbering_scheme": "hierarchical_standard",
            "heading_numbering_render_mode": "text",
        },
        numbering_registry=repository_numbering_registry(),
    )
    _bind_machine_source(workspace, source, input_id="machine-input-numbered")

    result = MdToDocxConverter().convert(context)

    assert result.success, result.error
    assert [item.code for item in result.diagnostics if item.level == "error"] == []
    package_path = Path(result.artifacts[0].staging_path)
    with ZipFile(package_path) as package:
        semantic_parts = [
            package.read(name) for name in package.namelist() if name.startswith("customXml/") and name.endswith(".xml")
        ]
    target_map = next(part for part in semantic_parts if TARGET_MAP_NAMESPACE.encode("utf-8") in part)
    reference_map = next(part for part in semantic_parts if REFERENCE_OCCURRENCE_MAP_NAMESPACE.encode("utf-8") in part)
    assert b"stable-heading" in target_map
    assert hashlib.sha256(source_text.encode("utf-8")).hexdigest().encode("ascii") in reference_map
    assert b"yaml_bad" not in target_map
