"""Current Markdown compatibility documentation guards."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]


def test_markdown_compatibility_points_to_current_regression_entrypoints() -> None:
    text = (ROOT / "docs" / "specs" / "markdown-compatibility.md").read_text(encoding="utf-8")
    current_paths = (
        "tests/golden/test_md_to_docx_old_baseline.py",
        "packages/plugins/markdown/tests/test_md_to_docx_*.py",
        "packages/plugins/markdown/tests/test_md_to_docx_formatting.py",
        "packages/plugins/markdown/tests/test_md_to_spreadsheet_*.py",
        "packages/core/tests/test_links_markdown_orchestrator.py",
    )
    for relative_path in current_paths:
        assert relative_path in text
        assert read_source_text(ROOT / relative_path)


def test_markdown_compatibility_keeps_current_safety_rules() -> None:
    text = (ROOT / "docs" / "specs" / "markdown-compatibility.md").read_text(encoding="utf-8")

    assert "must not fetch remote content implicitly" in text
    assert "must not disappear silently" in text
    assert "one precedence chain" in text


def test_markdown_compatibility_freezes_ref_occurrence_recovery() -> None:
    text = (ROOT / "docs" / "specs" / "markdown-compatibility.md").read_text(encoding="utf-8")

    assert "docwen-ref-occurrence-v1:<digest32>" in text
    assert "https://docwen.dev/schema/document-reference-occurrence-map/v1" in text
    assert (
        "docwen-ref-occurrence-map-v1\\0<source-sha256>\\0<start>\\0<end>\\0<authored-token>"
        "\\0<resolved-bookmark-name>\\0<cached-number>"
    ) in text
    assert "number plus Alias" in text
    assert "This occurrence wrapper is not a target, bookmark, second ID" in text


def test_markdown_compatibility_freezes_caption_order_and_fenced_anchor_kind() -> None:
    text = (ROOT / "docs" / "specs" / "markdown-compatibility.md").read_text(encoding="utf-8")
    normalized = " ".join(text.split())

    assert "A Figure logical object is immediately followed by its `DocWenFigureCaption` paragraph" in normalized
    assert "An ID-less caption is recovered only from that fixed direct adjacency" in normalized
    assert "It never guesses a pair or invents an ID" in normalized
    assert "Every ordinary anchor on a CommonMark fenced block uses `code_block`" in normalized
    assert "`fenced_block` is never serialized in this DOCX map" in normalized


def test_docx_recovery_resource_less_image_owner_contract_is_frozen() -> None:
    markdown = (ROOT / "docs" / "specs" / "markdown-compatibility.md").read_text(encoding="utf-8")
    machine = (ROOT / "docs" / "specs" / "machine-protocol-v1.md").read_text(encoding="utf-8")
    golden = (ROOT / "docs" / "specs" / "golden-regression-suite.md").read_text(encoding="utf-8")
    normalized_markdown = " ".join(markdown.split())
    normalized_machine = " ".join(machine.split())

    for text in (markdown, machine, golden):
        assert "![image omitted]()" in text
        assert "DOCX2MD-IMAGE-OWNER-RESOURCE-OMITTED" in text
    assert "no other empty-destination image is this carrier" in normalized_markdown
    assert "sidecar contains OCR presentation only" in normalized_markdown
    assert "primary owner is never replaced by an OCR-sidecar embed" in normalized_machine
    assert "With the default `image_mode=file`" in machine
    assert "if it cannot retain the authenticated image owner the task fails closed" in normalized_machine
    assert "does not set `evidence_schema`, `source`, `range`, `related_ranges`, or `fixes`" in normalized_machine


def test_fenced_source_occurrence_package_and_evidence_layers_are_frozen() -> None:
    markdown = (ROOT / "docs" / "specs" / "markdown-compatibility.md").read_text(encoding="utf-8")
    machine = (ROOT / "docs" / "specs" / "machine-protocol-v1.md").read_text(encoding="utf-8")
    golden = (ROOT / "docs" / "specs" / "golden-regression-suite.md").read_text(encoding="utf-8")
    normalized = " ".join(markdown.split())

    assert "https://docwen.dev/schema/document-fenced-source-map/v1" in markdown
    assert 'documentFencedSourceMap version="1"' in markdown
    assert "docwen-fenced-source-v1:" in markdown
    assert "fenced_sources" in markdown
    assert (
        "docwen-fenced-source-map-v1\\0<source_sha256>\\0<source_start>\\0<source_end>\\0<block_sha256>\\0<body_sha256>"
    ) in normalized
    assert "create zero bookmark, `SEQ`, or `REF` facts" in normalized
    assert "Direct/top-level, blockquote, and list-container" in normalized
    assert "source_oracle" in machine and "packaged" in machine and "roundtrip" in machine
    assert "neither is `task/completed.params.diagnostics` wire evidence" in machine
    assert "rust/Mermaid/query/view" in golden
    assert "partial payload wrapping" in golden


def test_nested_ordinary_anchor_topology_source_package_and_roundtrip_are_frozen() -> None:
    markdown = (ROOT / "docs" / "specs" / "markdown-compatibility.md").read_text(encoding="utf-8")
    machine = (ROOT / "docs" / "specs" / "machine-protocol-v1.md").read_text(encoding="utf-8")
    golden = (ROOT / "docs" / "specs" / "golden-regression-suite.md").read_text(encoding="utf-8")
    normalized = " ".join(markdown.split())

    assert "container_path" in markdown
    assert "longest proper prefix" in normalized
    assert "https://docwen.dev/schema/document-anchor-topology-map/v1" in markdown
    assert 'documentAnchorTopologyMap version="1"' in markdown
    assert "child_tag,parent_tag,sha256" in markdown
    assert "docwen-anchor-topology-edge-v1\\0<child_tag>\\0<parent_tag>" in normalized
    assert "inline multi-paragraph `list_item` anchor" in normalized
    assert "complete structural-container range" in normalized
    assert "28a486c7939e34bd8d6654ec694c0a7fdbf3f1af2aceb37d76db22d6b01124de" in markdown
    assert "acyclic forest" in normalized
    assert "proper contiguous subset" in normalized and "same ordered range" in normalized
    assert "outer quote ordinary SDT" in normalized and "inline fenced-source carrier SDT" in normalized
    for anchor_id in ("inner-fence", "outer-quote", "inner-list", "inner-quote", "outer-list"):
        assert anchor_id in markdown
        assert anchor_id in golden
    assert "emit no topology map" in golden
    assert "swapped tags" in golden and "reversed wrappers" in golden
    assert "Nested ordinary-anchor topology is layered" in machine
    assert "None of these is a Machine option, Bundle relation, consumer hierarchy" in machine
