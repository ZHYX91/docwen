"""Public documentation guards for the DocWen-owned round-trip sidecar."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]


def test_public_docs_freeze_single_file_sidecar_and_safe_fallback() -> None:
    contracts = (ROOT / "contracts/README.md").read_text(encoding="utf-8")
    machine = (ROOT / "docs/specs/machine-protocol-v1.md").read_text(encoding="utf-8")
    markdown = (ROOT / "docs/specs/markdown-compatibility.md").read_text(encoding="utf-8")
    combined = "\n".join((contracts, machine, markdown))
    normalized_markdown = " ".join(markdown.split())

    assert "application/vnd.docwen.round-trip-sidecar+zip" in contracts
    assert "resource_of(role=manifest, ordinal=0)" in contracts
    assert "docwen.round_trip_sidecar.v1" in machine
    assert "authored-source.md" in machine
    assert "neutral-document.json" in machine
    assert "numbering-export-plan.json" in machine
    assert "manifest.json" in machine
    assert "Missing or mismatched sidecars disable byte-exact source recovery" in contracts
    assert "A Word edit disables byte-exact recovery" in normalized_markdown
    assert "consumers do not recreate it from private inputs" in machine
    assert "round-trip sidecar directory" not in combined.lower()
