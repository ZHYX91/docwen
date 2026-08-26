"""Permanent oracle tests for the packaged physical-page Bundle verifier."""

from __future__ import annotations

import copy
import hashlib
from pathlib import Path
from typing import Any

import pytest
from scripts.release.verify_packaged_cli import _OFD_FIXTURE_SCRIPT, _verify_physical_page_bundle

pytestmark = pytest.mark.contract


def test_ofd_fixture_disables_dependency_logging_before_import() -> None:
    assert _OFD_FIXTURE_SCRIPT.index('loguru_logger.disable("easyofd")') < _OFD_FIXTURE_SCRIPT.index(
        "from easyofd import OFD"
    )


def _artifact(root: Path, artifact_id: str, kind: str, payload: bytes) -> dict[str, Any]:
    path = root / f"{artifact_id}.bin"
    path.write_bytes(payload)
    return {
        "artifact_id": artifact_id,
        "kind": kind,
        "media_type": "text/markdown" if kind != "resource" else "image/png",
        "locator": path.name,
        "logical_path": path.name,
        "suggested_name": path.name,
        "size_bytes": len(payload),
        "sha256": hashlib.sha256(payload).hexdigest(),
    }


def _canonical_terminal(root: Path) -> dict[str, Any]:
    artifacts = [_artifact(root, "document.main", "document", b"primary only\n")]
    statuses = ("success", "no_text", "recognition_failed", "success")
    fragment_payloads = (b"page one\n", b"", b"", b"page four\n")
    artifacts.extend(
        _artifact(root, f"fragment.{page}", "fragment", fragment_payloads[page - 1]) for page in range(1, 5)
    )
    artifacts.extend(_artifact(root, f"resource.{page}", "resource", bytes([page])) for page in range(1, 6))
    relations = [
        {
            "type": "fragment_of",
            "source_artifact_id": f"fragment.{page}",
            "target_artifact_id": "document.main",
            "role": "ocr_page",
            "ordinal": page - 1,
            "page_fragment": {
                "fragment_kind": "page",
                "page_index": page,
                "page_count": 4,
                "ocr_status": statuses[page - 1],
                "source_page": page,
            },
        }
        for page in range(1, 5)
    ]
    relations.extend(
        {
            "type": "resource_of",
            "source_artifact_id": f"resource.{page}",
            "target_artifact_id": f"fragment.{page}",
            "role": "image",
            "ordinal": page - 1,
            "page_resource": {"source_page": page},
        }
        for page in range(1, 5)
    )
    relations.append(
        {
            "type": "resource_of",
            "source_artifact_id": "resource.5",
            "target_artifact_id": "document.main",
            "role": "image",
            "ordinal": 4,
        }
    )
    diagnostics = [
        {
            "severity": "warning",
            "code": "OCR-BEST-EFFORT",
            "message": "best effort",
            "artifact_id": f"fragment.{page}",
        }
        for page in range(1, 5)
    ]
    diagnostics.append(
        {
            "severity": "warning",
            "code": "resource_page_unresolved",
            "message": "page unresolved",
            "artifact_id": "resource.5",
        }
    )
    return {
        "method": "task/completed",
        "params": {
            "bundle": {
                "schema": "docwen.artifact_bundle.v2",
                "layout_schema": "docwen.artifact_layout.v1",
                "task_id": "task.physical",
                "producer": {"name": "DocWen", "version": "0.9.0"},
                "artifacts": artifacts,
                "entries": [{"artifact_id": "document.main", "role": "primary", "ordinal": 0, "preferred": True}],
                "relations": relations,
            },
            "diagnostics": diagnostics,
        },
    }


def _verify(terminal: dict[str, Any], root: Path) -> None:
    _verify_physical_page_bundle(
        terminal=terminal,
        task_id="task.physical",
        staging_root=root,
        page_count=4,
        resource_count=5,
        ocr_enabled=True,
        keep_images=True,
        expected_statuses=("success", "no_text", "recognition_failed", "success"),
    )


def test_packaged_physical_page_verifier_accepts_canonical_p4_k5(tmp_path: Path) -> None:
    _verify(_canonical_terminal(tmp_path), tmp_path)


@pytest.mark.parametrize(
    "damage",
    [
        "extra_relation",
        "wrong_resource_source",
        "nonempty_failure",
        "duplicate_primary",
        "missing_diagnostic",
        "dangling_diagnostic",
        "unsafe_locator",
        "unknown_page_field",
    ],
)
def test_packaged_physical_page_verifier_rejects_semantic_damage(tmp_path: Path, damage: str) -> None:
    terminal = _canonical_terminal(tmp_path)
    damaged = copy.deepcopy(terminal)
    bundle = damaged["params"]["bundle"]
    if damage == "extra_relation":
        bundle["relations"].append(copy.deepcopy(bundle["relations"][0]))
    elif damage == "wrong_resource_source":
        relation = next(item for item in bundle["relations"] if item["source_artifact_id"] == "resource.1")
        relation["source_artifact_id"] = "resource.2"
    elif damage == "nonempty_failure":
        artifact = next(item for item in bundle["artifacts"] if item["artifact_id"] == "fragment.3")
        path = tmp_path / artifact["locator"]
        path.write_bytes(b"must stay empty")
        artifact["size_bytes"] = path.stat().st_size
        artifact["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
    elif damage == "duplicate_primary":
        primary = next(item for item in bundle["artifacts"] if item["artifact_id"] == "document.main")
        path = tmp_path / primary["locator"]
        path.write_bytes(path.read_bytes() + b"page one\n")
        primary["size_bytes"] = path.stat().st_size
        primary["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
    elif damage == "missing_diagnostic":
        damaged["params"]["diagnostics"] = [
            item for item in damaged["params"]["diagnostics"] if item.get("artifact_id") != "resource.5"
        ]
    elif damage == "dangling_diagnostic":
        damaged["params"]["diagnostics"].append(
            {
                "severity": "warning",
                "code": "other",
                "message": "dangling",
                "artifact_id": "artifact.missing",
            }
        )
    elif damage == "unsafe_locator":
        next(item for item in bundle["artifacts"] if item["artifact_id"] == "resource.1")["locator"] = "C:/escape.png"
    else:
        next(item for item in bundle["relations"] if item["source_artifact_id"] == "fragment.1")["page_fragment"][
            "unknown"
        ] = True

    with pytest.raises(RuntimeError):
        _verify(damaged, tmp_path)
