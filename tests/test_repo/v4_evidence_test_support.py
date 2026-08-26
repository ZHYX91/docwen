from __future__ import annotations

import hashlib
from collections.abc import Mapping
from pathlib import Path
from typing import cast

from scripts.release import v4_candidate_contract as contract
from scripts.release import v4_evidence_contract as evidence_contract


def record_ref(record: Mapping[str, object]) -> dict[str, object]:
    return {"caseId": record["caseId"], "layer": record["layer"], **evidence_contract.identity_core(record)}


def source_observation(root: Path, case_id: str, dimension: str) -> dict[str, object]:
    relative = f"contracts/oracles/docwen.markdown_semantics.v3/corpus/{case_id}.case.json"
    fixture = contract.read_json_object(root / relative, label=case_id)
    return {
        "schema": "docwen.v4_source_oracle_observation.v1",
        "fixture": contract.file_identity(root / relative, relative_to=root),
        "sourceSha256": hashlib.sha256(cast(str, fixture["source"]).encode()).hexdigest(),
        "expectedDiagnostics": evidence_contract._normalized_fixture_diagnostics(fixture),
        "invalidIdDimension": dimension,
    }


def source_bytes(root: Path, case_id: str) -> bytes:
    fixture = contract.read_json_object(
        root / f"contracts/oracles/docwen.markdown_semantics.v3/corpus/{case_id}.case.json", label=case_id
    )
    return cast(str, fixture["source"]).encode()


def wire_terminal(source: Mapping[str, object]) -> dict[str, object]:
    diagnostics = [
        {
            "severity": item["severity"],
            "code": item["code"],
            "message": "Synthetic closed wire diagnostic.",
            "evidence_schema": "docwen.machine.diagnostic_evidence.v1",
            "source": {
                "input_id": "invalid-id-dot.md",
                "sha256": source["sourceSha256"],
                "encoding": "utf-8",
                "coordinate_system": "unicode_code_point",
                "offset_base": 0,
                "range_end": "exclusive",
            },
            "range": item["range"],
            "related_ranges": item["relatedRanges"],
            "fixes": [],
        }
        for item in cast(list[dict[str, object]], source["expectedDiagnostics"])
    ]
    task_id = "task.synthetic.v4"
    return {
        "jsonrpc": "2.0",
        "method": "task/completed",
        "params": {
            "task_id": task_id,
            "sequence": 1,
            "bundle": {
                "schema": "docwen.artifact_bundle.v2",
                "layout_schema": "docwen.document_node.v1",
                "bundle_id": "bundle.synthetic.v4",
                "task_id": task_id,
                "producer": {
                    "name": "DocWen",
                    "product_version": "0.9.0",
                    "machine_protocol": "docwen.machine.v1",
                },
                "artifacts": [
                    {
                        "artifact_id": "document.synthetic",
                        "kind": "document",
                        "locator": "synthetic_20260820_120000_fromMarkdown/synthetic_20260820_120000_fromMarkdown.md",
                        "logical_path": "synthetic_20260820_120000_fromMarkdown/synthetic_20260820_120000_fromMarkdown.md",
                        "suggested_name": "synthetic.md",
                        "media_type": "text/markdown",
                        "size_bytes": 1,
                        "sha256": "a" * 64,
                    }
                ],
                "entries": [{"artifact_id": "document.synthetic", "role": "primary", "ordinal": 0, "preferred": True}],
                "relations": [],
            },
            "diagnostics": diagnostics,
            "metrics": {"duration_ms": 1, "input_bytes": 1, "output_bytes": 1},
        },
    }


def evidence_artifact(root: Path, name: str, content: bytes | object) -> dict[str, object]:
    path = root / "artifacts" / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content if isinstance(content, bytes) else contract.json_bytes(content))
    identity = contract.file_identity(path, relative_to=root)
    return {
        "relativePath": f"evidence/{identity['relativePath']}",
        "bytes": identity["bytes"],
        "sha256": identity["sha256"],
    }
