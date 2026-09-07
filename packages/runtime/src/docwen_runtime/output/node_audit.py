"""Prepare an optional audit document inside the conversion's unpublished root."""

from __future__ import annotations

import hashlib
import os
from dataclasses import replace
from pathlib import Path, PurePosixPath

from docwen_core.models.artifact import ArtifactManifest
from docwen_runtime.output.document_node import DocumentNodeLayoutPlan
from docwen_runtime.output.manifest import OutputManifestDocument, canonical_manifest_bytes


def stage_node_audit(plan: DocumentNodeLayoutPlan, root: str, document: OutputManifestDocument) -> ArtifactManifest:
    reserved = {PurePosixPath(item.logical_path or "").name.casefold() for item in plan.artifacts}
    name = "manifest.json"
    counter = 1
    while name.casefold() in reserved:
        name = f"manifest_{counter}.json"
        counter += 1
    document = replace(
        document,
        artifacts=tuple(
            (item.kind, item.suggested_name, item.media_type, item.is_primary)
            for item in plan.artifacts
            if item.kind != "manifest"
        ),
    )
    path = Path(root) / name
    with path.open("xb") as stream:
        stream.write(canonical_manifest_bytes(document))
        stream.flush()
        os.fsync(stream.fileno())
    logical = f"{plan.root_name}/{name}"
    return ArtifactManifest(
        artifact_id=f"audit-{hashlib.sha256(plan.identity.task_id.encode()).hexdigest()[:16]}",
        kind="manifest",
        staging_path=str(path),
        suggested_name=name,
        media_type="application/json",
        logical_path=logical,
        metadata={"document_node_role": "audit", "node_root": plan.root_name, "logical_path": logical},
    )
