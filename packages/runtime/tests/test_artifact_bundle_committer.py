"""Filesystem and graph safety contracts for ArtifactBundleCommitter."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from docwen_core.models import (
    BundleDraft,
    BundleDraftArtifact,
    BundleEntry,
    BundleRelation,
    PageFragmentSemantics,
)
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitError, ArtifactBundleCommitter

pytestmark = pytest.mark.contract


def _single_document(path: Path) -> BundleDraft:
    return BundleDraft(
        artifacts=(
            BundleDraftArtifact(
                artifact_id="artifact.primary",
                kind="document",
                path=str(path),
                suggested_name=path.name,
                media_type="text/markdown",
            ),
        ),
        entries=(BundleEntry("artifact.primary", "primary", 0, True),),
    )


def test_commit_pins_relative_locator_size_and_sha256(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    nested = staging / "documents"
    nested.mkdir(parents=True)
    output = nested / "result.md"
    output.write_bytes(b"# Result\n")

    bundle = ArtifactBundleCommitter().commit(
        task_id="task.commit",
        staging_root=str(staging),
        draft=_single_document(output),
    )

    artifact = bundle.artifacts[0]
    assert artifact.locator == "documents/result.md"
    assert artifact.size_bytes == len(b"# Result\n")
    assert artifact.sha256 == hashlib.sha256(b"# Result\n").hexdigest()
    assert bundle.to_dict()["schema"] == "docwen.artifact_bundle.v2"
    assert bundle.to_dict()["layout_schema"] == "docwen.artifact_layout.v1"
    assert bundle.to_dict()["artifacts"][0]["logical_path"] == "documents/result.md"


def test_commit_rejects_artifact_outside_staging(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    outside = tmp_path / "outside.md"
    outside.write_text("outside", encoding="utf-8")

    with pytest.raises(ArtifactBundleCommitError, match="outside the request-owned") as rejected:
        ArtifactBundleCommitter().commit(
            task_id="task.outside",
            staging_root=str(staging),
            draft=_single_document(outside),
        )
    assert rejected.value.code == "artifact_outside_staging"


def test_commit_rejects_dangling_graph_relation(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    output = staging / "result.md"
    output.write_text("result", encoding="utf-8")
    draft = _single_document(output)
    invalid = BundleDraft(
        artifacts=draft.artifacts,
        entries=draft.entries,
        relations=(
            BundleRelation(
                type="derived_from",
                source_artifact_id="artifact.primary",
                target_artifact_id="artifact.missing",
                role="source",
            ),
        ),
    )

    with pytest.raises(ArtifactBundleCommitError) as rejected:
        ArtifactBundleCommitter().commit(
            task_id="task.graph",
            staging_root=str(staging),
            draft=invalid,
        )
    assert rejected.value.code == "dangling_relation"


def test_commit_rejects_page_semantics_before_staging_root_or_artifact_io(tmp_path: Path) -> None:
    draft = BundleDraft(
        artifacts=(
            BundleDraftArtifact(
                "document.main", "document", str(tmp_path / "missing-main.md"), "main.md", "text/markdown"
            ),
            BundleDraftArtifact(
                "fragment.page.1", "fragment", str(tmp_path / "missing-page.md"), "page-001.md", "text/markdown"
            ),
        ),
        entries=(BundleEntry("document.main", "primary", 0, True),),
        relations=(
            BundleRelation(
                "fragment_of",
                "fragment.page.1",
                "document.main",
                "ocr_page",
                0,
                PageFragmentSemantics(0, 1, "success", 1),
            ),
        ),
    )

    with pytest.raises(ArtifactBundleCommitError) as rejected:
        ArtifactBundleCommitter().commit(
            task_id="task.page",
            staging_root=str(tmp_path / "missing-staging"),
            draft=draft,
        )

    assert rejected.value.code == "invalid_page_range"


def test_commit_rejects_duplicate_locator(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    output = staging / "result.md"
    output.write_text("result", encoding="utf-8")
    draft = BundleDraft(
        artifacts=(
            BundleDraftArtifact("artifact.one", "document", str(output), "one.md", "text/markdown"),
            BundleDraftArtifact("artifact.two", "document", str(output), "two.md", "text/markdown"),
        ),
        entries=(
            BundleEntry("artifact.one", "primary", 0, True),
            BundleEntry("artifact.two", "supplementary", 1, False),
        ),
    )

    with pytest.raises(ArtifactBundleCommitError) as rejected:
        ArtifactBundleCommitter().commit(
            task_id="task.duplicate",
            staging_root=str(staging),
            draft=draft,
        )
    assert rejected.value.code == "duplicate_artifact_locator"


def test_commit_rejects_symlink_artifact_when_supported(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    target = staging / "target.md"
    target.write_text("target", encoding="utf-8")
    link = staging / "link.md"
    try:
        link.symlink_to(target)
    except OSError:
        pytest.skip("symlink creation is unavailable on this host")

    with pytest.raises(ArtifactBundleCommitError) as rejected:
        ArtifactBundleCommitter().commit(
            task_id="task.link",
            staging_root=str(staging),
            draft=_single_document(link),
        )
    assert rejected.value.code == "artifact_path_is_link"


def test_discard_removes_only_named_staging_artifacts(tmp_path: Path) -> None:
    staging = tmp_path / "staging"
    nested = staging / "nested"
    nested.mkdir(parents=True)
    rejected = nested / "rejected.md"
    rejected.write_text("rejected", encoding="utf-8")
    retained = staging / "retained.md"
    retained.write_text("retained", encoding="utf-8")
    outside = tmp_path / "outside.md"
    outside.write_text("outside", encoding="utf-8")

    ArtifactBundleCommitter().discard(
        staging_root=str(staging),
        artifact_paths=[str(rejected), str(outside)],
    )

    assert not rejected.exists()
    assert not nested.exists()
    assert retained.read_text(encoding="utf-8") == "retained"
    assert outside.read_text(encoding="utf-8") == "outside"
