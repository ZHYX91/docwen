"""Runtime-owned validation and integrity commit for Artifact Bundle v2."""

from __future__ import annotations

import hashlib
import os
import sys
from pathlib import Path
from uuid import uuid4

from docwen_core.models.artifact_bundle import (
    ARTIFACT_BUNDLE_SCHEMA,
    ArtifactBundle,
    ArtifactBundleValidationError,
    BundleArtifact,
    BundleDraft,
    BundleDraftArtifact,
    BundleProducer,
    validate_artifact_bundle,
    validate_artifact_bundle_draft,
)
from docwen_core.version import __version__
from docwen_runtime.path_io import filesystem_path

_HASH_CHUNK_BYTES = 1024 * 1024


class ArtifactBundleCommitError(ValueError):
    """A draft could not be safely committed as a public bundle."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(f"{code}: {message}")
        self.code = code


class ArtifactBundleCommitter:
    """Turn a semantic draft inside request-owned staging into an integrity-pinned bundle."""

    def commit(self, *, task_id: str, staging_root: str, draft: BundleDraft) -> ArtifactBundle:
        try:
            validate_artifact_bundle_draft(draft)
        except ArtifactBundleValidationError as exc:
            raise ArtifactBundleCommitError(exc.code, str(exc)) from exc
        root = self._validated_root(staging_root)
        artifacts = tuple(self._commit_artifact(root, artifact) for artifact in draft.artifacts)
        bundle = ArtifactBundle(
            bundle_id=f"bundle.{uuid4().hex}",
            task_id=task_id,
            producer=BundleProducer(product_version=__version__),
            artifacts=artifacts,
            entries=draft.entries,
            relations=draft.relations,
            schema=ARTIFACT_BUNDLE_SCHEMA,
            layout_schema=(
                "docwen.document_node.v1"
                if any(artifact.media_type == "text/markdown" for artifact in draft.artifacts)
                and all(
                    artifact.logical_path is not None
                    for artifact in draft.artifacts
                    if artifact.media_type == "text/markdown"
                )
                else "docwen.artifact_layout.v1"
            ),
        )
        try:
            validate_artifact_bundle(bundle)
        except ArtifactBundleValidationError as exc:
            raise ArtifactBundleCommitError(exc.code, str(exc)) from exc
        return bundle

    def discard(self, *, staging_root: str, artifact_paths: list[str]) -> None:
        """Remove exact rejected paths without following links outside staging."""

        root = self._validated_root(staging_root)
        parents: set[Path] = set()
        for raw_path in artifact_paths:
            candidate = filesystem_path(raw_path, force_extended=sys.platform == "win32")
            if not candidate.is_absolute():
                continue
            try:
                lexical_relative = candidate.absolute().relative_to(root)
            except ValueError:
                continue
            current = root / lexical_relative
            if self._is_link_or_junction(current):
                self._remove_link(current)
                parents.add(current.parent)
                continue
            try:
                resolved = current.resolve(strict=True)
                resolved.relative_to(root)
            except (OSError, ValueError):
                continue
            if resolved.is_file():
                resolved.unlink()
                parents.add(resolved.parent)

        for parent in sorted(parents, key=lambda item: len(item.parts), reverse=True):
            current = parent
            while current != root:
                try:
                    current.rmdir()
                except OSError:
                    break
                current = current.parent

    @classmethod
    def _commit_artifact(cls, root: Path, draft_artifact: BundleDraftArtifact) -> BundleArtifact:
        raw_path = draft_artifact.path
        artifact_path = filesystem_path(raw_path, force_extended=sys.platform == "win32")
        if not artifact_path.is_absolute():
            raise ArtifactBundleCommitError("artifact_path_not_absolute", "draft artifact paths must be absolute")
        if cls._is_link_or_junction(artifact_path):
            raise ArtifactBundleCommitError("artifact_path_is_link", f"artifact is a link or junction: {raw_path!r}")
        try:
            resolved = artifact_path.resolve(strict=True)
        except OSError as exc:
            raise ArtifactBundleCommitError("artifact_missing", f"artifact does not exist: {raw_path!r}") from exc
        try:
            relative = resolved.relative_to(root)
        except ValueError as exc:
            raise ArtifactBundleCommitError(
                "artifact_outside_staging",
                f"artifact is outside the request-owned staging root: {raw_path!r}",
            ) from exc
        cls._reject_linked_descendant(root, relative)
        if not resolved.is_file():
            raise ArtifactBundleCommitError(
                "artifact_not_regular_file", f"artifact is not a regular file: {raw_path!r}"
            )

        locator = relative.as_posix()
        if not locator or locator == ".":
            raise ArtifactBundleCommitError("invalid_artifact_locator", "artifact locator must name a file")
        suggested_name = draft_artifact.suggested_name
        if suggested_name in {"", ".", ".."} or Path(suggested_name).name != suggested_name:
            raise ArtifactBundleCommitError(
                "invalid_suggested_name",
                f"suggested_name must be a safe basename: {suggested_name!r}",
            )

        size_bytes, sha256 = cls._integrity(resolved)
        return BundleArtifact(
            artifact_id=draft_artifact.artifact_id,
            kind=draft_artifact.kind,
            locator=locator,
            suggested_name=suggested_name,
            media_type=draft_artifact.media_type,
            size_bytes=size_bytes,
            sha256=sha256,
            logical_path=draft_artifact.logical_path or locator,
        )

    @classmethod
    def _validated_root(cls, staging_root: str) -> Path:
        raw_root = filesystem_path(staging_root, force_extended=sys.platform == "win32")
        if not raw_root.is_absolute():
            raise ArtifactBundleCommitError("staging_root_not_absolute", "staging_root must be absolute")
        if cls._is_link_or_junction(raw_root):
            raise ArtifactBundleCommitError("staging_root_is_link", "staging_root must not be a link or junction")
        try:
            root = raw_root.resolve(strict=True)
        except OSError as exc:
            raise ArtifactBundleCommitError("staging_root_missing", "staging_root must already exist") from exc
        if not root.is_dir():
            raise ArtifactBundleCommitError("staging_root_not_directory", "staging_root must be a directory")
        return root

    @classmethod
    def _reject_linked_descendant(cls, root: Path, relative: Path) -> None:
        current = root
        for part in relative.parts:
            current /= part
            if cls._is_link_or_junction(current):
                raise ArtifactBundleCommitError(
                    "artifact_path_is_link",
                    f"artifact path traverses a link or junction: {os.fspath(current)!r}",
                )

    @staticmethod
    def _is_link_or_junction(path: Path) -> bool:
        if path.is_symlink():
            return True
        is_junction = getattr(path, "is_junction", None)
        return bool(is_junction()) if callable(is_junction) else False

    @staticmethod
    def _remove_link(path: Path) -> None:
        is_junction = getattr(path, "is_junction", None)
        if callable(is_junction) and is_junction():
            path.rmdir()
        else:
            path.unlink()

    @staticmethod
    def _integrity(path: Path) -> tuple[int, str]:
        digest = hashlib.sha256()
        size_bytes = 0
        try:
            with path.open("rb") as stream:
                while chunk := stream.read(_HASH_CHUNK_BYTES):
                    size_bytes += len(chunk)
                    digest.update(chunk)
        except OSError as exc:
            raise ArtifactBundleCommitError(
                "artifact_read_failed", f"cannot read artifact: {os.fspath(path)!r}"
            ) from exc
        return size_bytes, digest.hexdigest()


__all__ = ["ArtifactBundleCommitError", "ArtifactBundleCommitter"]
