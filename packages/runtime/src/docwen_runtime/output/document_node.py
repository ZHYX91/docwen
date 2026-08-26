"""Central Markdown document-node layout planning and link relocation."""

from __future__ import annotations

import os
import posixpath
import re
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path, PurePosixPath
from urllib.parse import unquote

from docwen_core.models import (
    DOCUMENT_NODE_SCHEMA,
    ArtifactManifest,
    ConversionIdentity,
    DocumentNodePath,
    sanitize_node_label,
    validate_logical_path,
    validate_markdown_node_path,
)

MARKDOWN_MEDIA_TYPE = "text/markdown"
DOCUMENT_NODE_MANIFEST_MEDIA_TYPE = "application/vnd.docwen.document-node+json"

_MARKDOWN_LINK = re.compile(r"(?P<prefix>!?\[[^\]\n]*\]\()(?P<target>[^)\s]+)(?P<suffix>[^)]*\))")
_WIKI_LINK = re.compile(r"(?P<prefix>!?\[\[)(?P<body>[^\]\n]+)(?P<suffix>\]\])")
_HTML_LINK = re.compile(
    r"(?P<prefix>\b(?:src|href)\s*=\s*[\"'])(?P<target>[^\"']+)(?P<suffix>[\"'])",
    re.IGNORECASE,
)


@dataclass(frozen=True, slots=True)
class DocumentNodeLayoutPlan:
    identity: ConversionIdentity
    root_name: str
    artifacts: tuple[ArtifactManifest, ...]

    @property
    def root_path(self) -> PurePosixPath:
        return PurePosixPath(self.root_name)

    def rebase_root(self, root_name: str) -> DocumentNodeLayoutPlan:
        """Return the same semantic plan under a collision-resolved root."""

        rebased: list[ArtifactManifest] = []
        for artifact in self.artifacts:
            if artifact.logical_path is None:
                raise ValueError("planned artifact is missing logical_path")
            old = validate_logical_path(artifact.logical_path)
            relative = old.relative_to(self.root_path)
            if artifact.metadata.get("document_node_role") == "primary":
                relative = PurePosixPath(f"{root_name}.md")
            logical = (PurePosixPath(root_name) / relative).as_posix()
            if artifact.media_type == MARKDOWN_MEDIA_TYPE:
                validate_markdown_node_path(logical)
            rebased.append(
                replace(
                    artifact,
                    suggested_name=PurePosixPath(logical).name,
                    logical_path=logical,
                    metadata={**artifact.metadata, "node_root": root_name, "logical_path": logical},
                )
            )
        return DocumentNodeLayoutPlan(identity=self.identity, root_name=root_name, artifacts=tuple(rebased))


def has_markdown_artifacts(artifacts: list[ArtifactManifest] | tuple[ArtifactManifest, ...]) -> bool:
    return any(artifact.media_type == MARKDOWN_MEDIA_TYPE for artifact in artifacts)


def plan_document_node_layout(
    *,
    task_id: str,
    artifacts: list[ArtifactManifest] | tuple[ArtifactManifest, ...],
    input_path: str,
    created_at: datetime | None = None,
    root_collision: int = 0,
) -> DocumentNodeLayoutPlan:
    """Assign one root node and one ``X/X.md`` path to every Markdown."""

    markdown = [artifact for artifact in artifacts if artifact.media_type == MARKDOWN_MEDIA_TYPE]
    if not markdown:
        raise ValueError("document-node layout requires at least one Markdown artifact")
    preferred = [artifact for artifact in markdown if artifact.is_primary]
    primary = preferred[0] if preferred else markdown[0]
    source = Path(input_path)
    source_stem = source.stem or Path(primary.suggested_name).stem or "document"
    source_format = source.suffix.lstrip(".") or str(primary.metadata.get("source_format", "unknown"))
    identity = ConversionIdentity.create(
        task_id=task_id,
        source_stem=source_stem,
        source_format=source_format,
        created_at=created_at,
    )
    root_name = identity.node_name(collision=root_collision)
    root = DocumentNodePath(root_name)

    planned: list[ArtifactManifest] = []
    used_paths: set[str] = set()
    child_labels: dict[str, int] = {}
    resource_names: dict[str, int] = {}
    for artifact in artifacts:
        metadata = {
            **artifact.metadata,
            "document_node_schema": DOCUMENT_NODE_SCHEMA,
            "node_root": root_name,
            "source_suggested_name": artifact.suggested_name,
        }
        if artifact is primary:
            logical = root.markdown.as_posix()
            metadata["document_node_role"] = "primary"
        elif artifact.media_type == MARKDOWN_MEDIA_TYPE:
            label = _markdown_child_label(identity, artifact)
            key = label.casefold()
            child_labels[key] = child_labels.get(key, 0) + 1
            if child_labels[key] > 1:
                label = f"{label}_{child_labels[key]:02d}"
            child_name = identity.node_name(label)
            logical = root.child(child_name).markdown.as_posix()
            metadata["document_node_role"] = _markdown_role(artifact)
        else:
            resource_name = _portable_resource_name(artifact.suggested_name or Path(artifact.staging_path).name)
            key = resource_name.casefold()
            resource_names[key] = resource_names.get(key, 0) + 1
            if resource_names[key] > 1:
                stem, suffix = os.path.splitext(resource_name)
                resource_name = f"{stem}_{resource_names[key]:03d}{suffix}"
            logical = (root.directory / resource_name).as_posix()
            metadata["document_node_role"] = "resource"
        if logical.casefold() in used_paths:
            raise ValueError(f"duplicate logical artifact path: {logical}")
        used_paths.add(logical.casefold())
        if artifact.media_type == MARKDOWN_MEDIA_TYPE:
            validate_markdown_node_path(logical)
        else:
            validate_logical_path(logical)
        planned.append(
            replace(
                artifact,
                suggested_name=PurePosixPath(logical).name,
                logical_path=logical,
                metadata={**metadata, "logical_path": logical},
            )
        )

    return DocumentNodeLayoutPlan(identity=identity, root_name=root_name, artifacts=tuple(planned))


def relocated_markdown_bytes(
    artifact: ArtifactManifest,
    *,
    artifacts: tuple[ArtifactManifest, ...],
) -> bytes:
    """Read generated Markdown and rewrite known artifact links after relocation."""

    raw = Path(artifact.staging_path).read_bytes()
    bom = b"\xef\xbb\xbf" if raw.startswith(b"\xef\xbb\xbf") else b""
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw
    if artifact.logical_path is None:
        return raw
    source_parent = PurePosixPath(artifact.logical_path).parent.as_posix()
    replacements: dict[str, str] = {}
    for target in artifacts:
        if target.logical_path is None or target.artifact_id == artifact.artifact_id:
            continue
        relative = posixpath.relpath(target.logical_path, start=source_parent)
        for old in _artifact_reference_names(target):
            replacements.setdefault(old, relative)
    if not replacements:
        return bom + text.encode("utf-8")
    rewritten = _rewrite_known_links(text, replacements)
    return bom + rewritten.encode("utf-8")


def _markdown_child_label(identity: ConversionIdentity, artifact: ArtifactManifest) -> str:
    source_kind = str(artifact.metadata.get("source_kind", ""))
    if source_kind == "gongwen_attachment":
        ordinal = artifact.metadata.get("attachment_ordinal")
        suffix = f"{int(ordinal):02d}" if isinstance(ordinal, int) and ordinal > 0 else ""
        title = artifact.metadata.get("attachment_title")
        title_suffix = f"-{title}" if isinstance(title, str) and title.strip() else ""
        return sanitize_node_label(f"{identity.source_stem}_附件{suffix}{title_suffix}")
    raw = Path(artifact.suggested_name).stem or "子文档"
    if raw.casefold() == identity.source_stem.casefold():
        raw = f"{identity.source_stem}_子文档"
    return sanitize_node_label(raw)


def _markdown_role(artifact: ArtifactManifest) -> str:
    if artifact.metadata.get("source_kind") == "gongwen_attachment":
        return "attachment"
    if artifact.metadata.get("ocr") is True:
        return "ocr_fragment"
    return "derived_document"


def _portable_resource_name(value: str) -> str:
    path = PurePosixPath(value.replace("\\", "/"))
    basename = path.name or "resource"
    stem, suffix = os.path.splitext(basename)
    safe_stem = sanitize_node_label(stem or "resource", max_length=120)
    safe_suffix = re.sub(r"[^A-Za-z0-9._-]", "", suffix)[:24]
    return f"{safe_stem}{safe_suffix}"


def _artifact_reference_names(artifact: ArtifactManifest) -> set[str]:
    values = {
        artifact.suggested_name,
        Path(artifact.staging_path).name,
        artifact.staging_path,
        artifact.staging_path.replace("\\", "/"),
    }
    source_name = artifact.metadata.get("source_suggested_name")
    if isinstance(source_name, str) and source_name:
        values.add(source_name)
    return {value for value in values if value}


def _rewrite_known_links(text: str, replacements: dict[str, str]) -> str:
    normalized = {key.replace("\\", "/"): value for key, value in replacements.items()}

    def replace_target(raw: str) -> str:
        decoded = unquote(raw).replace("\\", "/")
        path, marker, anchor = decoded.partition("#")
        replacement = normalized.get(path) or normalized.get(PurePosixPath(path).name)
        if replacement is None:
            return raw
        return replacement + (f"#{anchor}" if marker else "")

    def markdown_sub(match: re.Match[str]) -> str:
        return f"{match.group('prefix')}{replace_target(match.group('target'))}{match.group('suffix')}"

    def wiki_sub(match: re.Match[str]) -> str:
        body = match.group("body")
        target, separator, alias = body.partition("|")
        replaced = replace_target(target.strip())
        suffix = f"|{alias}" if separator else ""
        return f"{match.group('prefix')}{replaced}{suffix}{match.group('suffix')}"

    def html_sub(match: re.Match[str]) -> str:
        return f"{match.group('prefix')}{replace_target(match.group('target'))}{match.group('suffix')}"

    text = _MARKDOWN_LINK.sub(markdown_sub, text)
    text = _WIKI_LINK.sub(wiki_sub, text)
    return _HTML_LINK.sub(html_sub, text)


__all__ = [
    "DOCUMENT_NODE_MANIFEST_MEDIA_TYPE",
    "DocumentNodeLayoutPlan",
    "has_markdown_artifacts",
    "plan_document_node_layout",
    "relocated_markdown_bytes",
]
