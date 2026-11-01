"""Central conversion layout planning and Markdown link relocation."""

from __future__ import annotations

import os
import posixpath
import re
from dataclasses import dataclass, replace
from datetime import datetime
from html import escape, unescape
from html.entities import html5
from pathlib import Path, PurePosixPath
from urllib.parse import unquote

from docwen_core.links import rewrite_markdown_resource_targets
from docwen_core.models import (
    DOCUMENT_NODE_SCHEMA,
    ArtifactManifest,
    ConversionIdentity,
    DocumentNodePath,
    sanitize_node_label,
    validate_artifact_ownership,
    validate_logical_path,
    validate_markdown_node_path,
)

MARKDOWN_MEDIA_TYPE = "text/markdown"

_WINDOWS_DRIVE_TARGET = re.compile(r"^[A-Za-z]:[\\/]")
_URI_SCHEME_TARGET = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")
_LOCAL_LINK_UNSAFE = frozenset("%#? \t\r\n<>\"'()[]|&")
_HTML_ATTRIBUTE_REFERENCE = re.compile(r"&#(?:[xX][0-9a-fA-F]+|[0-9]+);?|&[A-Za-z][A-Za-z0-9]*;?")


def _decode_html_attribute(value: str) -> str:
    """Decode references using HTML's attribute-context legacy-name rule."""

    def decode(match: re.Match[str]) -> str:
        token = match.group()[1:]
        if token.startswith("#"):
            return unescape(match.group())
        for length in range(len(token), 0, -1):
            name = token[:length]
            if name not in html5:
                continue
            following = token[length : length + 1] or value[match.end() : match.end() + 1]
            if (
                not name.endswith(";")
                and following
                and (following == "=" or (following.isascii() and following.isalnum()))
            ):
                return match.group()
            return html5[name] + token[length:]
        return match.group()

    return _HTML_ATTRIBUTE_REFERENCE.sub(decode, value)


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
            if (
                artifact.metadata.get("document_node_role") == "primary"
                and relative.parent == PurePosixPath(".")
                and relative.stem == self.root_name
            ):
                relative = PurePosixPath(f"{root_name}{relative.suffix}")
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


def plan_single_artifact_layout(
    *,
    task_id: str,
    artifacts: list[ArtifactManifest] | tuple[ArtifactManifest, ...],
    input_path: str,
    identity: ConversionIdentity | None = None,
) -> list[ArtifactManifest]:
    """Name a direct conversion result without inventing a document directory."""
    source = Path(input_path)
    identity = identity or ConversionIdentity.create(
        task_id=task_id,
        source_stem=source.stem or "document",
        source_format=source.suffix.lstrip(".") or "unknown",
        source_name=source.name,
    )
    planned: list[ArtifactManifest] = []
    for artifact in artifacts:
        if artifact.kind in {"manifest", "log"} or artifact.metadata.get("source") == "preconversion":
            planned.append(artifact)
            continue
        suffix = Path(artifact.suggested_name).suffix or Path(artifact.staging_path).suffix
        label = _csv_label(identity, artifact)
        name = f"{identity.node_name(label)}{suffix}"
        planned.append(
            replace(
                artifact,
                suggested_name=name,
                logical_path=name,
                metadata={**artifact.metadata, "source_suggested_name": artifact.suggested_name},
            )
        )
    return planned


def plan_document_node_layout(
    *,
    task_id: str,
    artifacts: list[ArtifactManifest] | tuple[ArtifactManifest, ...],
    input_path: str,
    created_at: datetime | None = None,
    root_collision: int = 0,
    identity: ConversionIdentity | None = None,
) -> DocumentNodeLayoutPlan:
    """Group a conversion's deliverables, keeping every Markdown at ``X/X.md``."""

    validate_artifact_ownership(artifacts)
    deliverables = [artifact for artifact in artifacts if artifact.kind not in {"manifest", "log"}]
    if not deliverables:
        raise ValueError("document-node layout requires at least one deliverable")
    preferred = [artifact for artifact in deliverables if artifact.is_primary]
    primary = preferred[0] if preferred else deliverables[0]
    source = Path(input_path)
    source_stem = source.stem or Path(primary.suggested_name).stem or "document"
    source_format = source.suffix.lstrip(".") or str(primary.metadata.get("source_format", "unknown"))
    identity = identity or ConversionIdentity.create(
        task_id=task_id,
        source_stem=source_stem,
        source_format=source_format,
        created_at=created_at,
        source_name=source.name,
    )
    root_name = identity.node_name(collision=root_collision)
    root = DocumentNodePath(root_name)

    planned: list[ArtifactManifest] = []
    used_paths: set[str] = set()
    child_labels: dict[str, int] = {}
    resource_names: dict[str, int] = {}
    csv_names: set[str] = set()
    reserved_csv_names = {
        identity.node_name(label).casefold()
        for artifact in artifacts
        if (label := _csv_label(identity, artifact)) is not None
    }
    by_id = {artifact.artifact_id: artifact for artifact in artifacts}
    directories: dict[str, PurePosixPath] = {}

    def depth(artifact: ArtifactManifest) -> int:
        count = 0
        while artifact.ownership is not None:
            artifact = by_id[artifact.ownership.owner_artifact_id]
            count += 1
        return count

    for artifact in sorted(artifacts, key=lambda item: (depth(item), item is not primary)):
        owner_directory = (
            directories[artifact.ownership.owner_artifact_id] if artifact.ownership is not None else root.directory
        )
        metadata = {
            **artifact.metadata,
            "document_node_schema": DOCUMENT_NODE_SCHEMA,
            "node_root": root_name,
            "source_suggested_name": artifact.suggested_name,
        }
        csv_label = _csv_label(identity, artifact)
        if csv_label is not None:
            csv_name = identity.node_name(csv_label)
            collision = 1
            while csv_name.casefold() in csv_names or (collision > 1 and csv_name.casefold() in reserved_csv_names):
                index = artifact.metadata.get("sheet_index", artifact.metadata.get("table_index", len(csv_names)))
                ordinal = int(index) + 1 if isinstance(index, int) else len(csv_names) + 1
                disambiguator = f"_{ordinal}" if collision == 1 else f"_{ordinal}_{collision}"
                csv_name = identity.node_name(f"{csv_label}{disambiguator}")
                collision += 1
            csv_names.add(csv_name.casefold())
            logical = (root.directory / f"{csv_name}.csv").as_posix()
            metadata["document_node_role"] = "primary" if artifact is primary else "worksheet"
        elif artifact is primary:
            suffix = (
                ".md"
                if artifact.media_type == MARKDOWN_MEDIA_TYPE
                else Path(artifact.suggested_name).suffix or Path(artifact.staging_path).suffix
            )
            logical = (root.directory / f"{root_name}{suffix}").as_posix()
            metadata["document_node_role"] = "primary"
        elif artifact.media_type == MARKDOWN_MEDIA_TYPE:
            label = _markdown_child_label(identity, artifact)
            key = (owner_directory / label).as_posix().casefold()
            child_labels[key] = child_labels.get(key, 0) + 1
            if child_labels[key] > 1:
                label = f"{label}_{child_labels[key]:02d}"
            child_name = identity.node_name(label)
            logical = (owner_directory / child_name / f"{child_name}.md").as_posix()
            metadata["document_node_role"] = _markdown_role(artifact)
        else:
            resource_name = _portable_resource_name(artifact.suggested_name or Path(artifact.staging_path).name)
            key = (owner_directory / resource_name).as_posix().casefold()
            resource_names[key] = resource_names.get(key, 0) + 1
            if resource_names[key] > 1:
                stem, suffix = os.path.splitext(resource_name)
                resource_name = f"{stem}_{resource_names[key]:03d}{suffix}"
            logical = (owner_directory / resource_name).as_posix()
            metadata["document_node_role"] = "resource"
        if logical.casefold() in used_paths:
            raise ValueError(f"duplicate logical artifact path: {logical}")
        used_paths.add(logical.casefold())
        directories[artifact.artifact_id] = PurePosixPath(logical).parent
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

    planned_by_id = {artifact.artifact_id: artifact for artifact in planned}
    return DocumentNodeLayoutPlan(
        identity=identity, root_name=root_name, artifacts=tuple(planned_by_id[item.artifact_id] for item in artifacts)
    )


def _csv_label(identity: ConversionIdentity, artifact: ArtifactManifest) -> str | None:
    if artifact.media_type != "text/csv" or (
        artifact.ownership is not None and artifact.ownership.relation_type == "resource_of"
    ):
        return None
    sheet_name = artifact.metadata.get("sheet_name")
    if not isinstance(sheet_name, str) or not sheet_name.strip():
        index = artifact.metadata.get("sheet_index", artifact.metadata.get("table_index", 0))
        sheet_name = f"Sheet{int(index) + 1}" if isinstance(index, int) else "Sheet1"
    return f"{identity.source_stem}_{sheet_name}"


def relocated_markdown_bytes(
    artifact: ArtifactManifest,
    *,
    artifacts: tuple[ArtifactManifest, ...],
) -> bytes:
    """Read generated Markdown and rewrite known artifact links after relocation."""

    raw = Path(artifact.staging_path).read_bytes()
    if artifact.ownership is not None and artifact.ownership.relation_type == "resource_of":
        return raw
    bom = b"\xef\xbb\xbf" if raw.startswith(b"\xef\xbb\xbf") else b""
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw
    if artifact.logical_path is None:
        return raw
    source_parent = PurePosixPath(artifact.logical_path).parent.as_posix()
    declared_targets: dict[str, set[str]] = {}
    for target in artifacts:
        if target.logical_path is None or target.artifact_id == artifact.artifact_id:
            continue
        if (
            target.ownership is not None
            and target.ownership.relation_type == "resource_of"
            and target.ownership.owner_artifact_id != artifact.artifact_id
        ):
            continue
        relative = posixpath.relpath(target.logical_path, start=source_parent)
        for old in _artifact_reference_names(target):
            declared_targets.setdefault(_reference_identity(old), set()).add(relative)
    if not declared_targets:
        return bom + text.encode("utf-8")
    replacements = {name: next(iter(paths)) for name, paths in declared_targets.items() if len(paths) == 1}
    ambiguous = {name for name, paths in declared_targets.items() if len(paths) > 1}
    rewritten = _rewrite_known_links(text, replacements, ambiguous=ambiguous)
    return bom + rewritten.encode("utf-8")


def _markdown_child_label(identity: ConversionIdentity, artifact: ArtifactManifest) -> str:
    if artifact.ownership is not None and artifact.ownership.label:
        return sanitize_node_label(f"{identity.source_stem}_{artifact.ownership.label}")
    raw = Path(artifact.suggested_name).stem or "子文档"
    if raw.casefold() == identity.source_stem.casefold():
        raw = f"{identity.source_stem}_子文档"
    return sanitize_node_label(raw)


def _markdown_role(artifact: ArtifactManifest) -> str:
    if artifact.ownership is not None:
        return artifact.ownership.role
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
        values.add(PurePosixPath(source_name.replace("\\", "/")).name)
    return {value for value in values if value}


def _is_external_link_target(raw: str) -> bool:
    """Return whether a link target is lexically a URI or network location."""

    raw_path = raw.partition("#")[0]
    candidate = raw_path.lstrip()
    normalized = candidate.replace("\\", "/")
    if normalized.startswith("//"):
        return True
    if _WINDOWS_DRIVE_TARGET.match(candidate):
        return False
    return _URI_SCHEME_TARGET.match(candidate) is not None


def _serialize_local_link_path(path: str) -> str:
    """Percent-escape delimiters that would change Markdown/HTML link structure."""

    result: list[str] = []
    for char in path:
        if char not in _LOCAL_LINK_UNSAFE:
            result.append(char)
            continue
        result.extend(f"%{byte:02X}" for byte in char.encode("utf-8"))
    return "".join(result)


def _reference_identity(value: str) -> str:
    normalized = value.replace("\\", "/")
    offset = 0
    while normalized.startswith("./", offset):
        offset += 2
    return normalized[offset:]


def _rewrite_known_links(text: str, replacements: dict[str, str], *, ambiguous: set[str] | None = None) -> str:
    normalized = {_reference_identity(key): value for key, value in replacements.items()}
    ambiguous = {_reference_identity(key) for key in (ambiguous or set())}

    def replace_target(raw: str, syntax: str) -> str:
        semantic = _decode_html_attribute(raw) if syntax == "html" else raw
        if _is_external_link_target(semantic):
            return raw
        path_and_query, marker, anchor = semantic.partition("#")
        path, query_marker, query = path_and_query.partition("?")
        decoded_path = _reference_identity(unquote(path))
        if decoded_path in ambiguous:
            raise ValueError(f"ambiguous producer-declared artifact reference: {path}")
        replacement = normalized.get(decoded_path)
        if replacement is None:
            return raw
        serialized = _serialize_local_link_path(replacement)
        serialized += (f"?{query}" if query_marker else "") + (f"#{anchor}" if marker else "")
        return escape(serialized, quote=True) if syntax == "html" else serialized

    return rewrite_markdown_resource_targets(text, replace_target)


__all__ = [
    "DocumentNodeLayoutPlan",
    "has_markdown_artifacts",
    "plan_document_node_layout",
    "relocated_markdown_bytes",
]
