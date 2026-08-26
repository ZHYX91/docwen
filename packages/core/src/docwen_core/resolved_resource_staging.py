"""Request-private staging for v4 resolved-document embedded resources.

The neutral document already authenticates every occurrence and payload.  This
module rechecks the typed snapshot at the physical boundary, writes only the
embedded bytes, and rewrites each complete authored token by its source range.
It never resolves an authored locator against a source directory.
"""

from __future__ import annotations

import hashlib
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path

from docwen_core.links._markdown_inline import parse_inline_link, parse_markdown_destination
from docwen_core.links._patterns import WIKI_EMBED_PATTERN
from docwen_core.models.resolved_numbering import (
    ResolvedDocument,
    ResolvedEmbeddedResource,
    ResolvedResourceOccurrence,
)
from docwen_core.models.semantic_document import SemanticBibliographyFragment
from docwen_core.semantic_bibliography import parse_semantic_bibliography

_WIKI_EMBED_RE = re.compile(WIKI_EMBED_PATTERN)
_WIKI_SIZE_RE = re.compile(r"(?P<width>[1-9][0-9]*)(?:x(?P<height>[1-9][0-9]*))?")
_EXTENSION_FOR_MEDIA_TYPE = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/webp": ".webp",
}


class ResolvedResourceStagingError(ValueError):
    """A typed resource snapshot cannot be staged without path inference."""


@dataclass(frozen=True, slots=True)
class ResolvedTextEdit:
    """One authenticated source replacement and its exact rendered range."""

    source_start: int
    source_end: int
    replacement: str
    result_start: int
    result_end: int

    def __post_init__(self) -> None:
        coordinates = (self.source_start, self.source_end, self.result_start, self.result_end)
        if any(type(value) is not int for value in coordinates):
            raise ResolvedResourceStagingError("text edit coordinates must be exact integers")
        if (
            self.source_start < 0
            or self.source_end <= self.source_start
            or self.result_start < 0
            or self.result_end <= self.result_start
            or type(self.replacement) is not str
            or not self.replacement
            or self.result_end - self.result_start != len(self.replacement)
        ):
            raise ResolvedResourceStagingError("text edit is not closed and canonical")


@dataclass(frozen=True, slots=True)
class ResolvedTextProjection:
    """Closed coordinate projection from authored Markdown to rendered Markdown."""

    source_length: int
    result_length: int
    edits: tuple[ResolvedTextEdit, ...]

    def __post_init__(self) -> None:
        if (
            type(self.source_length) is not int
            or type(self.result_length) is not int
            or self.source_length < 0
            or self.result_length < 0
            or type(self.edits) is not tuple
        ):
            raise ResolvedResourceStagingError("text projection bounds are not closed and canonical")
        shift = 0
        previous_end = 0
        for edit in self.edits:
            if (
                not isinstance(edit, ResolvedTextEdit)
                or edit.source_start < previous_end
                or edit.source_end > self.source_length
                or edit.result_start != edit.source_start + shift
            ):
                raise ResolvedResourceStagingError("text projection edits are not exact and ordered")
            shift += (edit.result_end - edit.result_start) - (edit.source_end - edit.source_start)
            previous_end = edit.source_end
        if self.result_length != self.source_length + shift:
            raise ResolvedResourceStagingError("text projection result length is inconsistent")

    def project_range(self, source_start: int, source_end: int) -> tuple[int, int]:
        """Project one non-resource source range without text search or diffing."""

        if source_start < 0 or source_end <= source_start or source_end > self.source_length:
            raise ResolvedResourceStagingError("source range is outside the authored Markdown")
        for edit in self.edits:
            if source_start < edit.source_end and source_end > edit.source_start:
                raise ResolvedResourceStagingError("source range overlaps a resource replacement")
        return self._project_boundary(source_start), self._project_boundary(source_end)

    def _project_boundary(self, source_offset: int) -> int:
        shift = 0
        for edit in self.edits:
            if source_offset <= edit.source_start:
                break
            if source_offset < edit.source_end:
                raise ResolvedResourceStagingError("source boundary falls inside a resource replacement")
            shift += (edit.result_end - edit.result_start) - (edit.source_end - edit.source_start)
        return source_offset + shift


@dataclass(frozen=True, slots=True)
class ResolvedResourceBinding:
    """One internal rendering projection and its request-owned resource paths."""

    rendered_markdown: str
    linked_paths: tuple[tuple[str, Path], ...]
    bibliography: SemanticBibliographyFragment | None
    text_projection: ResolvedTextProjection

    def path_for(self, resource_id: str) -> Path:
        for candidate, path in self.linked_paths:
            if candidate == resource_id:
                return path
        raise ResolvedResourceStagingError("resolved resource ID has no staged path")


def bind_resolved_document_resources(
    document: ResolvedDocument,
    resource_root: str | Path,
) -> ResolvedResourceBinding:
    """Bind exact image occurrences to embedded bytes under a fresh directory.

    All range/token/payload checks happen before the directory is published.
    Replacements are applied from the end of the authenticated source so an
    earlier replacement can never shift a later occurrence's coordinates.
    """

    root = Path(resource_root)
    linked = {resource.resource_id: resource for resource in document.resources if resource.role == "linked_resource"}
    bibliography_resources = tuple(resource for resource in document.resources if resource.role == "bibliography")
    if len(linked) + len(bibliography_resources) != len(document.resources):
        raise ResolvedResourceStagingError("resolved document contains an unknown resource role")
    if len(bibliography_resources) > 1:
        raise ResolvedResourceStagingError("resolved document contains multiple bibliographies")

    replacements = _resource_replacements(document, linked, root)
    text_projection = _text_projection(len(document.authored_markdown), replacements)
    rendered = document.authored_markdown
    for edit in reversed(text_projection.edits):
        rendered = rendered[: edit.source_start] + edit.replacement + rendered[edit.source_end :]
    if len(rendered) != text_projection.result_length:
        raise ResolvedResourceStagingError("rendered Markdown contradicts its closed text projection")

    bibliography = parse_semantic_bibliography(bibliography_resources[0].content) if bibliography_resources else None
    paths = _stage_linked_resources(tuple(linked.values()), root) if linked else ()
    return ResolvedResourceBinding(rendered, paths, bibliography, text_projection)


def _text_projection(
    source_length: int,
    replacements: tuple[tuple[ResolvedResourceOccurrence, str], ...],
) -> ResolvedTextProjection:
    shift = 0
    edits: list[ResolvedTextEdit] = []
    for occurrence, replacement in replacements:
        result_start = occurrence.source_start + shift
        result_end = result_start + len(replacement)
        edits.append(
            ResolvedTextEdit(
                source_start=occurrence.source_start,
                source_end=occurrence.source_end,
                replacement=replacement,
                result_start=result_start,
                result_end=result_end,
            )
        )
        shift += len(replacement) - (occurrence.source_end - occurrence.source_start)
    return ResolvedTextProjection(source_length, source_length + shift, tuple(edits))


def _resource_replacements(
    document: ResolvedDocument,
    linked: dict[str, ResolvedEmbeddedResource],
    root: Path,
) -> tuple[tuple[ResolvedResourceOccurrence, str], ...]:
    output: list[tuple[ResolvedResourceOccurrence, str]] = []
    previous_end = -1
    used: set[str] = set()
    for occurrence in document.resource_occurrences:
        if occurrence.source_start < previous_end or occurrence.source_end <= occurrence.source_start:
            raise ResolvedResourceStagingError("resolved resource occurrence ranges are invalid")
        token = document.authored_markdown[occurrence.source_start : occurrence.source_end]
        if token != occurrence.authored_token:
            raise ResolvedResourceStagingError("resolved resource token changed after admission")
        if hashlib.sha256(token.encode("utf-8")).hexdigest() != occurrence.source_slice_sha256:
            raise ResolvedResourceStagingError("resolved resource token hash changed after admission")
        resource = linked.get(occurrence.resource_id)
        if resource is None:
            raise ResolvedResourceStagingError("resolved occurrence has no embedded linked resource")
        path = root / _resource_filename(resource)
        replacement = _bind_complete_token(
            occurrence.authored_token,
            occurrence.authored_locator,
            path,
        )
        output.append((occurrence, replacement))
        used.add(resource.resource_id)
        previous_end = occurrence.source_end
    if used != set(linked):
        raise ResolvedResourceStagingError("embedded linked resource has no authenticated occurrence")
    return tuple(output)


def _bind_complete_token(token: str, authored_locator: str, path: Path) -> str:
    physical = path.absolute().as_posix()
    markdown = parse_inline_link(token, 0, image=True)
    if markdown is not None and markdown.end == len(token):
        destination = parse_markdown_destination(markdown.target, allow_image_size=True)
        if destination is None or destination.destination != authored_locator:
            raise ResolvedResourceStagingError("Markdown image locator contradicts its occurrence")
        return f"![{markdown.label}](<{physical}>{destination.suffix})"

    wiki = _WIKI_EMBED_RE.fullmatch(token)
    if wiki is None or wiki.group(1) != authored_locator:
        raise ResolvedResourceStagingError("wiki image locator contradicts its occurrence")
    display = wiki.group(2)
    size = _WIKI_SIZE_RE.fullmatch(display or "")
    if size is not None:
        width = size.group("width")
        height = size.group("height") or ""
        return f"![](<{physical}> ={width}x{height})"
    label = _escape_markdown_label(display or "")
    return f"![{label}](<{physical}>)"


def _escape_markdown_label(value: str) -> str:
    return value.replace("\\", "\\\\").replace("]", "\\]")


def _resource_filename(resource: ResolvedEmbeddedResource) -> str:
    extension = _EXTENSION_FOR_MEDIA_TYPE.get(resource.media_type)
    if extension is None:
        raise ResolvedResourceStagingError("linked resource media type is not stageable")
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9-]{0,63}", resource.resource_id):
        raise ResolvedResourceStagingError("linked resource ID is not a safe filename")
    return f"{resource.resource_id}{extension}"


def _stage_linked_resources(
    resources: tuple[ResolvedEmbeddedResource, ...],
    root: Path,
) -> tuple[tuple[str, Path], ...]:
    parent = root.parent
    if not root.is_absolute() or not parent.is_dir() or parent.is_symlink() or root.exists() or root.is_symlink():
        raise ResolvedResourceStagingError("request resource directory is not fresh")
    temporary = Path(tempfile.mkdtemp(prefix=f".{root.name}.", dir=parent))
    created: list[Path] = []
    try:
        for resource in resources:
            if (
                len(resource.content) != resource.size_bytes
                or hashlib.sha256(resource.content).hexdigest() != resource.sha256
            ):
                raise ResolvedResourceStagingError("embedded resource bytes changed after admission")
            target = temporary / _resource_filename(resource)
            created.append(target)
            with target.open("xb") as stream:
                stream.write(resource.content)
                stream.flush()
                os.fsync(stream.fileno())
            if target.read_bytes() != resource.content:
                raise ResolvedResourceStagingError("staged resource differs from authenticated bytes")
        temporary.rename(root)
    except BaseException:
        for path in reversed(created):
            path.unlink(missing_ok=True)
        temporary.rmdir()
        raise
    return tuple((resource.resource_id, root / _resource_filename(resource)) for resource in resources)


__all__ = [
    "ResolvedResourceBinding",
    "ResolvedResourceStagingError",
    "ResolvedTextEdit",
    "ResolvedTextProjection",
    "bind_resolved_document_resources",
]
