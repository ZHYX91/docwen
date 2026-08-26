"""Portable identity and path rules for DocWen document nodes.

Every Markdown document published by DocWen is a directory node. If the
portable logical path of a Markdown artifact is ``A/A.md``, ``A`` is both its
node identity and its direct parent. The rules in this module are pure so the
GUI, CLI, Runtime, and Machine adapters share one implementation.
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import PurePosixPath
from typing import Final

DOCUMENT_NODE_SCHEMA: Final = "docwen.document_node.v1"

_ILLEGAL_FILENAME = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_WHITESPACE = re.compile(r"\s+")
_WINDOWS_RESERVED = frozenset(
    {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        *(f"COM{index}" for index in range(1, 10)),
        *(f"LPT{index}" for index in range(1, 10)),
    }
)
_SOURCE_TAGS: Final[dict[str, str]] = {
    "md": "Md",
    "markdown": "Markdown",
    "doc": "Doc",
    "docx": "Docx",
    "rtf": "Rtf",
    "odt": "Odt",
    "pdf": "Pdf",
    "ofd": "Ofd",
    "xps": "Xps",
    "html": "Html",
    "htm": "Html",
    "mhtml": "Mhtml",
    "mht": "Mhtml",
    "epub": "Epub",
    "txt": "Txt",
    "csv": "Csv",
    "tsv": "Tsv",
    "xls": "Xls",
    "xlsx": "Xlsx",
    "ppt": "Ppt",
    "pptx": "Pptx",
    "png": "Png",
    "jpg": "Jpg",
    "jpeg": "Jpeg",
    "tif": "Tif",
    "tiff": "Tiff",
    "bmp": "Bmp",
    "gif": "Gif",
    "webp": "Webp",
}


class DocumentNodeValidationError(ValueError):
    """A document-node identity or logical path violates the portable rules."""


def canonical_source_tag(source_format: str) -> str:
    """Return the stable public ``fromX`` tag for an admitted source format."""

    normalized = source_format.strip().lower().lstrip(".")
    if not normalized:
        return "Unknown"
    known = _SOURCE_TAGS.get(normalized)
    if known is not None:
        return known
    tokens = re.findall(r"[A-Za-z0-9]+", normalized)
    if not tokens:
        return "Unknown"
    return "".join(token[:1].upper() + token[1:].lower() for token in tokens)


def sanitize_node_label(label: str, *, max_length: int = 96) -> str:
    """Return a cross-platform filename segment with a stable hash fallback."""

    normalized = unicodedata.normalize("NFC", label)
    normalized = _ILLEGAL_FILENAME.sub("_", normalized)
    normalized = _WHITESPACE.sub(" ", normalized).strip(" .")
    if not normalized:
        normalized = "document"
    if normalized.upper() in _WINDOWS_RESERVED:
        normalized = f"_{normalized}"
    if len(normalized) <= max_length:
        return normalized
    digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:10]
    keep = max(1, max_length - len(digest) - 1)
    return f"{normalized[:keep].rstrip(' .')}_{digest}"


@dataclass(frozen=True, slots=True)
class ConversionIdentity:
    """One immutable identity shared by every artifact in a conversion task."""

    task_id: str
    source_stem: str
    source_format: str
    created_at: datetime

    @classmethod
    def create(
        cls,
        *,
        task_id: str,
        source_stem: str,
        source_format: str,
        created_at: datetime | None = None,
    ) -> ConversionIdentity:
        instant = created_at or datetime.now().astimezone()
        if instant.tzinfo is None or instant.utcoffset() is None:
            raise DocumentNodeValidationError("created_at must be timezone-aware")
        return cls(
            task_id=task_id,
            source_stem=sanitize_node_label(source_stem),
            source_format=source_format.strip().lower().lstrip(".") or "unknown",
            created_at=instant,
        )

    @property
    def timestamp(self) -> str:
        return self.created_at.strftime("%Y%m%d_%H%M%S")

    @property
    def created_at_utc(self) -> str:
        return self.created_at.astimezone(UTC).isoformat().replace("+00:00", "Z")

    @property
    def source_tag(self) -> str:
        return canonical_source_tag(self.source_format)

    def node_name(self, label: str | None = None, *, collision: int = 0) -> str:
        safe_label = sanitize_node_label(label or self.source_stem)
        collision_token = f"_{collision:03d}" if collision else ""
        return f"{safe_label}_{self.timestamp}{collision_token}_from{self.source_tag}"


@dataclass(frozen=True, slots=True)
class DocumentNodePath:
    """A normalized portable node path and its main Markdown path."""

    node_name: str
    parent: PurePosixPath = field(default_factory=lambda: PurePosixPath("."))

    def __post_init__(self) -> None:
        if sanitize_node_label(self.node_name) != self.node_name:
            raise DocumentNodeValidationError(f"unsafe node name: {self.node_name!r}")
        _validate_relative_parts(self.parent)

    @property
    def directory(self) -> PurePosixPath:
        return PurePosixPath(self.node_name) if self.parent == PurePosixPath(".") else self.parent / self.node_name

    @property
    def markdown(self) -> PurePosixPath:
        return self.directory / f"{self.node_name}.md"

    def child(self, node_name: str) -> DocumentNodePath:
        return DocumentNodePath(node_name=node_name, parent=self.directory)


def validate_logical_path(value: str) -> PurePosixPath:
    """Validate and return a normalized, relative POSIX artifact path."""

    if "\\" in value or not value or value.startswith("/"):
        raise DocumentNodeValidationError(f"logical path must be relative POSIX: {value!r}")
    path = PurePosixPath(value)
    _validate_relative_parts(path)
    if path.as_posix() != value:
        raise DocumentNodeValidationError(f"logical path must be normalized: {value!r}")
    return path


def validate_markdown_node_path(value: str) -> PurePosixPath:
    """Require the universal ``X/X.md`` invariant for a logical path."""

    path = validate_logical_path(value)
    if path.suffix.lower() != ".md" or path.parent.name != path.stem:
        raise DocumentNodeValidationError(f"Markdown must use X/X.md layout: {value!r}")
    return path


def _validate_relative_parts(path: PurePosixPath) -> None:
    if path == PurePosixPath("."):
        return
    if path.is_absolute() or any(part in {"", ".", ".."} for part in path.parts):
        raise DocumentNodeValidationError(f"unsafe relative path: {path.as_posix()!r}")


__all__ = [
    "DOCUMENT_NODE_SCHEMA",
    "ConversionIdentity",
    "DocumentNodePath",
    "DocumentNodeValidationError",
    "canonical_source_tag",
    "sanitize_node_label",
    "validate_logical_path",
    "validate_markdown_node_path",
]
