"""FileRef — reference to an input file for conversion."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class FileRef:
    """Reference to an input file.

    Carries enough metadata for route resolution and plugin dispatch
    without requiring the file to be opened or parsed.

    Validation note: *format* and *category* are free-form strings.
    Production code should validate them against ``FORMAT_CATEGORY``
    from ``docwen_core.formats.categories`` before constructing a
    ``FileRef``.  The class itself does not enforce the constraint
    so that unknown/experimental formats can still be represented.
    """

    path: str
    """Absolute path to the input file."""

    format: str
    """Detected format (e.g. ``"markdown"``, ``"docx"``, ``"pdf"``).
    Should be a lowercase key present in ``FORMAT_CATEGORY``."""

    category: str
    """Detected category (e.g. ``"document"``, ``"spreadsheet"``, ``"image"``).
    Should be one of the ``CATEGORY_*`` constants in ``docwen_core.formats``."""

    encoding: str = "utf-8"
    """Text encoding (meaningful for text-based formats)."""

    warning_message: str = ""
    """Warning message from format validation (e.g. extension/content
    mismatch).  Empty string when everything is fine."""

    size_bytes: int = 0
    """File size in bytes (0 = unknown)."""

    input_kind: str = "resource"
    """Provider-neutral input kind (``document`` or ``resource``)."""

    input_role: str = "source"
    """Role of this input within the request's declared input shape."""

    logical_path: str = ""
    """Case-sensitive POSIX path in the request's virtual input root."""

    media_type: str = ""
    """Declared media type carried unchanged for typed request resources.

    Source-file admission may replace this value with the media type derived
    from the admitted content format. Non-source resources are not routed by
    suffix or content sniffing, so their declared media type remains part of
    the request contract.
    """

    metadata: dict[str, Any] = field(default_factory=dict)
    """Extra metadata from detection (page count, sheet names, etc.).
    A shallow copy is made on serialisation so the caller retains ownership."""

    def to_dict(self) -> dict[str, Any]:
        """Serialise to a plain dict."""
        return {
            "path": self.path,
            "format": self.format,
            "category": self.category,
            "encoding": self.encoding,
            "warning_message": self.warning_message,
            "size_bytes": self.size_bytes,
            "input_kind": self.input_kind,
            "input_role": self.input_role,
            "logical_path": self.logical_path,
            "media_type": self.media_type,
            "metadata": dict(self.metadata),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> FileRef:
        """Deserialise from a plain dict."""
        return cls(
            path=data["path"],
            format=data["format"],
            category=data["category"],
            encoding=data.get("encoding", "utf-8"),
            warning_message=data.get("warning_message", ""),
            size_bytes=data.get("size_bytes", 0),
            input_kind=data.get("input_kind", "resource"),
            input_role=data.get("input_role", "source"),
            logical_path=data.get("logical_path", ""),
            media_type=data.get("media_type", ""),
            metadata=dict(data.get("metadata", {})),
        )
