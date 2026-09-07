"""ConversionRequest — the input contract for a conversion task."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docwen_core.models.conversion_manifest import ConversionManifestContext
from docwen_core.models.document_node import ConversionIdentity
from docwen_core.models.file_ref import FileRef

PRECONVERSION_INTERMEDIATES_OPTION = "_docwen_preconversion_intermediates"
"""Internal request option carrying pre-conversion artifacts to finalize."""


@dataclass(slots=True)
class OutputPolicy:
    """Policy controlling where and how output files are written.

    Defined here (not in protocols) because it is a plain data object
    that crosses every boundary.
    """

    output_dir: str | None = None
    """Explicit output parent directory.  ``None`` means same-as-input."""

    output_path: str | None = None
    """Exact path for ungrouped output; mutually exclusive with ``output_dir``.

    Conversions from or to Markdown publish result directories and must use
    ``output_dir``.  The sole exception is an explicit in-place transform
    where ``output_path`` is the input path itself.
    """

    date_subfolder: str = ""
    """Date sub-folder pattern.

    Valid values:
    - ``""`` — no date sub-folder
    - ``"iso"`` — ISO 8601 (``YYYY-MM-DD``)
    - ``"compact"`` — compact (``YYYYMMDD``)
    - ``"chinese"`` — Chinese (``YYYY年MM月DD日``)
    """

    overwrite_mode: str = "rename"
    """Overwrite strategy.

    Valid values:
    - ``"error"`` — fail when the target already exists
    - ``"rename"`` — append a numeric suffix
    - ``"overwrite"`` — replace existing file
    - ``"skip"`` — do nothing if target exists
    """

    write_artifacts: bool = True
    """Whether staging artifacts are published to a final destination.

    Diagnostic operations may set this to ``False`` to return structured
    findings without creating a user-visible report file.  Staging remains
    private and is cleaned by the runtime in either mode.
    """

    open_after_done: bool = False
    """Whether to open the output folder after completion.

    .. note::

        This is a GUI presentation-layer behaviour.  CLI and headless
        modes will ignore it.  It lives in ``OutputPolicy`` (not in a
        GUI-specific model) so that the full user intent travels through
        the application → runtime chain without requiring the runtime to
        know about GUI concepts.
    """

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_dir": self.output_dir,
            "output_path": self.output_path,
            "date_subfolder": self.date_subfolder,
            "overwrite_mode": self.overwrite_mode,
            "write_artifacts": self.write_artifacts,
            "open_after_done": self.open_after_done,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> OutputPolicy:
        return cls(
            output_dir=data.get("output_dir"),
            output_path=data.get("output_path"),
            date_subfolder=data.get("date_subfolder", ""),
            overwrite_mode=data.get("overwrite_mode", "rename"),
            write_artifacts=data.get("write_artifacts", True),
            open_after_done=data.get("open_after_done", False),
        )


@dataclass(slots=True)
class ConversionRequest:
    """A request to convert one or more input files.

    This is the primary input contract between application layer and runtime.

    When multiple input files are provided, the runtime MAY split them
    into individual ``WorkerRequest`` instances (one per file).  For
    operations that require multiple inputs (e.g. merge, compare), the
    runtime may pass the full list to a single worker.
    """

    request_id: str
    """Unique identifier for this request (UUID or similar)."""

    input_refs: list[FileRef]
    """One or more input file references."""

    target_format: str
    """Desired output format (e.g. ``"docx"``, ``"md"``, ``"pdf"``)."""

    action_name: str = ""
    """Named action override (``"validate"``, ``"merge_tables"``, etc.)."""

    options: dict[str, Any] = field(default_factory=dict)
    """Typed conversion options (merged from config + CLI/GUI overrides)."""

    output_policy: OutputPolicy = field(default_factory=OutputPolicy)
    """Output placement policy."""

    config_snapshot: dict[str, Any] = field(default_factory=dict)
    """Read-only config snapshot available to plugins."""

    manifest_context: ConversionManifestContext | None = None
    """Optional typed sidecar-manifest context frozen at request admission."""

    conversion_identity: ConversionIdentity | None = None
    """Runtime-frozen source label and timestamp, shared by rendering and publication."""

    @property
    def source_stem(self) -> str:
        if self.conversion_identity is not None:
            return self.conversion_identity.source_stem
        source = next((ref for ref in self.input_refs if ref.input_role in {"source", "neutral_document"}), None)
        return Path(source.logical_path or source.path).stem if source is not None else "document"

    def to_dict(self) -> dict[str, Any]:
        return {
            "request_id": self.request_id,
            "input_refs": [r.to_dict() for r in self.input_refs],
            "target_format": self.target_format,
            "action_name": self.action_name,
            "options": dict(self.options),
            "output_policy": self.output_policy.to_dict(),
            "config_snapshot": dict(self.config_snapshot),
            "manifest_context": self.manifest_context.to_dict() if self.manifest_context else None,
            "conversion_identity": self.conversion_identity.to_dict() if self.conversion_identity else None,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConversionRequest:
        return cls(
            request_id=data["request_id"],
            input_refs=[FileRef.from_dict(r) for r in data["input_refs"]],
            target_format=data["target_format"],
            action_name=data.get("action_name", ""),
            options=dict(data.get("options", {})),
            output_policy=OutputPolicy.from_dict(data.get("output_policy", {})),
            config_snapshot=dict(data.get("config_snapshot", {})),
            manifest_context=(
                ConversionManifestContext.from_dict(data["manifest_context"])
                if isinstance(data.get("manifest_context"), dict)
                else None
            ),
            conversion_identity=(
                ConversionIdentity.from_dict(data["conversion_identity"])
                if isinstance(data.get("conversion_identity"), dict)
                else None
            ),
        )
