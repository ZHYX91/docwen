"""Execution context protocols — the context passed to converters.

The protocols form a layered family so each converter declares exactly the
dependencies it needs (interface segregation):

    ConverterContext            (narrow base — request, workspace, config,
                                 progress, cancellation, logger)
      ├─ NumberingConverterContext   (+ request-owned numbering policy)
      ├─ ProofreadConverterContext   (+ proofread_rules)
      └─ PluginExecutionContext      (+ numbering, proofread and OCR presentation policy)

``ConverterContext`` is what *internal* converters (called via hub/intermediate
links) accept.  ``NumberingConverterContext`` / ``ProofreadConverterContext``
narrow that by one extra field for converters that actually inject numbering
or proofread rules.  ``PluginExecutionContext`` is the full context the runtime
hands to plugin *entry* ``convert()`` methods; it extends ``ConverterContext``
because it provides everything a narrow converter needs plus the injected
    numbering-policy, proofread-rule and OCR-presentation fields.

Callers that build an intermediate-format context (hub links) should use the
shared concrete ``HubConversionContext`` (see ``hub_context``) rather than
hand-rolling a private proxy.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Protocol

from docwen_core.models.request import ConversionRequest

if TYPE_CHECKING:
    from docwen_core.docx_styles import DocumentStyleCatalog
    from docwen_core.export_semantics import MarkdownExportSemantics
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.proofread import ProofreadRules
    from docwen_core.text.heading_numbering import HeadingCleanupRules


class WorkspaceHandle(Protocol):
    """Handle to the task workspace.

    Plugins use this to get staging paths and register artifacts.
    They MUST NOT write directly to the final output directory.

    .. warning::

        ``input_path`` and ``staging_dir`` expose raw filesystem paths.
        Plugins are trusted to write only within ``staging_dir``.
        Path-traversal safeguards are the runtime's responsibility.
    """

    @property
    def input_path(self) -> str:
        """Absolute path to the input file (convenience)."""
        ...

    @property
    def staging_dir(self) -> str:
        """Absolute path to the staging directory for this task."""
        ...

    def input_resources(self, role: str | None = None) -> tuple[FileRef, ...]:
        """Return only the request-declared inputs, optionally filtered by role."""
        ...

    def resource_by_logical_path(self, logical_path: str) -> FileRef | None:
        """Resolve an exact case-sensitive key in the declared virtual input root."""
        ...

    def create_artifact_path(self, kind: str, suffix: str) -> str:
        """Create and return a writable path inside the staging directory.

        Args:
            kind: Artifact kind (see ``ArtifactManifest`` kind constants:
                  ``ARTIFACT_KIND_PRIMARY``, ``ARTIFACT_KIND_AUXILIARY``,
                  ``ARTIFACT_KIND_IMAGE``, ``ARTIFACT_KIND_LOG``).
            suffix: File suffix including the dot (e.g. ``".docx"``, ``".png"``).

        Returns:
            An absolute path the plugin can write to.
        """
        ...

    def add_artifact(self, manifest: ArtifactManifest) -> None:
        """Register an ``ArtifactManifest`` with the workspace.

        The runtime uses registered artifacts during finalisation.
        """
        ...

    @property
    def registered_artifacts(self) -> list[ArtifactManifest]:
        """Return artifacts registered by the active conversion chain."""
        ...


class ProgressSink(Protocol):
    """Sink for progress events during conversion.

    Plugins call methods on this sink; the runtime serialises the
    events and delivers them to the application layer.
    """

    def report_progress(self, percent: float, message: str = "") -> None:
        """Report conversion progress (0.0 – 100.0)."""
        ...

    def report_diagnostic(self, level: str, message: str, code: str = "", location: str = "") -> None:
        """Report a diagnostic message (info / warning / error)."""
        ...

    def report_artifact_ready(self, artifact_id: str, suggested_name: str) -> None:
        """Announce that an artifact is ready for finalisation."""
        ...


class CancellationTokenView(Protocol):
    """Read-only view of a cancellation token.

    Plugins check this periodically in long-running loops.
    The token is the ONLY mechanism for cooperative cancellation —
    plugins must not access ``threading.Event``, Qt signals, or other
    concurrency primitives directly.

    This Protocol is the public API type.  The concrete implementation
    returned by ``CancellationToken.view()`` is the internal
    ``_CancellationTokenViewImpl`` — plugins should type-annotate
    against this Protocol, not the concrete class.
    """

    @property
    def is_cancelled(self) -> bool:
        """Return ``True`` if cancellation has been requested."""
        ...

    def check(self) -> None:
        """Raise ``CancellationRequested`` if cancellation has been requested.

        Plugins should call this before expensive operations, in long loops,
        and before large file writes.
        """
        ...


class PluginLogger(Protocol):
    """Minimal logger interface available to plugins.

    Plugins use this instead of importing ``logging`` or ``loguru``
    directly, so the runtime can control log routing.
    """

    def debug(self, message: str, **extra: object) -> None: ...
    def info(self, message: str, **extra: object) -> None: ...
    def warning(self, message: str, **extra: object) -> None: ...
    def error(self, message: str, **extra: object) -> None: ...


class ReadOnlyConfigView(Protocol):
    """Read-only view of configuration available to plugins.

    Plugins MUST NOT hold mutable references to config objects.
    """

    def get(self, key: str, default: object = None) -> object:
        """Get a config value by dotted key (e.g. ``"plugins.docx.image_mode"``)."""
        ...

    def get_plugin_config(self, plugin_id: str) -> dict[str, object]:
        """Get the entire config sub-tree for a plugin."""
        ...


class ConverterContext(Protocol):
    """Narrow execution context for internal converters.

    Internal converters (called via hub / intermediate-format links, not
    directly by the runtime) accept this protocol rather than the full
    ``PluginExecutionContext``.  It exposes only the six fields an internal
    converter touches: the request, the workspace, a read-only config view,
    a progress sink, a cancellation token view, and a logger.

    Converters that additionally need the request-owned numbering policy or the
    proofread rules bundle use ``NumberingConverterContext`` /
    ``ProofreadConverterContext`` instead.  ``PluginExecutionContext`` extends
    this protocol and is what the runtime hands to plugin entry ``convert()``
    methods.
    """

    @property
    def request(self) -> ConversionRequest:
        """The conversion request being executed."""
        ...

    @property
    def workspace(self) -> WorkspaceHandle:
        """Workspace handle for staging writes."""
        ...

    @property
    def config(self) -> ReadOnlyConfigView:
        """Read-only config view."""
        ...

    @property
    def progress(self) -> ProgressSink:
        """Progress sink for reporting."""
        ...

    @property
    def cancellation(self) -> CancellationTokenView:
        """Cancellation token for cooperative cancellation.

        Returns a read-only view (the ``CancellationTokenView`` Protocol).
        Plugins MUST NOT cast this back to a mutable ``CancellationToken``.
        """
        ...

    @property
    def logger(self) -> PluginLogger:
        """Logger for plugin diagnostics."""
        ...


class NumberingConverterContext(ConverterContext, Protocol):
    """Converter context that also exposes request-owned numbering policy.

    Used by internal converters that inject a numbering scheme registry
    (e.g. markdown numbering, document→markdown heading numbering).
    """

    @property
    def numbering_registry(self) -> Any:
        """Numbering scheme registry for heading numbering.

        Returns ``None`` when no registry has been configured (e.g. in
        tests that use a partial mock context).
        """
        ...

    @property
    def heading_cleanup_rules(self) -> HeadingCleanupRules:
        """Immutable compiled cleanup rules for this request."""
        ...


class DocumentStyleConverterContext(NumberingConverterContext, Protocol):
    """DOCX converter context with request-owned style locale data."""

    @property
    def document_style_catalog(self) -> DocumentStyleCatalog:
        """Strict immutable catalog for the exact admitted request locale."""
        ...


class ProofreadConverterContext(ConverterContext, Protocol):
    """Converter context that also exposes the proofread rules bundle.

    Used by internal converters that apply proofread rules (the proofread
    validators).
    """

    @property
    def proofread_rules(self) -> ProofreadRules | None:
        """Pure proofread rule bundle injected by the runtime."""
        ...


class PluginExecutionContext(ConverterContext, Protocol):
    """The complete execution context passed to a plugin's ``convert()`` method.

    This is the ONLY way plugins interact with the outside world.  It extends
    ``ConverterContext`` (the narrow base internal converters use) with the
    runtime-injected numbering-policy and proofread-rule fields.  Internal
    converters that need only one policy group accept the narrower
    ``NumberingConverterContext`` or ``ProofreadConverterContext`` instead.
    """

    @property
    def numbering_registry(self) -> Any:
        """Numbering scheme registry for heading numbering.

        Injected by the runtime at assembly time so plugins do not need
        to import ``docwen_runtime.numbering.registry``.  Returns ``None``
        when no registry has been configured (e.g. in tests that use a
        partial mock context).
        """
        ...

    @property
    def heading_cleanup_rules(self) -> HeadingCleanupRules:
        """Immutable compiled cleanup rules for this request."""
        ...

    @property
    def document_style_catalog(self) -> DocumentStyleCatalog | None:
        """Strict immutable DOCX style catalog when the route produces DOCX."""
        ...

    @property
    def proofread_rules(self) -> ProofreadRules | None:
        """Pure proofread rule bundle injected by the runtime."""
        ...

    @property
    def ocr_blockquote_title(self) -> str:
        """Request-owned Markdown title fragment for inline OCR output."""
        ...

    @property
    def markdown_export_semantics(self) -> MarkdownExportSemantics:
        """Immutable generic Markdown export policy frozen at admission."""
        ...
