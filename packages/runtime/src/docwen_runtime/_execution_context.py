"""RuntimeExecutionContext — concrete implementation of PluginExecutionContext.

This is the **only** execution context passed to plugins by the runtime.
It wires together the actual runtime components (workspace, cancellation,
config, progress, logger) behind the core protocol interfaces.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from docwen_core.cancellation import CancellationToken
from docwen_core.events.task_events import (
    make_artifact_ready,
    make_diagnostic,
    make_task_progress,
)
from docwen_core.models.result import ConversionDiagnostic
from docwen_core.models.task import TaskEvent
from docwen_core.protocols.execution_context import (
    CancellationTokenView,
    PluginLogger,
    ProgressSink,
    ReadOnlyConfigView,
)

if TYPE_CHECKING:
    from docwen_core.docx_styles import DocumentStyleCatalog
    from docwen_core.export_semantics import MarkdownExportSemantics
    from docwen_core.models.proofread import ProofreadRules
    from docwen_core.models.request import ConversionRequest
    from docwen_core.text.heading_numbering import HeadingCleanupRules
    from docwen_runtime.workspace.manager import WorkspaceHandle


class _RuntimeProgressSink:
    """Concrete ProgressSink that creates ``TaskEvent`` instances.

    Each ``report_*`` call constructs a proper ``TaskEvent`` via the
    standard factory functions and forwards it through the *on_event*
    callback so it reaches the application layer through
    ``TaskManager`` → ``RuntimePortAdapter`` → ``PresenterPort``.

    Uses a *shared_seq* mutable list (``[int]``) so sequence numbers
    are globally monotonic across both TaskManager-emitted and
    plugin-emitted events.
    """

    def __init__(
        self,
        task_id: str,
        on_event: Callable[[TaskEvent], None] | None = None,
        shared_seq: list[int] | None = None,
    ) -> None:
        self._task_id = task_id
        self._on_event = on_event
        self._shared_seq = shared_seq if shared_seq is not None else [0]
        self.events: list[TaskEvent] = []
        self.diagnostics: list[ConversionDiagnostic] = []

    def _next_seq(self) -> int:
        """Return the current sequence number and increment."""
        val = self._shared_seq[0]
        self._shared_seq[0] += 1
        return val

    def _emit(self, event: TaskEvent) -> None:
        """Store the event and forward it to the listener."""
        self.events.append(event)
        if self._on_event is not None:
            self._on_event(event)

    def report_progress(self, percent: float, message: str = "") -> None:
        event = make_task_progress(self._task_id, self._next_seq(), percent, message)
        self._emit(event)

    def report_diagnostic(self, level: str, message: str, code: str = "", location: str = "") -> None:
        self.diagnostics.append(
            ConversionDiagnostic(
                level=level,
                message=message,
                code=code,
                location=location,
            )
        )
        event = make_diagnostic(self._task_id, self._next_seq(), level, message, code, location)
        self._emit(event)

    def report_artifact_ready(self, artifact_id: str, suggested_name: str) -> None:
        event = make_artifact_ready(self._task_id, self._next_seq(), artifact_id, suggested_name)
        self._emit(event)


class _RuntimePluginLogger:
    """Concrete PluginLogger for runtime use."""

    def __init__(self, task_id: str) -> None:
        self.task_id = task_id
        self.messages: list[dict[str, Any]] = []

    def debug(self, message: str, **extra: object) -> None:
        self.messages.append({"level": "debug", "message": message, "task_id": self.task_id, **extra})

    def info(self, message: str, **extra: object) -> None:
        self.messages.append({"level": "info", "message": message, "task_id": self.task_id, **extra})

    def warning(self, message: str, **extra: object) -> None:
        self.messages.append({"level": "warning", "message": message, "task_id": self.task_id, **extra})

    def error(self, message: str, **extra: object) -> None:
        self.messages.append({"level": "error", "message": message, "task_id": self.task_id, **extra})


class _RuntimeReadOnlyConfigView:
    """Concrete ReadOnlyConfigView backed by a plain dict."""

    def __init__(self, values: dict[str, Any]) -> None:
        self._values = dict(values)

    def get(self, key: str, default: object = None) -> object:
        if key in self._values:
            return self._values[key]
        current: object = self._values
        for part in key.split("."):
            if not isinstance(current, dict) or part not in current:
                return default
            current = current[part]
        return current

    def get_plugin_config(self, plugin_id: str) -> dict[str, object]:
        plugin_cfg = self._values.get(plugin_id, {})
        if isinstance(plugin_cfg, dict):
            return dict(plugin_cfg)
        return {}


class RuntimeExecutionContext:
    """Concrete execution context passed to plugin ``convert()``.

    Satisfies ``docwen_core.protocols.execution_context.PluginExecutionContext``.

    Plugins receive this object and interact with it through the
    protocol interfaces — they never import or reference this class
    directly.
    """

    def __init__(
        self,
        request: ConversionRequest,
        workspace: WorkspaceHandle,
        config_snapshot: dict[str, Any],
        cancellation_token: CancellationToken,
        *,
        on_event: Callable[[TaskEvent], None] | None = None,
        shared_seq: list[int] | None = None,
        numbering_registry: Any = None,
        heading_cleanup_rules: HeadingCleanupRules = (),
        proofread_rules: ProofreadRules | None = None,
        ocr_blockquote_title: str = "",
        document_style_catalog: DocumentStyleCatalog | None = None,
        markdown_export_semantics: MarkdownExportSemantics,
    ) -> None:
        task_id = request.request_id

        self._request = request
        self._workspace = workspace
        self._config = _RuntimeReadOnlyConfigView(config_snapshot)
        self._progress = _RuntimeProgressSink(task_id, on_event=on_event, shared_seq=shared_seq)
        self._cancellation = cancellation_token
        self._logger = _RuntimePluginLogger(task_id)
        self._numbering_registry = numbering_registry
        self._heading_cleanup_rules = tuple(heading_cleanup_rules)
        self._proofread_rules = proofread_rules
        self._ocr_blockquote_title = ocr_blockquote_title
        self._document_style_catalog = document_style_catalog
        self._markdown_export_semantics = markdown_export_semantics

    @property
    def request(self) -> ConversionRequest:
        return self._request

    @property
    def workspace(self) -> WorkspaceHandle:
        """Return the workspace handle.

        Plugins receive this as the ``WorkspaceHandle`` protocol.
        They MUST NOT cast it to the concrete type or access
        implementation details.
        """
        return self._workspace

    @property
    def config(self) -> ReadOnlyConfigView:
        """Read-only config view (Protocol type)."""
        return self._config

    @property
    def progress(self) -> ProgressSink:
        """Progress sink (Protocol type).

        Plugin calls to ``report_progress()`` etc. create ``TaskEvent``
        instances that are forwarded through the *on_event* callback
        to the ``TaskManager`` and ultimately to the application layer.
        """
        return self._progress

    @property
    def reported_diagnostics(self) -> list[ConversionDiagnostic]:
        """Return diagnostics streamed by the plugin during this execution."""
        return list(self._progress.diagnostics)

    @property
    def cancellation(self) -> CancellationTokenView:
        """Return a **read-only view** of the cancellation token.

        Plugins CANNOT call ``cancel()`` — they only see
        ``is_cancelled`` and ``check()``.
        """
        return self._cancellation.view()

    @property
    def logger(self) -> PluginLogger:
        """Logger (Protocol type)."""
        return self._logger

    @property
    def numbering_registry(self) -> Any:
        """Numbering scheme registry injected by the runtime."""
        return self._numbering_registry

    @property
    def heading_cleanup_rules(self) -> HeadingCleanupRules:
        """Immutable heading cleanup rules compiled from the request snapshot."""
        return self._heading_cleanup_rules

    @property
    def ocr_blockquote_title(self) -> str:
        """Localized OCR title resolved from this request's snapshot."""
        return self._ocr_blockquote_title

    @property
    def markdown_export_semantics(self) -> MarkdownExportSemantics:
        """Immutable generic Markdown policy captured at request admission."""
        return self._markdown_export_semantics

    @property
    def document_style_catalog(self) -> DocumentStyleCatalog | None:
        """Return the strict request-owned DOCX style catalog, when applicable."""
        return self._document_style_catalog

    @property
    def proofread_rules(self) -> ProofreadRules | None:
        """Pure proofread rule data injected by the runtime."""
        return self._proofread_rules
