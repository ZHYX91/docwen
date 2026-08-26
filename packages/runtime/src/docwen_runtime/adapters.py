"""Adapters that expose runtime capabilities through application port interfaces.

These classes implement the port protocols defined in
``docwen_application.ports`` so that bundle (or tests) can inject
runtime functionality into the application layer without the
application layer importing runtime internals.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import replace
from typing import TYPE_CHECKING, Any

from docwen_core.detection import enforce_file_admission
from docwen_core.models.conversion_manifest import ConversionManifestContext
from docwen_core.models.request import ConversionRequest
from docwen_core.models.task import TaskEvent
from docwen_runtime._request_admission import admit_markdown_ocr_options

if TYPE_CHECKING:
    from docwen_runtime.config.loader import ConfigLoader
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.manifest import OutputManifestWriter

_AGGREGATE_ACTIONS: frozenset[str] = frozenset(
    {
        "merge_pdfs",
        "merge_tables",
        "merge_images_to_tiff",
    }
)


class RuntimePortAdapter:
    """Adapts ``TaskManager`` to the ``RuntimePort`` protocol.

    This is the bridge that allows the application layer to delegate
    conversion execution to the runtime without importing runtime
    internals directly.

    Satisfies ``docwen_application.ports.runtime.RuntimePort`` structurally
    (no explicit inheritance — duck typing via Protocol).

    Events emitted during execution are collected in ``_collected_events``
    and can also be forwarded to an optional *event_callback* passed at
    construction time.

    When a *config_loader* is supplied, each admitted request captures an
    immutable configuration snapshot and derives missing Markdown OCR options
    from that same snapshot. No process-wide converter policy is mutated.
    """

    def __init__(
        self,
        task_manager: TaskManager,
        *,
        event_callback: Callable[[TaskEvent], None] | None = None,
        config_loader: ConfigLoader | None = None,
        capability_provider: Callable[[], dict[str, Any]] | None = None,
        output_manifest_writer: OutputManifestWriter | None = None,
    ) -> None:
        self._task_manager = task_manager
        self._available = True
        self._event_callback = event_callback
        self._collected_events: list[TaskEvent] = []
        self._config_loader = config_loader
        self._capability_provider = capability_provider
        self._output_manifest_writer = output_manifest_writer

    def execute(self, request: Any) -> Any:
        """Execute a conversion request.

        Accepts a ``ConversionRequest``.  Returns a ``ConversionResult``
        for single-file requests and aggregate requests, or
        ``list[ConversionResult]`` for regular batch requests (multiple
        ``input_refs`` without an aggregate action).

        Events emitted during execution are collected in
        ``_collected_events`` and forwarded to *event_callback* if set.

        When *config_loader* is available, the ``config_snapshot`` is
        populated from the merged configuration if the request does not
        already carry one.  Missing Markdown OCR language/locale options
        are then projected from that exact snapshot without mutating the
        caller's request.
        """
        req = request if isinstance(request, ConversionRequest) else ConversionRequest.from_dict(request)
        req = enforce_file_admission(req)

        # Admit one authoritative snapshot, then derive missing OCR language
        # and locale from that exact value.  Explicit request option keys
        # win, and no later live loader read can split the pair across reloads.
        config_snapshot = req.config_snapshot
        snapshot_already_admitted = req.manifest_context is not None
        if not config_snapshot and self._config_loader is not None and not snapshot_already_admitted:
            config_snapshot = self._config_loader.config.as_dict()
        req = admit_markdown_ocr_options(req, config_snapshot)
        if req.manifest_context is None:
            manifest_context = ConversionManifestContext.from_request_inputs(
                req.input_refs,
                req.config_snapshot,
            )
        else:
            manifest_context = req.manifest_context
        if req.manifest_context is None and manifest_context.policy.save_to_output:
            req = replace(
                req,
                manifest_context=manifest_context,
            )

        self._collected_events.clear()

        def collect_events(event: TaskEvent) -> None:
            self._collected_events.append(event)
            if self._event_callback is not None:
                self._event_callback(event)

        primary_count = sum(item.input_role in {"source", "neutral_document"} for item in req.input_refs)
        if primary_count == 1 or req.action_name in _AGGREGATE_ACTIONS:
            result = self._task_manager.execute_single(req, on_event=collect_events)
        else:
            result = self._task_manager.execute_batch(req, on_event=collect_events)
        return self.persist_output_manifests(req, result)

    def persist_output_manifests(self, request: Any, result: Any) -> Any:
        """Apply the optional sidecar policy without making it a base-port requirement."""
        writer = self._output_manifest_writer
        if writer is None:
            return result
        req = request if isinstance(request, ConversionRequest) else ConversionRequest.from_dict(request)
        if req.manifest_context is None:
            config_snapshot = req.config_snapshot
            if not config_snapshot and self._config_loader is not None:
                config_snapshot = self._config_loader.config.as_dict()
                req = replace(req, config_snapshot=config_snapshot)
            req = replace(
                req,
                manifest_context=ConversionManifestContext.from_request_inputs(
                    req.input_refs,
                    req.config_snapshot,
                ),
            )
        return writer.persist(req, result)

    def cancel(self, task_id: str) -> None:
        """Request cancellation of a running task."""
        self._task_manager.cancel(task_id)

    def reserve_cancellation(self, task_id: str) -> None:
        """Reserve Runtime cancellation state for an Application handoff."""
        self._task_manager.reserve_cancellation(task_id)

    def release_cancellation(self, task_id: str) -> None:
        """Release Runtime cancellation state after Application handoff."""
        self._task_manager.release_cancellation(task_id)

    @property
    def is_available(self) -> bool:
        """Whether the runtime is initialized and ready."""
        return self._available

    @property
    def collected_events(self) -> list[TaskEvent]:
        """Events collected during the most recent ``execute()`` call.

        Returns a copy so callers can't mutate internal state.
        """
        return list(self._collected_events)

    def describe_capabilities(self) -> dict[str, Any]:
        """Reflect the loaded runtime composition and current machine gates."""

        provider = self._capability_provider
        if provider is None:
            raise RuntimeError("runtime_capability_discovery_unavailable")
        return provider()

    def shutdown(self) -> None:
        """Shut down the runtime adapter."""
        self._task_manager.cancel_all()
        self._available = False
