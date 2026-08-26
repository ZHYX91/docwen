"""HubConversionContext — shared concrete context for intermediate-format links.

When a converter produces an intermediate hub file (e.g. PPT→PPTX→MD,
spreadsheet→XLSX→MD) it must hand a context to the downstream converter whose
``request`` points at the hub file and whose ``workspace`` exposes that hub
file as ``input_path``.  This module provides the single shared implementation
of that remapping so plugins do not each hand-roll a private proxy.

``HubConversionContext`` satisfies ``ConverterContext`` (the narrow protocol).
``HubWorkspaceHandle`` satisfies ``WorkspaceHandle`` by delegating every call
to the original workspace while overriding ``input_path``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.request import ConversionRequest
    from docwen_core.protocols.execution_context import (
        CancellationTokenView,
        ConverterContext,
        PluginLogger,
        ProgressSink,
        ReadOnlyConfigView,
        WorkspaceHandle,
    )


class HubWorkspaceHandle:
    """Workspace handle that exposes a hub intermediate file as ``input_path``.

    Every other operation (``staging_dir``, ``create_artifact_path``,
    ``add_artifact``) delegates to the original workspace.  Only
    ``input_path`` is overridden to point at the hub file produced by the
    upstream link.
    """

    def __init__(self, delegate: WorkspaceHandle, hub_input_path: str) -> None:
        self._delegate = delegate
        self.input_path = hub_input_path

    @property
    def staging_dir(self) -> str:
        return self._delegate.staging_dir

    def input_resources(self, role: str | None = None):
        return self._delegate.input_resources(role)

    def resource_by_logical_path(self, logical_path: str):
        return self._delegate.resource_by_logical_path(logical_path)

    def create_artifact_path(self, kind: str, suffix: str) -> str:
        return self._delegate.create_artifact_path(kind, suffix)

    def add_artifact(self, manifest: ArtifactManifest) -> None:
        self._delegate.add_artifact(manifest)

    @property
    def registered_artifacts(self) -> list[ArtifactManifest]:
        """Expose artifacts registered through the delegated workspace."""
        return self._delegate.registered_artifacts


class HubConversionContext:
    """Concrete ``ConverterContext`` for intermediate-format conversion links.

    Wraps an existing context (typically the plugin entry ``PluginExecutionContext``)
    and overrides ``request`` and ``workspace`` so a downstream internal
    converter receives a context that points at the hub intermediate file.
    The remaining fields (``config`` / ``progress`` / ``cancellation`` /
    ``logger``) are forwarded unchanged from the wrapped context.
    """

    def __init__(self, base: ConverterContext, request: ConversionRequest, workspace: HubWorkspaceHandle) -> None:
        self._base = base
        self.request = request
        self.workspace = workspace

    @property
    def config(self) -> ReadOnlyConfigView:
        return self._base.config

    @property
    def progress(self) -> ProgressSink:
        return self._base.progress

    @property
    def cancellation(self) -> CancellationTokenView:
        return self._base.cancellation

    @property
    def logger(self) -> PluginLogger:
        return self._base.logger
