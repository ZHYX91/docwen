"""Core protocols — abstract interfaces for plugins and runtime components."""

from docwen_core.protocols.converter import ConverterPlugin
from docwen_core.protocols.execution_context import (
    CancellationTokenView,
    ConverterContext,
    DocumentStyleConverterContext,
    NumberingConverterContext,
    PluginExecutionContext,
    PluginLogger,
    ProgressSink,
    ProofreadConverterContext,
    ReadOnlyConfigView,
    WorkspaceHandle,
)
from docwen_core.protocols.hub_context import HubConversionContext, HubWorkspaceHandle

__all__ = [
    "CancellationTokenView",
    "ConverterContext",
    "ConverterPlugin",
    "DocumentStyleConverterContext",
    "HubConversionContext",
    "HubWorkspaceHandle",
    "NumberingConverterContext",
    "PluginExecutionContext",
    "PluginLogger",
    "ProgressSink",
    "ProofreadConverterContext",
    "ReadOnlyConfigView",
    "WorkspaceHandle",
]
