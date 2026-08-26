"""Fake ConverterPlugin for testing plugin registries."""

from __future__ import annotations

from docwen_core.models.result import ConversionResult
from docwen_core.protocols.execution_context import PluginExecutionContext


class FakePlugin:
    """A fake plugin conforming to ConverterPlugin protocol."""

    def __init__(self, manifest) -> None:
        self._manifest = manifest

    @property
    def manifest(self):
        return self._manifest

    def can_handle(self, source_format, target_format, action_name=""):
        for r in self._manifest.routes:
            if (
                r.source_format == source_format
                and r.target_format == target_format
                and (r.action_name == action_name or not action_name)
            ):
                return True
        return False

    def convert(self, context: PluginExecutionContext) -> ConversionResult:
        raise NotImplementedError("Fake plugin does not convert")
