"""Fake implementation of PluginExecutionContext protocol."""

from __future__ import annotations

from typing import Any


class FakeExecutionContext:
    """Fake execution context wiring all support fakes together.

    Fields use ``Any`` intentionally — this fake exists to glue together
    individual fakes (FakePluginLogger, FakeProgressSink, etc.) without
    importing every protocol from the production codebase.
    """

    def __init__(
        self,
        request: Any,
        workspace: Any,
        config: Any,
        progress: Any,
        cancellation: Any,
        logger: Any,
        *,
        numbering_registry: Any = None,
        heading_cleanup_rules: Any = (),
        proofread_rules: Any = None,
        ocr_blockquote_title: str = "",
        document_style_catalog: Any = None,
        markdown_export_semantics: Any = None,
    ) -> None:
        self._request = request
        self._workspace = workspace
        self._config = config
        self._progress = progress
        self._cancellation = cancellation
        self._logger = logger
        self._numbering_registry = numbering_registry
        self._heading_cleanup_rules = tuple(heading_cleanup_rules or ())
        self._proofread_rules = proofread_rules
        self._ocr_blockquote_title = ocr_blockquote_title
        self._document_style_catalog = document_style_catalog
        self._markdown_export_semantics = markdown_export_semantics

    @property
    def request(self) -> Any:
        return self._request

    @property
    def workspace(self) -> Any:
        return self._workspace

    @property
    def config(self) -> Any:
        return self._config

    @property
    def progress(self) -> Any:
        return self._progress

    @property
    def cancellation(self) -> Any:
        return self._cancellation

    @property
    def logger(self) -> Any:
        return self._logger

    @property
    def numbering_registry(self) -> Any:
        return self._numbering_registry

    @property
    def heading_cleanup_rules(self) -> Any:
        return self._heading_cleanup_rules

    @property
    def proofread_rules(self) -> Any:
        return self._proofread_rules

    @property
    def ocr_blockquote_title(self) -> str:
        return self._ocr_blockquote_title

    @property
    def markdown_export_semantics(self) -> Any:
        return self._markdown_export_semantics

    @property
    def document_style_catalog(self) -> Any:
        return self._document_style_catalog
