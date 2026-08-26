"""Speaker-notes failure reporting for presentation conversion."""

from __future__ import annotations

from types import SimpleNamespace
from typing import cast
from unittest.mock import Mock

import pytest

from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_core.protocols.execution_context import ConverterContext
from docwen_plugin_presentation.pptx_md.converter import PptxToMarkdownConverter
from docwen_plugin_presentation.pptx_md.request_policy import (
    PresentationMarkdownRequestPolicy,
)

pytestmark = pytest.mark.unit


def test_requested_notes_read_failure_is_recorded() -> None:
    class BrokenNotesSlide:
        shapes: tuple[object, ...] = ()

        @property
        def notes_slide(self) -> object:
            raise RuntimeError("notes relationship is corrupt")

    logger = Mock()
    context = cast(ConverterContext, SimpleNamespace(logger=logger))
    policy = PresentationMarkdownRequestPolicy(
        export=MarkdownExportSemantics(
            image_extraction_mode="file",
            ocr_placement_mode="main_md",
            image_link_style="markdown_embed",
            md_file_link_style="markdown_link",
        ),
        ocr_blockquote_title="",
    )
    payload_stats = {
        "charts": 0,
        "audio": 0,
        "video": 0,
        "chart_snapshot_unavailable_locations": [],
        "notes_unavailable_locations": [],
    }

    lines, *_ = PptxToMarkdownConverter()._process_slide(
        BrokenNotesSlide(),
        3,
        context,
        {"export_notes": True},
        request_policy=policy,
        payload_stats=payload_stats,
    )

    assert lines[:2] == ["## Slide 3", ""]
    assert payload_stats["notes_unavailable_locations"] == ["slide 3: RuntimeError: notes relationship is corrupt"]
    logger.warning.assert_called_once()
