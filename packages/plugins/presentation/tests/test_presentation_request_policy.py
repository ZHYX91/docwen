from __future__ import annotations

from types import SimpleNamespace
from typing import cast

import pytest
from tests.support.config import FakeConfigView

from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_core.protocols.execution_context import ConverterContext
from docwen_plugin_presentation.pptx_md.request_policy import (
    build_presentation_markdown_request_policy,
)

pytestmark = pytest.mark.unit


def _context(snapshot: dict[str, object], *, title: str | None = None) -> ConverterContext:
    return cast(
        ConverterContext,
        SimpleNamespace(
            request=SimpleNamespace(config_snapshot=snapshot),
            config=FakeConfigView(snapshot),
            ocr_blockquote_title=title,
        ),
    )


def test_nonempty_snapshot_freezes_complete_presentation_export_policy() -> None:
    snapshot: dict[str, object] = {
        "link": {
            "format": {
                "image_link_style": "wiki_embed",
                "md_file_link_style": "wiki_link",
            }
        },
        "conversion": {
            "export": {
                "base64_compress_enabled": False,
                "base64_compress_threshold_kb": 7,
            },
            "ocr_output": {
                "show_blockquote_title": True,
                "blockquote_title_override_by_locale": {"en_US": "Snapshot override"},
            },
        },
        "export": {
            "to_md_image_extraction_mode": "embed",
            "to_md_ocr_placement_mode": "image_md",
        },
        "output": {"intermediate_files": {"save_to_output": True}},
        "gui": {"language": {"locale": "en_US"}},
    }
    policy = build_presentation_markdown_request_policy(
        _context(snapshot, title="Request title"),
        {},
    )

    assert policy.export.image_link_style == "wiki_embed"
    assert policy.export.md_file_link_style == "wiki_link"
    assert policy.export.image_extraction_mode == "embed"
    assert policy.export.ocr_placement_mode == "image_md"
    assert policy.export.export_base64_compress_enabled is False
    assert policy.export.export_base64_compress_threshold_kb == 7
    assert policy.export.save_intermediate_files is True
    assert policy.ocr_blockquote_title == "Request title"


def test_explicit_options_override_snapshot_and_base64_forces_inline_ocr() -> None:
    snapshot: dict[str, object] = {
        "link": {"format": {"image_link_style": "wiki_embed"}},
        "export": {
            "to_md_image_extraction_mode": "file",
            "to_md_ocr_placement_mode": "image_md",
        },
    }

    policy = build_presentation_markdown_request_policy(
        _context(snapshot, title=""),
        {
            "image_mode": "base64",
            "ocr_placement": "image_md",
            "image_link_style": "markdown_embed",
        },
    )

    assert policy.export.image_extraction_mode == "base64"
    assert policy.export.ocr_placement_mode == "main_md"
    assert policy.export.image_link_style == "markdown_embed"


def test_empty_snapshot_uses_deterministic_defaults() -> None:
    policy = build_presentation_markdown_request_policy(_context({}), {})

    assert policy.export == MarkdownExportSemantics()
    assert policy.ocr_blockquote_title == ""


def test_nonempty_snapshot_preserves_authoritative_disabled_ocr_title() -> None:
    policy = build_presentation_markdown_request_policy(
        _context(
            {"conversion": {"ocr_output": {"show_blockquote_title": False}}},
            title="",
        ),
        {},
    )

    assert policy.ocr_blockquote_title == ""
