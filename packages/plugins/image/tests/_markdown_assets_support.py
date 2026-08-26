"""Integration tests for image→Markdown asset/embed handling.

Covers the "embed" image_mode path that produces ``{{IMAGE:...}}``
placeholders via the shared ``docwen_core.links`` utilities.

Verification targets from F-H2-008:
- ``format_image_placeholder`` is consumed by the image converter.
- Embed mode is a real user-visible path, not dead code.
"""

from __future__ import annotations

import base64
import re
import tempfile
from pathlib import Path
from typing import Any

import pytest

from docwen_core.links import format_image_placeholder, is_data_uri_image

pytestmark = [pytest.mark.golden, pytest.mark.contract]


def _ocr_success(text: str) -> Any:
    from docwen_core.text.ocr import OcrOutcome, OcrStatus

    return OcrOutcome(OcrStatus.SUCCESS, text=text)


def _build_fake_context(
    input_path: str,
    staging_dir: str,
    target_format: str,
    options: dict | None = None,
    action_name: str = "",
    extra_input_paths: list[str] | None = None,
    *,
    pre_cancelled: bool = False,
    config_snapshot: dict[str, object] | None = None,
    ocr_blockquote_title: str = "",
):
    """Build a converter context at the same policy boundary as Runtime."""
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.export_semantics import MarkdownExportSemantics
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    paths = [input_path, *(extra_input_paths or [])]
    file_refs = [FileRef(path=p, format=Path(p).suffix.lstrip("."), category="image") for p in paths]
    snapshot = config_snapshot or {}
    request = ConversionRequest(
        request_id="test-image-embed-001",
        input_refs=file_refs,
        target_format=target_format,
        action_name=action_name,
        options=options or {},
        config_snapshot=snapshot,
        output_policy=OutputPolicy(),
    )
    token = CancellationToken()
    if pre_cancelled:
        token.cancel("test cancellation")
    admitted_export_semantics = MarkdownExportSemantics.from_config_snapshot(snapshot)
    return FakeExecutionContext(
        request=request,
        workspace=FakeWorkspaceHandle(input_path, staging_dir),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=token.view(),
        logger=FakePluginLogger(),
        ocr_blockquote_title=ocr_blockquote_title,
        markdown_export_semantics=admitted_export_semantics,
    )


__all__ = (
    "Path",
    "_build_fake_context",
    "_ocr_success",
    "base64",
    "format_image_placeholder",
    "is_data_uri_image",
    "pytest",
    "pytestmark",
    "re",
    "tempfile",
)
