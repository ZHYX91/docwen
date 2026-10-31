"""Shared exact-two request fixtures for resolved-v4 conversion contracts."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.resolved_numbering import NUMBERING_EXPORT_PLAN_MEDIA_TYPE, RESOLVED_DOCUMENT_MEDIA_TYPE
from docwen_runtime.config.document_styles import build_document_style_catalog

from .conftest import PROJECT_ROOT

_FIXTURES = Path(__file__).parent / "fixtures" / "resolved_v4"
_NEUTRAL = _FIXTURES / "resolved-document.rich.json"
_PLAN = _FIXTURES / "numbering-export-plan.rich.json"


def _sha_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _refs(neutral: Path, plan: Path) -> tuple[FileRef, FileRef]:
    return (
        FileRef(
            path=str(neutral),
            format="markdown",
            category="document",
            size_bytes=neutral.stat().st_size,
            input_kind="document",
            input_role="neutral_document",
            logical_path="document.json",
            media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
        ),
        FileRef(
            path=str(plan),
            format="json",
            category="resource",
            size_bytes=plan.stat().st_size,
            input_kind="resource",
            input_role="numbering_export_plan",
            logical_path="numbering-plan.json",
            media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
        ),
    )


def _context(
    tmp_path: Path,
    refs: tuple[FileRef, ...],
    *,
    options: dict[str, Any] | None = None,
    style_locale: str = "zh_CN",
) -> tuple[FakeExecutionContext, FakeWorkspaceHandle]:
    staging = tmp_path / "staging"
    staging.mkdir(parents=True)
    request = ConversionRequest(
        request_id="resolved-v4-test",
        input_refs=list(refs),
        target_format="docx",
        options=options or {"locale": "zh_CN", "heading_merge_mode": "never"},
        output_policy=OutputPolicy(),
    )
    workspace = FakeWorkspaceHandle(refs[0].path if refs else "", str(staging), refs)
    styles = build_document_style_catalog(
        {"gui": {"language": {"locale": style_locale}}},
        locales_dir=PROJECT_ROOT / "i18n" / "locales",
    )
    return (
        FakeExecutionContext(
            request,
            workspace,
            FakeConfigView(),
            FakeProgressSink(),
            FakeCancellationTokenView(),
            FakePluginLogger(),
            document_style_catalog=styles,
        ),
        workspace,
    )
