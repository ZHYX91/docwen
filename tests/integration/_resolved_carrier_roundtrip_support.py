"""Small authored exact-two requests and real DOCX/Markdown conversion wiring."""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
    CaptionMaterialization,
    NumberingTarget,
    ResolvedDocumentTarget,
    canonicalize_numbering_plan,
)
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_runtime.config.document_styles import build_document_style_catalog
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle


def _context(root: Path, refs: tuple[FileRef, ...], *, options: dict[str, Any]) -> FakeExecutionContext:
    staging = root / "staging"
    staging.mkdir(parents=True)
    request = ConversionRequest(
        request_id="resolved-carrier-roundtrip",
        input_refs=list(refs),
        target_format="docx",
        options=options,
        output_policy=OutputPolicy(),
    )
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle(refs[0].path, str(staging), refs),
        FakeConfigView(),
        FakeProgressSink(),
        FakeCancellationTokenView(),
        FakePluginLogger(),
        document_style_catalog=build_document_style_catalog(
            {"conversion": {"output_locale": "zh_CN"}},
            locales_dir=Path(__file__).resolve().parents[2] / "i18n" / "locales",
        ),
    )


def sha_text(value: str) -> str:
    return hashlib.sha256(value.encode()).hexdigest()


def write_caption_document(
    root: Path,
    source: str,
    *,
    target_end: int | None,
    kind: str = "table",
    title: str = "Cap",
    target_id: str | None = None,
    structural_tables: bool = True,
    enabled: bool = False,
) -> Path:
    targets = (
        ()
        if target_end is None
        else (ResolvedDocumentTarget(0, target_end, sha_text(source[:target_end]), kind, target_id, None, title),)
    )
    materialization = (
        CaptionMaterialization(
            "simple_seq",
            kind.title(),
            "arabic_half",
            "continue",
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            "1",
            kind.title(),
            " ",
        )
        if enabled
        else None
    )
    plan_targets = (
        ()
        if target_end is None
        else (NumberingTarget(0, target_end, kind, enabled, target_id, "1" if enabled else None, materialization),)
    )
    return write_resolved_document(root, source, targets, plan_targets, structural_tables=structural_tables)


def write_resolved_document(
    root: Path,
    source: str,
    targets: tuple[ResolvedDocumentTarget, ...],
    plan_targets: tuple[NumberingTarget, ...],
    *,
    structural_tables: bool = True,
) -> Path:
    plan_body = {"heading_definitions": [], "heading_instances": [], "targets": [asdict(item) for item in plan_targets]}
    identity = {
        "input_id": "source",
        "source_sha256": sha_text(source),
        "plan_sha256": hashlib.sha256(canonicalize_numbering_plan(plan_body)).hexdigest(),
    }
    neutral = {
        "$schema": "urn:docwen:schema:resolved-document:v1",
        "schema": "docwen.resolved_document.v1",
        **identity,
        "document": {
            "authored_markdown": source,
            "targets": [asdict(item) for item in targets],
            "references": [],
            "resource_occurrences": [],
            "citations": [],
            "resources": [],
        },
    }
    plan = {
        "$schema": "urn:docwen:schema:numbering-export-plan:v1",
        "schema": "docwen.numbering_export_plan.v1",
        **identity,
        "plan": plan_body,
    }
    root.mkdir(parents=True, exist_ok=True)
    paths = (root / "document.json", root / "plan.json")
    for path, value in zip(paths, (neutral, plan), strict=True):
        path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    original = tuple(path.read_bytes() for path in paths)
    refs = tuple(
        FileRef(
            path=str(path),
            format="markdown" if index == 0 else "json",
            category="document" if index == 0 else "resource",
            size_bytes=path.stat().st_size,
            input_kind="document" if index == 0 else "resource",
            input_role="neutral_document" if index == 0 else "numbering_export_plan",
            logical_path=path.name,
            media_type=RESOLVED_DOCUMENT_MEDIA_TYPE if index == 0 else NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
        )
        for index, path in enumerate(paths)
    )
    context = _context(
        root,
        refs,
        options={
            "locale": "zh_CN",
            "heading_merge_mode": "never",
            "markdown_extensions": {"input": {"structural_tables": structural_tables}},
        },
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    assert tuple(path.read_bytes() for path in paths) == original
    return Path(next(item.staging_path for item in result.artifacts if item.is_primary))


def read_document(root: Path, document: Path) -> tuple[Any, str]:
    refs = (FileRef(path=str(document), format="docx", category="document"),)
    initial = _context(root, refs, options={})
    initial.request.target_format = "md"
    initial.request.options = {
        "remove_numbering": False,
        "markdown_extensions": {"output": {"captions_references": True}},
    }
    context = FakeExecutionContext(
        initial.request,
        initial.workspace,
        initial.config,
        initial.progress,
        initial.cancellation,
        initial.logger,
        document_style_catalog=initial.document_style_catalog,
    )
    result = DocxToMarkdownConverter().convert(context)
    text = "\n".join(
        Path(item.staging_path).read_text(encoding="utf-8")
        for item in result.artifacts
        if item.media_type == "text/markdown"
    )
    return result, text
