"""Application-owned Markdown→DOCX proofreading pipeline contracts."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest

from docwen_application.controller import ApplicationController
from docwen_application.ports.runtime import RuntimePort
from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import (
    POSTPROCESS_PROOFREAD_OPTION,
    ConversionRequest,
    OutputPolicy,
)
from docwen_core.models.result import ConversionDiagnostic, ConversionErrorInfo, ConversionMetrics, ConversionResult

pytestmark = pytest.mark.unit

_DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@pytest.mark.parametrize("failure", ["render", "proofread", "cancel_render", "cancel_proofread"])
def test_failure_or_cancellation_never_publishes_private_render(tmp_path, monkeypatch, failure):
    monkeypatch.setattr("docwen_core.detection.enforce_file_admission", lambda request: request)
    source = tmp_path / "note.md"
    source.write_text("# 原文\n公文才料（", encoding="utf-8")
    original = source.read_bytes()
    output = tmp_path / "public"
    runtime = MagicMock(spec=RuntimePort)
    controller = ApplicationController(runtime_port=runtime)
    private_paths = []

    def execute(stage):
        if not stage.action_name:
            private = Path(stage.output_policy.output_dir)
            private_paths.append(private)
            docx = private / "note.docx"
            docx.write_bytes(b"private")
            if failure == "render":
                return ConversionResult(
                    task_id=stage.request_id,
                    success=False,
                    error=ConversionErrorInfo(error_type="conversion_failed", message="render failure"),
                )
            if failure == "cancel_render":
                controller.cancel("pipeline-failure")
            return ConversionResult(
                task_id=stage.request_id,
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="render",
                        kind="primary",
                        staging_path=str(docx),
                        suggested_name="note.docx",
                        media_type=_DOCX_MEDIA_TYPE,
                        is_primary=True,
                    )
                ],
            )
        if failure == "cancel_proofread":
            controller.cancel("pipeline-failure")
        return ConversionResult(
            task_id=stage.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="cancelled" if failure.startswith("cancel") else "conversion_failed",
                message="proofread stopped",
            ),
        )

    runtime.execute.side_effect = execute
    result = controller.execute_single(
        ConversionRequest(
            request_id="pipeline-failure",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="docx",
            options={POSTPROCESS_PROOFREAD_OPTION: {"enable_typos_rule": True}},
            output_policy=OutputPolicy(output_dir=str(output)),
        )
    )
    assert not result.success and not result.artifacts
    assert result.error.error_type == ("cancelled" if failure.startswith("cancel") else "conversion_failed")
    assert runtime.execute.call_count == (1 if failure in {"render", "cancel_render"} else 2)
    assert not output.exists()
    assert private_paths and all(not path.exists() for path in private_paths)
    assert source.read_bytes() == original


def test_markdown_docx_proofread_is_two_private_then_public_runtime_stages(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setattr("docwen_core.detection.enforce_file_admission", lambda request: request)
    source = tmp_path / "note.md"
    source.write_text("# 标题\n正文", encoding="utf-8")
    published = tmp_path / "published"
    final_docx = published / "note_checked.docx"

    runtime = MagicMock(spec=RuntimePort)
    calls: list[ConversionRequest] = []

    def execute(request: ConversionRequest) -> ConversionResult:
        calls.append(request)
        if request.action_name == "":
            private_docx = Path(request.output_policy.output_dir or "") / "rendered.docx"
            private_docx.parent.mkdir(parents=True, exist_ok=True)
            private_docx.write_bytes(b"private-docx")
            return ConversionResult(
                task_id=request.request_id,
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="rendered",
                        kind="primary",
                        staging_path=str(private_docx),
                        suggested_name="note.docx",
                        media_type=_DOCX_MEDIA_TYPE,
                        size_bytes=private_docx.stat().st_size,
                        is_primary=True,
                    )
                ],
                diagnostics=[
                    ConversionDiagnostic(level="info", message="rendered", code="MD2DOCX-OK"),
                    ConversionDiagnostic(
                        level="info",
                        message=f"Published private node {private_docx.parent}",
                        code="FINALIZER_DONE",
                    ),
                ],
                metrics=ConversionMetrics(duration_ms=10.0, input_bytes=source.stat().st_size),
            )

        assert request.action_name == "validate"
        final_docx.parent.mkdir(parents=True, exist_ok=True)
        final_docx.write_bytes(b"proofread-docx")
        return ConversionResult(
            task_id=request.request_id,
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id="checked",
                    kind="primary",
                    staging_path=str(final_docx),
                    suggested_name=final_docx.name,
                    media_type=_DOCX_MEDIA_TYPE,
                    size_bytes=final_docx.stat().st_size,
                    is_primary=True,
                )
            ],
            diagnostics=[ConversionDiagnostic(level="info", message="checked", code="PROOFREAD-OK")],
            metrics=ConversionMetrics(duration_ms=5.0, input_bytes=12, output_bytes=final_docx.stat().st_size),
        )

    runtime.execute.side_effect = execute
    controller = ApplicationController(runtime_port=runtime)
    request = ConversionRequest(
        request_id="pipeline",
        input_refs=[
            FileRef(
                path=str(source),
                format="markdown",
                category="markdown",
                size_bytes=source.stat().st_size,
                logical_path="note.md",
            )
        ],
        target_format="docx",
        options={
            "remove_numbering": True,
            POSTPROCESS_PROOFREAD_OPTION: {
                "enable_symbol_pairing": True,
                "enable_symbol_correction": False,
                "enable_typos_rule": True,
                "enable_sensitive_word": False,
            },
        },
        output_policy=OutputPolicy(output_dir=str(published), overwrite_mode="rename"),
    )

    result = controller.execute_single(request)

    assert result.success is True
    assert [item.code for item in result.diagnostics] == [
        "MD2DOCX-OK",
        "PROOFREAD-OK",
        "POSTPROCESS_PROOFREAD_APPLIED",
    ]
    assert result.task_id == "pipeline"
    assert result.metrics.duration_ms == 15.0
    assert result.metrics.input_bytes == source.stat().st_size
    assert result.metrics.extra["proofread_postprocess"] is True
    assert result.artifacts[0].staging_path == str(final_docx)

    assert len(calls) == 2
    render, proofread = calls
    assert render.request_id == "pipeline-render"
    assert render.action_name == ""
    assert POSTPROCESS_PROOFREAD_OPTION not in render.options
    assert render.output_policy.output_dir != str(published)
    assert render.output_policy.group_outputs is True
    assert render.conversion_identity is not None
    private_output = render.output_policy.output_dir
    assert private_output is not None

    assert proofread.request_id == "pipeline"
    assert proofread.action_name == "validate"
    assert proofread.target_format == "docx"
    assert proofread.options == {
        "enable_symbol_pairing": True,
        "enable_symbol_correction": False,
        "enable_typos_rule": True,
        "enable_sensitive_word": False,
    }
    assert proofread.output_policy.output_dir == str(published)
    assert proofread.output_policy.group_outputs is True
    assert render.conversion_identity is not None
    assert proofread.conversion_identity is not None
    assert render.conversion_identity.task_id == "pipeline-render"
    assert proofread.conversion_identity.task_id == "pipeline"
    assert (
        render.conversion_identity.source_stem,
        render.conversion_identity.source_format,
        render.conversion_identity.created_at,
    ) == (
        proofread.conversion_identity.source_stem,
        proofread.conversion_identity.source_format,
        proofread.conversion_identity.created_at,
    )
    assert proofread.input_refs[0].path.endswith("rendered.docx")
    assert proofread.input_refs[0].format == "docx"
    assert not Path(private_output).exists()


def test_postprocess_proofread_rejects_non_docx_target(tmp_path: Path) -> None:
    source = tmp_path / "note.md"
    source.write_text("body", encoding="utf-8")
    controller = ApplicationController(runtime_port=MagicMock(spec=RuntimePort))
    request = ConversionRequest(
        request_id="wrong-target",
        input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
        target_format="pdf",
        options={POSTPROCESS_PROOFREAD_OPTION: {"enable_symbol_pairing": True}},
    )

    with pytest.raises(ValueError, match="only for ordinary DOCX"):
        controller.execute_single(request)
