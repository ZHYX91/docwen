"""Office bridge results preserve stage diagnostics and opt-in intermediates."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import cast

import pytest

from docwen_core.models.artifact import ARTIFACT_KIND_AUXILIARY, ArtifactManifest
from docwen_core.models.result import ConversionDiagnostic, ConversionMetrics, ConversionResult
from docwen_core.protocols.execution_context import PluginExecutionContext
from docwen_plugin_markdown.office_bridge import converter as bridge_module

pytestmark = pytest.mark.unit


class _Progress:
    def __init__(self) -> None:
        self.values: list[float] = []

    def report_progress(self, value: float, message: str = "") -> None:
        self.values.append(value)


class _Workspace:
    def __init__(self, root: Path) -> None:
        self.root = root
        self.input_path = str(root / "input.md")
        Path(self.input_path).write_text("Input", encoding="utf-8")
        self.added: list[ArtifactManifest] = []

    def create_artifact_path(self, name: str, suffix: str) -> str:
        return str(self.root / f"{name}{suffix}")

    def add_artifact(self, artifact: ArtifactManifest) -> None:
        self.added.append(artifact)


@pytest.mark.parametrize("save_intermediate", [False, True])
def test_bridge_propagates_stage_warnings_and_retains_only_requested_artifact(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, save_intermediate: bool
) -> None:
    progress = _Progress()
    workspace = _Workspace(tmp_path)
    docx = tmp_path / "input.docx"
    docx.write_bytes(b"DOCX-INTERMEDIATE")
    upstream = ArtifactManifest(
        artifact_id="task-docx",
        kind="primary",
        staging_path=str(docx),
        suggested_name="input.docx",
        is_primary=True,
    )
    warning = ConversionDiagnostic(level="warning", message="Fallback preserved as code", code="MERMAID-FALLBACK")
    stage = ConversionResult(
        task_id="task",
        success=True,
        artifacts=[upstream],
        diagnostics=[warning],
        metrics=ConversionMetrics(),
    )

    def write_target(_source: str, destination: str, **_kwargs):
        Path(destination).write_bytes(b"RTF-RESULT")
        return SimpleNamespace(success=True, backend="test", message="")

    monkeypatch.setattr(bridge_module, "convert_with_backend_priority", write_target)
    ctx = SimpleNamespace(
        request=SimpleNamespace(request_id="task"),
        workspace=workspace,
        progress=progress,
        cancellation=None,
        config={"output.intermediate_files.save_to_output": save_intermediate},
    )
    converter = bridge_module.MarkdownOfficeBridgeConverter()

    def intermediate(stage_context):
        stage_context.progress.report_progress(100, "Intermediate finished")
        return stage

    result = converter._convert_with_intermediate(cast(PluginExecutionContext, ctx), "rtf", intermediate, "docx")
    assert result.success
    assert "MERMAID-FALLBACK" in {item.code for item in result.diagnostics}
    assert [item.kind for item in result.artifacts] == (
        ["primary", ARTIFACT_KIND_AUXILIARY] if save_intermediate else ["primary"]
    )
    assert progress.values == [70.0, 72.0, 85.0]
    assert result.artifacts[0].suggested_name == "input.rtf"
    if save_intermediate:
        assert result.artifacts[1].suggested_name == "input.docx"


def test_bridge_failure_keeps_upstream_warning(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workspace = _Workspace(tmp_path)
    docx = tmp_path / "source.docx"
    docx.write_bytes(b"data")
    stage = ConversionResult(
        task_id="task",
        success=True,
        artifacts=[
            ArtifactManifest(
                artifact_id="intermediate",
                kind="primary",
                staging_path=str(docx),
                suggested_name="source.docx",
                is_primary=True,
            )
        ],
        diagnostics=[ConversionDiagnostic(level="warning", message="style conflict", code="STYLE-WARN")],
    )
    monkeypatch.setattr(
        bridge_module,
        "convert_with_backend_priority",
        lambda *_args, **_kwargs: SimpleNamespace(success=False, backend=None, message="office failed"),
    )
    ctx = SimpleNamespace(
        request=SimpleNamespace(request_id="task"),
        workspace=workspace,
        progress=_Progress(),
        cancellation=None,
        config={},
    )
    result = bridge_module.MarkdownOfficeBridgeConverter()._convert_with_intermediate(
        cast(PluginExecutionContext, ctx), "rtf", lambda _stage_context: stage, "docx"
    )
    assert not result.success
    assert "STYLE-WARN" in {item.code for item in result.diagnostics}
