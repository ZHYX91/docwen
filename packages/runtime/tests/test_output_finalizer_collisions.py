"""Real filesystem collision progress and cancellation at publication retries."""

from __future__ import annotations

import errno
import os
from pathlib import Path

import pytest

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.result import ConversionResult
from docwen_core.models.task import TaskEvent
from docwen_core.protocols.execution_context import PluginExecutionContext
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

from ._fake_closed_loop_support import FakeClosedLoopPlugin
from ._output_finalizer_transaction_safety_support import _artifact

pytestmark = [pytest.mark.integration, pytest.mark.release_gate]


def _dangling_link(link: Path, target: Path) -> None:
    try:
        link.symlink_to(target)
    except OSError as error:
        if os.name == "nt" and getattr(error, "winerror", None) == 1314:
            pytest.skip("Creating a Windows file symlink requires Developer Mode or the symlink privilege")
        raise
    assert link.is_symlink()
    assert not link.exists()


@pytest.mark.parametrize("occupied_suffix", [False, True])
def test_rename_preserves_dangling_base_and_numbered_entries(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, occupied_suffix: bool
) -> None:
    staging = tmp_path / "source.xlsx"
    staging.write_bytes(b"complete new output")
    output = tmp_path / "output"
    output.mkdir()
    if occupied_suffix:
        (output / "report.xlsx").write_bytes(b"existing output")
    link = output / ("report_001.xlsx" if occupied_suffix else "report.xlsx")
    missing = output / "missing-target.xlsx"
    _dangling_link(link, missing)
    original_link = link.lstat()
    original_target = os.readlink(link)
    real_publish = OutputFinalizer._publish_no_clobber
    attempts: list[str] = []

    def bounded_publish(source: str, destination: str) -> None:
        attempts.append(Path(destination).name)
        assert len(attempts) <= 3, "publication retried an occupied entry without making progress"
        real_publish(source, destination)

    monkeypatch.setattr(OutputFinalizer, "_publish_no_clobber", staticmethod(bounded_publish))

    result = OutputFinalizer().finalize(
        "dangling-entry",
        [_artifact(staging, "report.xlsx")],
        OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
        input_path=str(staging),
    )

    expected = "report_002.xlsx" if occupied_suffix else "report_001.xlsx"
    assert result.success, result.diagnostics
    assert attempts == [expected]
    assert Path(result.artifacts[0].staging_path) == output / expected
    assert (output / expected).read_bytes() == b"complete new output"
    assert staging.read_bytes() == b"complete new output"
    assert (link.lstat().st_dev, link.lstat().st_ino) == (original_link.st_dev, original_link.st_ino)
    assert os.readlink(link) == original_target
    assert not missing.exists()
    if occupied_suffix:
        assert (output / "report.xlsx").read_bytes() == b"existing output"
    assert not list(output.glob(".__docwen-finalizer-*"))


def test_dangling_entry_created_after_prepare_is_skipped(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    staging = tmp_path / "source.xlsx"
    staging.write_bytes(b"new output")
    output = tmp_path / "output"
    real_publish = OutputFinalizer._publish_no_clobber
    attempts: list[str] = []
    missing = output / "missing-target.xlsx"
    original_link: list[tuple[str, int, int]] = []

    def publish_after_competing_link(source: str, destination: str) -> None:
        attempts.append(Path(destination).name)
        assert len(attempts) <= 3, "publication retried the competing link"
        if len(attempts) == 1:
            _dangling_link(Path(destination), missing)
            identity = Path(destination).lstat()
            original_link.append((os.readlink(destination), identity.st_dev, identity.st_ino))
        real_publish(source, destination)

    monkeypatch.setattr(OutputFinalizer, "_publish_no_clobber", staticmethod(publish_after_competing_link))
    result = OutputFinalizer().finalize(
        "late-dangling-entry",
        [_artifact(staging, "report.xlsx")],
        OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
    )

    assert result.success, result.diagnostics
    assert attempts == ["report.xlsx", "report_001.xlsx"]
    assert (output / "report.xlsx").is_symlink()
    # Windows readlink returns the substitution path with a \\?\ prefix.
    # Compare the original link bytes/identity instead of its display spelling.
    identity = (output / "report.xlsx").lstat()
    assert (os.readlink(output / "report.xlsx"), identity.st_dev, identity.st_ino) == original_link[0]
    assert not missing.exists()
    assert (output / "report_001.xlsx").read_bytes() == staging.read_bytes()
    assert not list(output.glob(".__docwen-finalizer-*"))


def test_transient_competitor_cannot_make_retries_reuse_a_rejected_name(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    staging = tmp_path / "source.xlsx"
    staging.write_bytes(b"complete output")
    output = tmp_path / "output"
    real_publish = OutputFinalizer._publish_no_clobber
    attempts: list[str] = []

    def publish_during_transient_competitor(source: str, destination: str) -> None:
        target = Path(destination)
        attempts.append(target.name)
        if len(attempts) <= 2:
            target.write_bytes(b"competing writer")
            try:
                real_publish(source, destination)
            finally:
                target.unlink()
        else:
            assert len(attempts) == 3
            real_publish(source, destination)

    monkeypatch.setattr(OutputFinalizer, "_publish_no_clobber", staticmethod(publish_during_transient_competitor))
    result = OutputFinalizer().finalize(
        "transient-collision",
        [_artifact(staging, "report.xlsx")],
        OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
    )

    assert result.success, result.diagnostics
    assert attempts == ["report.xlsx", "report_001.xlsx", "report_002.xlsx"]
    assert sorted(path.name for path in output.iterdir()) == ["report_002.xlsx"]
    assert (output / "report_002.xlsx").read_bytes() == staging.read_bytes()


def test_repeated_collisions_can_cancel_before_any_publication(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    staging = tmp_path / "source.xlsx"
    staging.write_bytes(b"must not be published")
    output = tmp_path / "output"
    token = CancellationToken()
    real_publish = OutputFinalizer._publish_no_clobber
    attempts: list[str] = []

    def publish_with_collisions(source: str, destination: str) -> None:
        target = Path(destination)
        attempts.append(target.name)
        assert len(attempts) <= 3, "cancelled retry dispatched another publication"
        target.write_bytes(b"competing output")
        if len(attempts) == 3:
            token.cancel("stop repeated collisions")
        real_publish(source, destination)

    monkeypatch.setattr(OutputFinalizer, "_publish_no_clobber", staticmethod(publish_with_collisions))
    with pytest.raises(CancellationRequested):
        OutputFinalizer().finalize(
            "cancel-retries",
            [_artifact(staging, "report.xlsx")],
            OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
            cancellation=token.view(),
        )

    assert attempts == ["report.xlsx", "report_001.xlsx", "report_002.xlsx"]
    assert sorted(path.name for path in output.iterdir()) == attempts
    assert all(path.read_bytes() == b"competing output" for path in output.iterdir())
    assert staging.read_bytes() == b"must not be published"


def test_cancelled_collision_after_one_commit_reports_kept_outputs_and_stops_rest(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    sources = [tmp_path / f"source-{index}.xlsx" for index in range(3)]
    for index, source in enumerate(sources):
        source.write_bytes(f"complete-{index}".encode())
    output = tmp_path / "output"
    token = CancellationToken()
    real_publish = OutputFinalizer._publish_no_clobber
    attempts: list[str] = []

    def publish_then_cancel_on_collision(source: str, destination: str) -> None:
        target = Path(destination)
        attempts.append(target.name)
        assert len(attempts) <= 2, "cancelled partial publication continued with another artifact"
        if target.name == "second.xlsx":
            target.write_bytes(b"external authoritative output")
            token.cancel("stop on collision after first output")
        real_publish(source, destination)

    artifacts = [
        _artifact(source, name)
        for source, name in zip(sources, ["first.xlsx", "second.xlsx", "third.xlsx"], strict=True)
    ]
    for index, artifact in enumerate(artifacts):
        artifact.artifact_id = f"artifact-{index}"
    monkeypatch.setattr(OutputFinalizer, "_publish_no_clobber", staticmethod(publish_then_cancel_on_collision))
    result = OutputFinalizer().finalize(
        "cancel-partial-retries",
        artifacts,
        OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
        cancellation=token.view(),
    )

    assert not result.success
    assert result.error is not None and result.error.diagnostic_code == "FINALIZER_PARTIAL"
    assert [artifact.artifact_id for artifact in result.artifacts] == ["artifact-0"]
    assert result.metrics.output_bytes == len(b"complete-0")
    assert result.metrics.extra["placed_artifacts"] == 1
    assert result.metrics.extra["failed_artifacts"] == 2
    assert any(item.code == "FINALIZER_RETRY_CANCELLED" for item in result.diagnostics)
    assert attempts == ["first.xlsx", "second.xlsx"]
    assert sorted(path.name for path in output.iterdir()) == ["first.xlsx", "second.xlsx"]
    assert (output / "first.xlsx").read_bytes() == sources[0].read_bytes()
    assert (output / "second.xlsx").read_bytes() == b"external authoritative output"


def test_name_scan_does_not_treat_inspection_error_as_available(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    staging = tmp_path / "source.xlsx"
    staging.write_bytes(b"new output")
    output = tmp_path / "output"
    output.mkdir()
    (output / "report.xlsx").write_bytes(b"existing output")
    real_lstat = Path.lstat

    def denied_candidate(path: Path):
        if path.name == "report_001.xlsx":
            raise PermissionError(errno.EACCES, "cannot inspect output entry", str(path))
        return real_lstat(path)

    monkeypatch.setattr(Path, "lstat", denied_candidate)
    result = OutputFinalizer().finalize(
        "uninspectable-candidate",
        [_artifact(staging, "report.xlsx")],
        OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
    )

    assert not result.success
    assert result.artifacts == []
    assert any("cannot inspect output entry" in item.message for item in result.diagnostics)
    assert sorted(path.name for path in output.iterdir()) == ["report.xlsx"]
    assert (output / "report.xlsx").read_bytes() == b"existing output"


def test_task_manager_keeps_partial_outputs_when_a_later_publication_retry_is_cancelled(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    class ThreeArtifactPlugin(FakeClosedLoopPlugin):
        def convert(self, context: PluginExecutionContext) -> ConversionResult:
            artifacts: list[ArtifactManifest] = []
            for index, name in enumerate(("first.pdf", "second.pdf", "third.pdf")):
                path = Path(context.workspace.create_artifact_path(f"artifact-{index}", ".pdf"))
                path.write_bytes(f"complete-{index}".encode())
                artifact = _artifact(path, name)
                artifact.artifact_id = f"artifact-{index}"
                context.workspace.add_artifact(artifact)
                artifacts.append(artifact)
            return ConversionResult(task_id=context.request.request_id, success=True, artifacts=artifacts)

    registry = PluginRegistry()
    registry.register(ThreeArtifactPlugin("three_artifacts", "txt", "pdf"))
    workspaces = WorkspaceManager(root_dir=str(tmp_path / "workspace"))
    manager = TaskManager(registry, RouteResolver(registry), workspaces, OutputFinalizer())
    input_path = tmp_path / "source.txt"
    input_path.write_text("source content", encoding="utf-8")
    output = tmp_path / "output"
    request = ConversionRequest(
        request_id="manager-partial-cancel",
        input_refs=[FileRef(path=str(input_path), format="txt", category="text")],
        target_format="pdf",
        output_policy=OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
    )
    events: list[TaskEvent] = []
    attempts: list[str] = []
    real_publish = OutputFinalizer._publish_no_clobber

    def publish_then_cancel_on_collision(source: str, destination: str) -> None:
        target = Path(destination)
        if target.parent == output:
            attempts.append(target.name)
            assert len(attempts) <= 2, "runtime continued publication after the cancelled retry"
            if target.name == "second.pdf":
                target.write_bytes(b"external authoritative output")
                manager.cancel(request.request_id)
        real_publish(source, destination)

    monkeypatch.setattr(OutputFinalizer, "_publish_no_clobber", staticmethod(publish_then_cancel_on_collision))
    try:
        result = manager.execute_single(request, on_event=events.append)
    finally:
        workspaces.cleanup_all()

    assert not result.success
    assert result.error is not None and result.error.diagnostic_code == "FINALIZER_PARTIAL"
    assert result.error.error_type == "output_finalization_failed"
    assert [artifact.artifact_id for artifact in result.artifacts] == ["artifact-0"]
    assert Path(result.artifacts[0].staging_path) == output / "first.pdf"
    assert result.metrics.output_bytes == len(b"complete-0")
    assert result.metrics.extra["placed_artifacts"] == 1
    assert result.metrics.extra["failed_artifacts"] == 2
    assert [event.event_type for event in events][-1] == "task_failed"
    assert not {"task_cancelled", "task_completed"}.intersection(event.event_type for event in events)
    assert attempts == ["first.pdf", "second.pdf"]
    assert sorted(path.name for path in output.iterdir()) == ["first.pdf", "second.pdf"]
    assert (output / "first.pdf").read_bytes() == b"complete-0"
    assert (output / "second.pdf").read_bytes() == b"external authoritative output"
    assert input_path.read_text(encoding="utf-8") == "source content"
