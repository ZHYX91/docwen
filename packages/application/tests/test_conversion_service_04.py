"""Focused tests split from test_conversion_service.py."""

from __future__ import annotations

from ._conversion_service_support import (
    Any,
    ArtifactBundle,
    ArtifactManifest,
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionResult,
    ConversionService,
    ConversionServiceError,
    Path,
    _Committer,
    _Controller,
    _request,
    pytest,
)

pytestmark = pytest.mark.integration


def test_failed_runtime_result_never_creates_bundle(tmp_path: Path) -> None:
    class _FailedController(_Controller):
        def execute_single(self, request: Any) -> ConversionResult:
            rejected = Path(request.output_policy.output_dir) / "rejected.docx"
            rejected.write_bytes(b"rejected")
            return ConversionResult(
                task_id=request.request_id,
                success=False,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="artifact.rejected",
                        kind="primary",
                        staging_path=str(rejected),
                        suggested_name=rejected.name,
                    )
                ],
                error=ConversionErrorInfo(error_type="conversion_failed", message="failed"),
            )

    source = tmp_path / "source.md"
    source.write_text("body", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_FailedController(), _Committer())
    task_id = service.accept(service.plan(_request(source, staging)).plan_id, "task.failed")

    outcome = service.execute_accepted(task_id)
    assert outcome.state == "failed"
    assert outcome.bundle is None
    assert list(staging.iterdir()) == []


def test_failed_runtime_result_rejects_artifact_bound_diagnostic(tmp_path: Path) -> None:
    class _FailedController(_Controller):
        def execute_single(self, request: Any) -> ConversionResult:
            rejected = Path(request.output_policy.output_dir) / "rejected.docx"
            rejected.write_bytes(b"rejected")
            return ConversionResult(
                task_id=request.request_id,
                success=False,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="artifact.rejected",
                        kind="primary",
                        staging_path=str(rejected),
                        suggested_name=rejected.name,
                    )
                ],
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message="page failed",
                        code="recognition_failed",
                        artifact_id="artifact.rejected",
                    )
                ],
                error=ConversionErrorInfo(error_type="conversion_failed", message="failed"),
            )

    source = tmp_path / "source.md"
    source.write_text("body", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_FailedController(), _Committer())
    task_id = service.accept(service.plan(_request(source, staging)).plan_id, "task.failed-bound")

    with pytest.raises(ConversionServiceError) as rejected:
        service.execute_accepted(task_id)

    assert rejected.value.code == "dangling_diagnostic_artifact"
    assert list(staging.iterdir()) == []


def test_bundle_commit_failure_discards_rejected_output(tmp_path: Path) -> None:
    class _RejectingCommitter(_Committer):
        def commit(self, *, task_id: str, staging_root: str, draft: Any) -> ArtifactBundle:
            raise ValueError("invalid bundle")

    source = tmp_path / "source.md"
    source.write_text("body", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _RejectingCommitter())
    task_id = service.accept(service.plan(_request(source, staging)).plan_id, "task.rejected")

    with pytest.raises(ConversionServiceError) as rejected:
        service.execute_accepted(task_id)
    assert rejected.value.code == "bundle_commit_failed"
    assert list(staging.iterdir()) == []


def test_dangling_diagnostic_artifact_is_rejected_before_bundle_commit(tmp_path: Path) -> None:
    class _DiagnosticController(_Controller):
        def execute_single(self, request: Any) -> ConversionResult:
            result = super().execute_single(request)
            result.diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message="unknown page",
                    code="resource_page_unresolved",
                    artifact_id="resource.missing",
                )
            )
            return result

    class _ObservingCommitter(_Committer):
        called = False

        def commit(self, *, task_id: str, staging_root: str, draft: Any) -> ArtifactBundle:
            self.called = True
            return super().commit(task_id=task_id, staging_root=staging_root, draft=draft)

    source = tmp_path / "source.md"
    source.write_text("body", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    committer = _ObservingCommitter()
    service = ConversionService(_DiagnosticController(), committer)
    task_id = service.accept(service.plan(_request(source, staging)).plan_id, "task.diagnostic")

    with pytest.raises(ConversionServiceError) as rejected:
        service.execute_accepted(task_id)

    assert rejected.value.code == "dangling_diagnostic_artifact"
    assert committer.called is False
    assert list(staging.iterdir()) == []
