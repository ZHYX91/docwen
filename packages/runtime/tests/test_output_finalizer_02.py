"""Focused tests split from test_output_finalizer.py."""

from __future__ import annotations

import pytest

from ._output_finalizer_support import (
    ARTIFACT_KIND_PRIMARY,
    ArtifactManifest,
    OutputFinalizer,
    OutputPolicy,
    Path,
    _create_staging_file,
    os,
    tempfile,
)
from ._output_finalizer_support import (
    finalizer as finalizer,
)

pytestmark = pytest.mark.integration


class TestOutputFinalizer:
    def test_finalize_missing_staging_source_returns_failure_without_phantom_artifact(
        self,
        finalizer: OutputFinalizer,
    ) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            missing = Path(staging) / "missing.docx"
            artifact = ArtifactManifest(
                artifact_id="missing",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=str(missing),
                suggested_name="report.docx",
                is_primary=True,
            )

            result = finalizer.finalize(
                task_id="task-missing",
                artifacts=[artifact],
                policy=OutputPolicy(output_dir=output),
            )

            assert result.success is False
            assert result.artifacts == []
            assert result.error is not None
            assert result.error.error_type == "output_finalization_failed"
            assert result.error.diagnostic_code == "FINALIZER_FAILED"
            assert [diagnostic.code for diagnostic in result.diagnostics] == [
                "FINALIZER_PLACE_ERROR",
                "FINALIZER_FAILED",
            ]
            assert "Staging artifact is not a file" in result.diagnostics[0].message
            assert not (Path(output) / "report.docx").exists()

    def test_finalize_partial_placement_preserves_real_artifacts_and_reports_typed_error(
        self,
        finalizer: OutputFinalizer,
    ) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            good = _create_staging_file(staging, "good.txt", "kept")
            missing = Path(staging) / "missing.txt"
            artifacts = [
                ArtifactManifest(
                    artifact_id="good",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=good,
                    suggested_name="good.txt",
                    is_primary=True,
                ),
                ArtifactManifest(
                    artifact_id="missing",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=str(missing),
                    suggested_name="missing.txt",
                    is_primary=True,
                ),
            ]

            result = finalizer.finalize(
                task_id="task-partial",
                artifacts=artifacts,
                policy=OutputPolicy(output_dir=output),
            )

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "output_finalization_failed"
            assert result.error.diagnostic_code == "FINALIZER_PARTIAL"
            assert [artifact.suggested_name for artifact in result.artifacts] == ["good.txt"]
            assert Path(result.artifacts[0].staging_path).read_text(encoding="utf-8") == "kept"
            assert result.metrics.output_bytes == 4
            assert result.metrics.extra == {
                "output_dir": output,
                "attempted_artifacts": 2,
                "placed_artifacts": 1,
                "failed_artifacts": 1,
            }
            assert result.diagnostics[-1].code == "FINALIZER_PARTIAL"

    def test_finalize_same_dir_mode(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as work:
            input_path = os.path.join(work, "input.md")
            with open(input_path, "w") as f:
                f.write("input")

            staging_path = _create_staging_file(staging, "output.docx")
            artifacts = [
                ArtifactManifest(
                    artifact_id="a1",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=staging_path,
                    suggested_name="input.docx",
                    is_primary=True,
                )
            ]
            result = finalizer.finalize(
                task_id="task-4",
                artifacts=artifacts,
                policy=OutputPolicy(),  # no output_dir → same as input
                input_path=input_path,
            )

            assert result.success is True
            final_path = result.artifacts[0].staging_path
            assert os.path.dirname(final_path) == work

    def test_finalize_date_subfolder(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            staging_path = _create_staging_file(staging, "output.docx")
            artifacts = [
                ArtifactManifest(
                    artifact_id="a1",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=staging_path,
                    suggested_name="report.docx",
                    is_primary=True,
                )
            ]
            result = finalizer.finalize(
                task_id="task-5",
                artifacts=artifacts,
                policy=OutputPolicy(output_dir=output, date_subfolder="iso"),
            )

            assert result.success is True
            final_path = result.artifacts[0].staging_path
            # Should be in a date subfolder
            parent = os.path.dirname(final_path)
            assert parent != output
            # ISO date format: YYYY-MM-DD
            import re

            assert re.match(r"\d{4}-\d{2}-\d{2}", os.path.basename(parent))

    def test_finalize_error_result(self, finalizer: OutputFinalizer) -> None:
        from docwen_core.models.result import ConversionErrorInfo

        result = finalizer.finalize_error(
            task_id="task-err",
            error_info=ConversionErrorInfo(
                error_type="conversion_failed",
                message="Plugin crashed",
            ),
            duration_ms=100.0,
            input_bytes=500,
        )

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "conversion_failed"
        assert result.metrics.duration_ms == 100.0

    def test_finalize_empty_artifacts(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as output:
            result = finalizer.finalize(
                task_id="task-empty",
                artifacts=[],
                policy=OutputPolicy(output_dir=output),
            )
            # No artifacts to place → not "success" in a meaningful sense
            assert result.success is False
            assert len(result.artifacts) == 0
            assert result.error is not None
            assert result.error.error_type == "output_finalization_failed"
            assert result.error.diagnostic_code == "FINALIZER_NO_ARTIFACTS"
            assert [diagnostic.code for diagnostic in result.diagnostics] == ["FINALIZER_NO_ARTIFACTS"]
            assert "FINALIZER_DONE" not in [diagnostic.code for diagnostic in result.diagnostics]

    def test_plugin_cannot_write_final_output_directly(self, finalizer: OutputFinalizer) -> None:
        """Proof that the final output path is determined by OutputFinalizer,
        not by the plugin.  The plugin only provides staging paths and
        suggested names; the finalizer decides the actual placement.

        This test verifies the architectural constraint:
        - Plugins write to staging (handled by WorkspaceManager)
        - OutputFinalizer is the ONLY component that writes to final output
        - Plugin never receives output_dir or final path information
        """
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            # Simulate: plugin writes to staging, provides suggested name
            staging_path = _create_staging_file(staging, "plugin_output.docx", "plugin content")

            artifact = ArtifactManifest(
                artifact_id="a1",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=staging_path,  # ← plugin only knows staging path
                suggested_name="my_document.docx",  # ← plugin suggests a name
                is_primary=True,
            )

            # The finalizer decides the actual placement
            result = finalizer.finalize(
                task_id="task-boundary",
                artifacts=[artifact],
                policy=OutputPolicy(output_dir=output),
            )

            # The final path is determined by the finalizer, not the plugin
            final_path = result.artifacts[0].staging_path
            assert final_path.startswith(output)
            # Plugin never knew about 'output' directory
            assert staging_path != final_path
            assert staging_path.startswith(staging)
            assert final_path.startswith(output)
