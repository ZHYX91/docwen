"""Focused tests split from test_output_finalizer.py."""

from __future__ import annotations

from ._output_finalizer_support import (
    ARTIFACT_KIND_PRIMARY,
    ArtifactManifest,
    OutputFinalizer,
    OutputPolicy,
    Path,
    ThreadPoolExecutor,
    _create_staging_file,
    os,
    pytest,
    tempfile,
    threading,
)
from ._output_finalizer_support import (
    finalizer as finalizer,
)

pytestmark = pytest.mark.integration


class TestOutputFinalizer:
    def test_finalize_exact_primary_output_path(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            staging_path = _create_staging_file(staging, "generated.docx")
            exact_path = os.path.join(output, "chosen-name.docx")
            artifacts = [
                ArtifactManifest(
                    artifact_id="primary",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=staging_path,
                    suggested_name="plugin-name.docx",
                    is_primary=True,
                )
            ]

            result = finalizer.finalize(
                task_id="exact-output",
                artifacts=artifacts,
                policy=OutputPolicy(output_path=exact_path, overwrite_mode="error"),
            )

            assert result.success is True
            assert result.artifacts[0].staging_path == exact_path
            assert Path(exact_path).read_text() == "test content"

    def test_finalize_exact_output_refuses_existing_target(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            staging_path = _create_staging_file(staging, "generated.docx", "new")
            exact_path = Path(output) / "existing.docx"
            exact_path.write_text("original")
            artifacts = [
                ArtifactManifest(
                    artifact_id="primary",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=staging_path,
                    suggested_name="plugin-name.docx",
                    is_primary=True,
                )
            ]

            result = finalizer.finalize(
                task_id="exact-conflict",
                artifacts=artifacts,
                policy=OutputPolicy(output_path=str(exact_path), overwrite_mode="error"),
            )

            assert result.success is False
            assert exact_path.read_text() == "original"
            assert result.artifacts == []

    def test_finalize_single_artifact(self, finalizer: OutputFinalizer) -> None:
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
                task_id="task-1",
                artifacts=artifacts,
                policy=OutputPolicy(output_dir=output),
            )

            assert result.success is True
            assert len(result.artifacts) == 1
            final_path = result.artifacts[0].staging_path
            assert os.path.isfile(final_path)
            assert final_path.startswith(output)
            assert os.path.basename(final_path) == "report.docx"

    def test_finalize_respects_rename_policy(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            # Pre-create a file at the target path
            existing = os.path.join(output, "report.docx")
            with open(existing, "w") as f:
                f.write("existing")

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
                task_id="task-1",
                artifacts=artifacts,
                policy=OutputPolicy(output_dir=output, overwrite_mode="rename"),
            )

            assert result.success is True
            final_name = os.path.basename(result.artifacts[0].staging_path)
            assert final_name == "report_001.docx"

    @pytest.mark.parametrize("overwrite_mode", ["error", "rename", "overwrite", "skip"])
    def test_finalize_copies_identical_nonprimary_input_into_document_node(
        self,
        finalizer: OutputFinalizer,
        overwrite_mode: str,
    ) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as work:
            input_path = Path(work) / "source.png"
            input_path.write_bytes(b"same retained image bytes")
            primary_path = Path(staging) / "source.md"
            primary_path.write_text("![[source.png]]", encoding="utf-8")
            retained_path = Path(staging) / "source.png"
            retained_path.write_bytes(input_path.read_bytes())
            artifacts = [
                ArtifactManifest(
                    artifact_id="primary",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=str(primary_path),
                    suggested_name="source.md",
                    media_type="text/markdown",
                    is_primary=True,
                ),
                ArtifactManifest(
                    artifact_id="retained",
                    kind="image",
                    staging_path=str(retained_path),
                    suggested_name="source.png",
                    media_type="image/png",
                    is_primary=False,
                ),
            ]

            result = finalizer.finalize(
                task_id="same-dir-retained-input",
                artifacts=artifacts,
                policy=OutputPolicy(overwrite_mode=overwrite_mode),
                input_path=str(input_path),
            )

            assert result.success is True
            retained = next(artifact for artifact in result.artifacts if artifact.kind == "image")
            assert Path(retained.staging_path).parent == Path(result.metrics.extra["document_node_root"])
            assert Path(retained.staging_path).read_bytes() == input_path.read_bytes()
            assert retained.metadata["document_node_committed"] is True
            assert not (Path(work) / "source_001.png").exists()
            main = next(artifact for artifact in result.artifacts if artifact.is_primary)
            assert Path(main.staging_path).read_text(encoding="utf-8") == "![[source.png]]"

    def test_finalize_changed_retained_input_is_isolated_inside_document_node(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as work:
            input_path = Path(work) / "source.png"
            input_path.write_bytes(b"original input")
            primary_path = Path(staging) / "source.md"
            primary_path.write_text("![[source.png]]", encoding="utf-8")
            retained_path = Path(staging) / "source.png"
            retained_path.write_bytes(b"changed retained artifact")
            artifacts = [
                ArtifactManifest(
                    artifact_id="primary",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=str(primary_path),
                    suggested_name="source.md",
                    media_type="text/markdown",
                    is_primary=True,
                ),
                ArtifactManifest(
                    artifact_id="retained",
                    kind="image",
                    staging_path=str(retained_path),
                    suggested_name="source.png",
                    media_type="image/png",
                    is_primary=False,
                ),
            ]

            result = finalizer.finalize(
                task_id="same-dir-changed-retained-input",
                artifacts=artifacts,
                policy=OutputPolicy(overwrite_mode="rename"),
                input_path=str(input_path),
            )

            assert result.success is True
            retained = next(artifact for artifact in result.artifacts if artifact.artifact_id == "retained")
            assert Path(retained.staging_path).read_bytes() == b"changed retained artifact"
            assert input_path.read_bytes() == b"original input"
            assert not (Path(work) / "source_001.png").exists()

    def test_finalize_serializes_same_output_directory_across_instances(self, monkeypatch) -> None:
        """Collision detection and placement must be atomic for one output directory."""
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            staging_paths = [
                _create_staging_file(staging, f"source-{index}.md", f"content-{index}") for index in range(2)
            ]
            original_prepare_artifact = OutputFinalizer._prepare_artifact
            state_lock = threading.Lock()
            start_barrier = threading.Barrier(2)
            overlapping_call = threading.Event()
            active_calls = 0
            call_count = 0
            max_active_calls = 0

            def tracked_prepare_artifact(
                artifact,
                output_dir,
                overwrite_mode,
                input_path,
                cancellation,
            ):
                nonlocal active_calls, call_count, max_active_calls
                with state_lock:
                    active_calls += 1
                    call_count += 1
                    current_call = call_count
                    max_active_calls = max(max_active_calls, active_calls)
                    if active_calls > 1:
                        overlapping_call.set()
                if current_call == 1:
                    overlapping_call.wait(0.5)
                try:
                    return original_prepare_artifact(
                        artifact,
                        output_dir,
                        overwrite_mode,
                        input_path,
                        cancellation,
                    )
                finally:
                    with state_lock:
                        active_calls -= 1

            monkeypatch.setattr(
                OutputFinalizer,
                "_prepare_artifact",
                staticmethod(tracked_prepare_artifact),
            )

            def finalize(index: int):
                start_barrier.wait()
                artifact = ArtifactManifest(
                    artifact_id=f"a{index}",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=staging_paths[index],
                    suggested_name="report.md",
                    is_primary=True,
                )
                return OutputFinalizer().finalize(
                    task_id=f"task-{index}",
                    artifacts=[artifact],
                    policy=OutputPolicy(output_dir=output, overwrite_mode="rename"),
                )

            with ThreadPoolExecutor(max_workers=2) as executor:
                results = list(executor.map(finalize, range(2)))

            assert all(result.success for result in results)
            assert max_active_calls == 1
            assert sorted(os.path.basename(result.artifacts[0].staging_path) for result in results) == [
                "report.md",
                "report_001.md",
            ]
            assert {Path(result.artifacts[0].staging_path).read_text(encoding="utf-8") for result in results} == {
                "content-0",
                "content-1",
            }

    def test_finalize_overwrite_mode(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            existing = os.path.join(output, "report.docx")
            with open(existing, "w") as f:
                f.write("old content")

            staging_path = _create_staging_file(staging, "output.docx", "new content")
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
                task_id="task-2",
                artifacts=artifacts,
                policy=OutputPolicy(output_dir=output, overwrite_mode="overwrite"),
            )

            assert result.success is True
            with open(existing) as f:
                assert f.read() == "new content"

    def test_finalize_skip_mode(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            existing = os.path.join(output, "report.docx")
            with open(existing, "w") as f:
                f.write("old content")

            staging_path = _create_staging_file(staging, "output.docx", "new content")
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
                task_id="task-3",
                artifacts=artifacts,
                policy=OutputPolicy(output_dir=output, overwrite_mode="skip"),
            )

            assert result.success is True
            # The file content should NOT change
            with open(existing) as f:
                assert f.read() == "old content"
            # The artifact should be marked as skipped
            assert result.artifacts[0].metadata.get("skipped") is True

    def test_finalize_skip_existing_target_remains_a_real_skipped_artifact(
        self,
        finalizer: OutputFinalizer,
    ) -> None:
        """Skip may reuse a real destination even when staging is already gone."""
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            existing = Path(output) / "report.docx"
            existing.write_text("existing", encoding="utf-8")
            missing_staging = Path(staging) / "missing.docx"
            artifact = ArtifactManifest(
                artifact_id="a1",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=str(missing_staging),
                suggested_name=existing.name,
                is_primary=True,
            )

            result = finalizer.finalize(
                task_id="task-skip-existing",
                artifacts=[artifact],
                policy=OutputPolicy(output_dir=output, overwrite_mode="skip"),
            )

            assert result.success is True
            assert result.error is None
            assert result.artifacts[0].staging_path == str(existing)
            assert result.artifacts[0].metadata["skipped"] is True

    def test_finalize_skip_rejects_existing_directory_as_artifact(
        self,
        finalizer: OutputFinalizer,
    ) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            existing_directory = Path(output) / "report.docx"
            existing_directory.mkdir()
            artifact = ArtifactManifest(
                artifact_id="a1",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=str(Path(staging) / "missing.docx"),
                suggested_name=existing_directory.name,
                is_primary=True,
            )

            result = finalizer.finalize(
                task_id="task-skip-directory",
                artifacts=[artifact],
                policy=OutputPolicy(output_dir=output, overwrite_mode="skip"),
            )

            assert result.success is False
            assert result.artifacts == []
            assert result.error is not None
            assert result.error.error_type == "output_finalization_failed"
            assert result.error.diagnostic_code == "FINALIZER_FAILED"
            assert result.diagnostics[0].code == "FINALIZER_PLACE_ERROR"
            assert "not a file" in result.diagnostics[0].message

    def test_finalize_overwrite_rejects_existing_directory_as_artifact(
        self,
        finalizer: OutputFinalizer,
    ) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            staging_path = _create_staging_file(staging, "staging.docx", "new content")
            existing_directory = Path(output) / "report.docx"
            existing_directory.mkdir()
            artifact = ArtifactManifest(
                artifact_id="a1",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=staging_path,
                suggested_name=existing_directory.name,
                is_primary=True,
            )

            result = finalizer.finalize(
                task_id="task-overwrite-directory",
                artifacts=[artifact],
                policy=OutputPolicy(output_dir=output, overwrite_mode="overwrite"),
            )

            assert result.success is False
            assert result.artifacts == []
            assert result.error is not None
            assert result.error.error_type == "output_finalization_failed"
            assert result.error.diagnostic_code == "FINALIZER_FAILED"
            assert result.diagnostics[0].code == "FINALIZER_PLACE_ERROR"
            assert "not a file" in result.diagnostics[0].message
            assert list(existing_directory.iterdir()) == []

    def test_finalize_rejects_unknown_overwrite_mode_without_mutating_existing_output(
        self,
        finalizer: OutputFinalizer,
    ) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            staging_path = _create_staging_file(staging, "staging.docx", "new content")
            existing = Path(output) / "report.docx"
            existing.write_text("existing content", encoding="utf-8")
            artifact = ArtifactManifest(
                artifact_id="a1",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=staging_path,
                suggested_name=existing.name,
                is_primary=True,
            )

            result = finalizer.finalize(
                task_id="task-invalid-overwrite-mode",
                artifacts=[artifact],
                policy=OutputPolicy(output_dir=output, overwrite_mode="invalid"),
            )

            assert result.success is False
            assert result.artifacts == []
            assert result.error is not None
            assert result.error.error_type == "output_finalization_failed"
            assert result.error.diagnostic_code == "FINALIZER_FAILED"
            assert result.diagnostics[0].code == "FINALIZER_PLACE_ERROR"
            assert "Unknown overwrite mode" in result.diagnostics[0].message
            assert existing.read_text(encoding="utf-8") == "existing content"

    def test_finalize_allows_safe_relative_subpath(self, finalizer: OutputFinalizer) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            staging_path = _create_staging_file(staging, "sheet.csv", "a,b\n1,2\n")
            artifact = ArtifactManifest(
                artifact_id="a1",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=staging_path,
                suggested_name=os.path.join("report_fromMd", "report_Sheet1_fromMd.csv"),
                is_primary=True,
            )

            result = finalizer.finalize(
                task_id="task-subpath",
                artifacts=[artifact],
                policy=OutputPolicy(output_dir=output),
            )

            assert result.success is True
            final_path = result.artifacts[0].staging_path
            assert os.path.isfile(final_path)
            assert os.path.dirname(final_path) == os.path.join(output, "report_fromMd")
            assert result.artifacts[0].suggested_name == os.path.join("report_fromMd", "report_Sheet1_fromMd.csv")

    @pytest.mark.parametrize(
        "suggested_name",
        [
            os.path.join("..", "escape.csv"),
            os.path.abspath("escape.csv"),
            ".",
        ],
    )
    def test_finalize_rejects_unsafe_suggested_name(self, finalizer: OutputFinalizer, suggested_name: str) -> None:
        with tempfile.TemporaryDirectory() as staging, tempfile.TemporaryDirectory() as output:
            staging_path = _create_staging_file(staging, "sheet.csv", "a,b\n1,2\n")
            artifact = ArtifactManifest(
                artifact_id="a1",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=staging_path,
                suggested_name=suggested_name,
                is_primary=True,
            )

            result = finalizer.finalize(
                task_id="task-unsafe",
                artifacts=[artifact],
                policy=OutputPolicy(output_dir=output),
            )

            assert result.success is False
            assert result.artifacts == []
            assert result.error is not None
            assert result.error.error_type == "output_finalization_failed"
            assert result.error.diagnostic_code == "FINALIZER_FAILED"
            assert result.diagnostics[0].code == "FINALIZER_PLACE_ERROR"
            assert "Unsafe artifact suggested_name" in result.diagnostics[0].message
            assert result.diagnostics[-1].code == "FINALIZER_FAILED"
