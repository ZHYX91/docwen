"""Tests for WorkspaceManager and WorkspaceHandle."""

from __future__ import annotations

import hashlib
import os
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
)
from docwen_runtime.path_io import filesystem_path
from docwen_runtime.workspace.manager import WorkspaceHandle, WorkspaceManager

pytestmark = pytest.mark.integration


class TestWorkspaceHandle:
    def test_input_path(self) -> None:
        handle = WorkspaceHandle("/tmp/input.md", "/tmp/staging")
        assert handle.input_path == "/tmp/input.md"

    def test_staging_dir(self) -> None:
        handle = WorkspaceHandle("/tmp/input.md", "/tmp/staging/task1")
        assert handle.staging_dir == "/tmp/staging/task1"

    def test_create_artifact_path_unique(self) -> None:
        handle = WorkspaceHandle("/tmp/in.md", "/tmp/staging")
        p1 = handle.create_artifact_path("primary", ".docx")
        p2 = handle.create_artifact_path("primary", ".docx")
        assert p1 != p2  # unique paths
        assert p1.startswith("/tmp/staging")
        assert p2.startswith("/tmp/staging")
        assert p1.endswith(".docx")
        assert p2.endswith(".docx")

    def test_create_artifact_path_sanitizes_kind(self) -> None:
        handle = WorkspaceHandle("/tmp/in.md", "/tmp/staging")
        path = handle.create_artifact_path("aux/image", ".png")
        assert "/" not in os.path.basename(path)
        assert "aux_image" in os.path.basename(path)

    def test_add_artifact_registers(self) -> None:
        handle = WorkspaceHandle("/tmp/in.md", "/tmp/staging")
        art = ArtifactManifest(
            artifact_id="a1",
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path="/tmp/staging/out.docx",
            suggested_name="out.docx",
        )
        handle.add_artifact(art)
        assert len(handle.registered_artifacts) == 1
        assert handle.registered_artifacts[0].artifact_id == "a1"

    def test_registered_artifacts_is_copy(self) -> None:
        handle = WorkspaceHandle("/tmp/in.md", "/tmp/staging")
        art = ArtifactManifest(artifact_id="a1", kind="primary", staging_path="/s/o.docx", suggested_name="o.docx")
        handle.add_artifact(art)
        copy1 = handle.registered_artifacts
        copy1.clear()
        # internal list unchanged
        assert len(handle.registered_artifacts) == 1


class TestWorkspaceManager:
    @pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
    def test_typed_workspace_materializes_an_ordinary_long_path(self, tmp_path: Path) -> None:
        source = tmp_path
        while len(str(source / "source.md")) < 280:
            source /= "long-path-segment-0123456789"
        logical_source = source / "source.md"
        io_source = filesystem_path(logical_source, force_extended=True)
        io_source.parent.mkdir(parents=True)
        payload = b"# long path\n"
        io_source.write_bytes(payload)
        ref = FileRef(
            path=str(logical_source),
            format="markdown",
            category="markdown",
            size_bytes=len(payload),
            input_kind="document",
            input_role="source",
            logical_path="documents/source.md",
            media_type="text/markdown",
            metadata={
                "machine_input_size_bytes": len(payload),
                "machine_input_sha256": hashlib.sha256(payload).hexdigest(),
            },
        )
        manager = WorkspaceManager(root_dir=str(tmp_path / "workspaces"))

        handle = manager.create("long-input", str(logical_source), (ref,))

        assert not ref.path.startswith("\\\\?\\")
        assert Path(handle.input_path).read_bytes() == payload

    def test_resolved_numbering_inputs_are_copied_as_one_request(self, tmp_path: Path) -> None:
        neutral = tmp_path / "resolved-document.json"
        plan = tmp_path / "numbering-export-plan.json"
        neutral.write_bytes(b'{"schema_id":"docwen.resolved_document.v1"}')
        plan.write_bytes(b'{"schema_id":"docwen.numbering_export_plan.v1"}')

        def ref(
            path: Path,
            *,
            kind: str,
            role: str,
            logical_path: str,
            media_type: str,
        ) -> FileRef:
            data = path.read_bytes()
            return FileRef(
                path=str(path),
                format="markdown" if kind == "document" else "resource",
                category="markdown" if kind == "document" else "other",
                size_bytes=len(data),
                input_kind=kind,
                input_role=role,
                logical_path=logical_path,
                media_type=media_type,
                metadata={
                    "machine_input_size_bytes": len(data),
                    "machine_input_sha256": hashlib.sha256(data).hexdigest(),
                },
            )

        refs = (
            ref(
                neutral,
                kind="document",
                role="neutral_document",
                logical_path="request/resolved-document.json",
                media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
            ),
            ref(
                plan,
                kind="resource",
                role="numbering_export_plan",
                logical_path="request/numbering-export-plan.json",
                media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
            ),
        )
        manager = WorkspaceManager(root_dir=str(tmp_path / "workspaces"))

        handle = manager.create("resolved-numbering", str(neutral), refs)

        copied = handle.input_resources()
        assert [item.input_role for item in copied] == ["neutral_document", "numbering_export_plan"]
        assert [Path(item.path).read_bytes() for item in copied] == [neutral.read_bytes(), plan.read_bytes()]
        assert handle.input_path == copied[0].path
        assert handle.input_resources("neutral_document") == (copied[0],)
        assert handle.input_resources("numbering_export_plan") == (copied[1],)

    @pytest.mark.parametrize("mutated_role", ["neutral_document", "numbering_export_plan"])
    def test_resolved_numbering_copy_rejects_post_admission_mutation(
        self,
        tmp_path: Path,
        mutated_role: str,
    ) -> None:
        neutral = tmp_path / "resolved-document.json"
        plan = tmp_path / "numbering-export-plan.json"
        neutral.write_bytes(b"admitted neutral")
        plan.write_bytes(b"admitted plan")

        def ref(path: Path, *, kind: str, role: str) -> FileRef:
            data = path.read_bytes()
            return FileRef(
                path=str(path),
                format="markdown" if kind == "document" else "resource",
                category="markdown" if kind == "document" else "other",
                size_bytes=len(data),
                input_kind=kind,
                input_role=role,
                logical_path=f"request/{role}.json",
                media_type=(
                    RESOLVED_DOCUMENT_MEDIA_TYPE if role == "neutral_document" else NUMBERING_EXPORT_PLAN_MEDIA_TYPE
                ),
                metadata={
                    "machine_input_size_bytes": len(data),
                    "machine_input_sha256": hashlib.sha256(data).hexdigest(),
                },
            )

        refs = (
            ref(neutral, kind="document", role="neutral_document"),
            ref(plan, kind="resource", role="numbering_export_plan"),
        )
        (neutral if mutated_role == "neutral_document" else plan).write_bytes(b"mutated after admission")
        workspace_root = tmp_path / "workspaces"
        manager = WorkspaceManager(root_dir=str(workspace_root))

        with pytest.raises(ValueError, match="typed input copy failed integrity verification"):
            manager.create("resolved-numbering-mutated", str(neutral), refs)

        assert len(manager) == 0
        assert not list(workspace_root.iterdir())

    def test_typed_workspace_rejects_plan_without_primary_document(self, tmp_path: Path) -> None:
        plan = tmp_path / "numbering-export-plan.json"
        plan.write_bytes(b"plan")
        data = plan.read_bytes()
        plan_ref = FileRef(
            path=str(plan),
            format="resource",
            category="other",
            size_bytes=len(data),
            input_kind="resource",
            input_role="numbering_export_plan",
            logical_path="request/numbering-export-plan.json",
            media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
            metadata={
                "machine_input_size_bytes": len(data),
                "machine_input_sha256": hashlib.sha256(data).hexdigest(),
            },
        )
        workspace_root = tmp_path / "workspaces"
        manager = WorkspaceManager(root_dir=str(workspace_root))

        with pytest.raises(ValueError, match="typed input set has no primary document input"):
            manager.create("missing-neutral-document", str(plan), (plan_ref,))

        assert len(manager) == 0
        assert not list(workspace_root.iterdir())

    def test_typed_inputs_use_opaque_physical_names_and_preserve_logical_keys(self, tmp_path) -> None:
        source = tmp_path / "source.md"
        resource = tmp_path / "resource.png"
        source.write_text("![x](assets/x.png)\n", encoding="utf-8")
        resource.write_bytes(b"declared png")

        def ref(path, *, kind: str, role: str, logical_path: str, media_type: str) -> FileRef:
            data = path.read_bytes()
            return FileRef(
                path=str(path),
                format="markdown" if kind == "document" else "png",
                category="markdown" if kind == "document" else "image",
                size_bytes=len(data),
                input_kind=kind,
                input_role=role,
                logical_path=logical_path,
                media_type=media_type,
                metadata={
                    "machine_input_size_bytes": len(data),
                    "machine_input_sha256": hashlib.sha256(data).hexdigest(),
                },
            )

        refs = (
            ref(
                resource,
                kind="resource",
                role="linked_resource",
                logical_path="docs/payload:stream.png",
                media_type="image/png",
            ),
            ref(
                source,
                kind="document",
                role="source",
                logical_path="Docs/CON.md",
                media_type="text/markdown",
            ),
        )
        manager = WorkspaceManager(root_dir=str(tmp_path / "workspaces"))

        handle = manager.create("typed", str(source), refs)

        copied = handle.input_resources()
        assert [item.logical_path for item in copied] == ["docs/payload:stream.png", "Docs/CON.md"]
        assert [Path(item.path).name for item in copied] == ["input-0000.png", "input-0001.md"]
        assert handle.resource_by_logical_path("docs/payload:stream.png") is copied[0]
        assert [item.media_type for item in copied] == ["image/png", "text/markdown"]
        assert handle.input_path == copied[1].path
        assert Path(copied[0].path).read_bytes() == resource.read_bytes()
        assert Path(copied[1].path).read_bytes() == source.read_bytes()

    @pytest.mark.parametrize("logical_path", ["../escape.md", "a\\..\\escape.md", "file:a.md", "/root.md"])
    def test_typed_workspace_rejects_invalid_logical_paths(self, tmp_path, logical_path: str) -> None:
        source = tmp_path / "source.md"
        source.write_text("body", encoding="utf-8")
        ref = FileRef(
            path=str(source),
            format="markdown",
            category="markdown",
            input_kind="document",
            input_role="source",
            logical_path=logical_path,
        )
        manager = WorkspaceManager(root_dir=str(tmp_path / "workspaces"))

        with pytest.raises(ValueError, match="invalid typed input logical_path"):
            manager.create("typed", str(source), (ref,))

        assert len(manager) == 0

    def test_typed_workspace_rejects_symlink_source_without_reading_target(self, tmp_path) -> None:
        target = tmp_path / "target.md"
        link = tmp_path / "link.md"
        target.write_text("sentinel", encoding="utf-8")
        try:
            link.symlink_to(target)
        except OSError as exc:
            pytest.skip(f"file symlink unavailable: {exc}")
        ref = FileRef(
            path=str(link),
            format="markdown",
            category="markdown",
            input_kind="document",
            input_role="source",
            logical_path="source.md",
        )
        manager = WorkspaceManager(root_dir=str(tmp_path / "workspaces"))

        with pytest.raises(ValueError, match="link or junction"):
            manager.create("typed", str(link), (ref,))

        assert target.read_text(encoding="utf-8") == "sentinel"
        assert len(manager) == 0

    def test_create_workspace(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            handle = mgr.create("task-1", "/tmp/input.md")
            assert os.path.isdir(handle.staging_dir)
            assert handle.staging_dir.startswith(tmpdir)

    def test_workspace_name_is_short_opaque_and_cannot_escape_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            hostile_task_id = "../" * 20 + "request-" + "x" * 500

            handle = mgr.create(hostile_task_id, "/tmp/input.md")
            workspace_dir = os.path.dirname(handle.staging_dir)

            assert os.path.dirname(workspace_dir) == tmpdir
            assert os.path.basename(workspace_dir).startswith("dw-")
            assert hostile_task_id not in workspace_dir
            assert len(os.path.basename(workspace_dir)) <= 16

    def test_create_workspace_creates_missing_configured_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = os.path.join(tmpdir, "missing", "workspaces")
            mgr = WorkspaceManager(root_dir=root)

            handle = mgr.create("task-1", "/tmp/input.md")

            assert os.path.isdir(handle.staging_dir)
            assert os.path.dirname(os.path.dirname(handle.staging_dir)) == root

    def test_get_workspace(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            mgr.create("task-1", "/tmp/input.md")
            handle = mgr.get("task-1")
            assert handle is not None
            assert handle.input_path == "/tmp/input.md"

    def test_duplicate_task_id_fails_without_leaking_the_first_workspace(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            first = mgr.create("task-1", "/tmp/input.md")

            with pytest.raises(ValueError, match="Workspace already exists"):
                mgr.create("task-1", "/tmp/other.md")

            assert mgr.get("task-1") is first
            assert os.path.isdir(first.staging_dir)
            assert len(os.listdir(tmpdir)) == 1

    def test_get_nonexistent(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            assert mgr.get("nonexistent") is None

    def test_cleanup_removes_directory(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            handle = mgr.create("task-1", "/tmp/input.md")
            ws_dir = os.path.dirname(handle.staging_dir)
            assert os.path.isdir(ws_dir)

            mgr.cleanup("task-1")
            assert not os.path.isdir(ws_dir)
            assert mgr.get("task-1") is None

    def test_cleanup_nonexistent_noop(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            mgr.cleanup("nonexistent")  # no error

    def test_cleanup_all(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            h1 = mgr.create("task-1", "/tmp/a.md")
            h2 = mgr.create("task-2", "/tmp/b.md")
            d1 = os.path.dirname(h1.staging_dir)
            d2 = os.path.dirname(h2.staging_dir)

            mgr.cleanup_all()
            assert len(mgr) == 0
            assert not os.path.isdir(d1)
            assert not os.path.isdir(d2)

    def test_multiple_workspaces_isolated(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            h1 = mgr.create("task-1", "/tmp/a.md")
            h2 = mgr.create("task-2", "/tmp/b.md")
            assert h1.staging_dir != h2.staging_dir
            assert len(mgr) == 2

    def test_concurrent_workspaces_are_unique_and_cleanup_all_removes_them(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            with ThreadPoolExecutor(max_workers=8) as executor:
                handles = list(executor.map(lambda index: mgr.create(f"task-{index}", "/tmp/input.md"), range(32)))

            workspace_dirs = {os.path.dirname(handle.staging_dir) for handle in handles}
            assert len(workspace_dirs) == 32
            assert len(mgr) == 32

            mgr.cleanup_all()

            assert len(mgr) == 0
            assert not os.listdir(tmpdir)

    @pytest.mark.parametrize("root_dir", ["", ".", "relative/workspaces"])
    def test_relative_or_empty_root_is_rejected(self, root_dir: str) -> None:
        with pytest.raises(ValueError, match="workspace root must be absolute"):
            WorkspaceManager(root_dir=root_dir)

    def test_filesystem_root_is_rejected(self) -> None:
        filesystem_root = Path.cwd().anchor
        with pytest.raises(ValueError, match="must not be a filesystem root"):
            WorkspaceManager(root_dir=filesystem_root)

    def test_handle_writes_to_staging(self) -> None:
        """Plugin writes to staging via create_artifact_path."""
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            handle = mgr.create("task-write", "/tmp/input.md")
            path = handle.create_artifact_path("primary", ".txt")

            # Write something
            with open(path, "w") as f:
                f.write("hello")

            assert os.path.isfile(path)
            with open(path) as f:
                assert f.read() == "hello"

    def test_plugin_cannot_write_outside_staging(self) -> None:
        """Plugins are given staging_dir — writing elsewhere is misuse."""
        with tempfile.TemporaryDirectory() as tmpdir:
            mgr = WorkspaceManager(root_dir=tmpdir)
            handle = mgr.create("task-boundary", "/tmp/input.md")

            # The plugin can technically write anywhere, but the
            # workspace manager only tracks what's in staging.
            # This test verifies that create_artifact_path always
            # returns paths within staging_dir.
            path = handle.create_artifact_path("primary", ".docx")
            assert path.startswith(handle.staging_dir)
