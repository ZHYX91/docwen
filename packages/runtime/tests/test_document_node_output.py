from __future__ import annotations

import hashlib
import os
from datetime import datetime
from pathlib import Path

import pytest

from docwen_core.models import ArtifactOwnership, OutputPolicy
from docwen_runtime.output.finalizer import OutputFinalizer

from ._document_node_output_support import _artifact

pytestmark = pytest.mark.integration


@pytest.fixture
def frozen_node_clock(monkeypatch: pytest.MonkeyPatch) -> None:
    local_timezone = datetime.now().astimezone().tzinfo

    class FrozenDateTime(datetime):
        @classmethod
        def now(cls, tz=None):
            value = cls(2026, 8, 20, 12, 0, 0, tzinfo=local_timezone)
            return value if tz is None else value.astimezone(tz)

    monkeypatch.setattr("docwen_core.models.document_node.datetime", FrozenDateTime)


def test_markdown_bundle_is_published_as_one_document_node(tmp_path: Path) -> None:
    source = tmp_path / "公文.docx"
    source.write_bytes(b"source")
    staging = tmp_path / "staging"
    staging.mkdir()
    main = staging / "main.md"
    attachment = staging / "attachment.md"
    image = staging / "seal.png"
    main.write_text("[附件](公文_attachments.md)\n![印章](seal.png)\n", encoding="utf-8")
    attachment.write_text("![印章](seal.png)\n", encoding="utf-8")
    image.write_bytes(b"png")
    output = tmp_path / "output"

    result = OutputFinalizer().finalize(
        "task.node",
        [
            _artifact(
                main,
                artifact_id="main",
                suggested_name="公文.md",
                media_type="text/markdown",
                primary=True,
            ),
            _artifact(
                attachment,
                artifact_id="attachment",
                suggested_name="公文_attachments.md",
                media_type="text/markdown",
                ownership=ArtifactOwnership("attachment_of", "main", "attachment", 0, "附件"),
            ),
            _artifact(
                image,
                artifact_id="image",
                suggested_name="seal.png",
                media_type="image/png",
            ),
        ],
        OutputPolicy(output_dir=str(output), overwrite_mode="error"),
        input_path=str(source),
    )

    assert result.success is True
    markdown = [artifact for artifact in result.artifacts if artifact.media_type == "text/markdown"]
    assert len(markdown) == 2
    assert all(Path(artifact.staging_path).parent.name == Path(artifact.staging_path).stem for artifact in markdown)
    root = Path(result.metrics.extra["document_node_root"])
    assert root.name.startswith("公文_") and root.name.endswith("_fromDocx")
    assert (root / f"{root.name}.md").is_file()
    assert not (root / "docwen-node.json").exists()
    assert len(result.artifacts) == 3
    for artifact in result.artifacts:
        data = Path(artifact.staging_path).read_bytes()
        assert artifact.size_bytes == len(data)
        assert artifact.sha256 == hashlib.sha256(data).hexdigest()
        assert artifact.logical_path == Path(artifact.staging_path).relative_to(output).as_posix()
    main_text = (root / f"{root.name}.md").read_text(encoding="utf-8")
    attachment_artifact = next(item for item in markdown if not item.is_primary)
    attachment_path = Path(attachment_artifact.staging_path)
    assert attachment_path.relative_to(root).as_posix() in main_text
    assert "../seal.png" in attachment_path.read_text(encoding="utf-8")


def test_document_node_failure_leaves_no_partial_root(tmp_path: Path) -> None:
    source = tmp_path / "report.docx"
    source.write_bytes(b"source")
    missing = tmp_path / "missing.md"
    output = tmp_path / "output"

    result = OutputFinalizer().finalize(
        "task.node.fail",
        [
            _artifact(
                missing,
                artifact_id="main",
                suggested_name="report.md",
                media_type="text/markdown",
                primary=True,
            )
        ],
        OutputPolicy(output_dir=str(output), overwrite_mode="error"),
        input_path=str(source),
    )

    assert result.success is False
    assert result.artifacts == []
    assert list(output.glob("report_*_fromDocx")) == []
    assert list(output.glob(".__docwen-node-*")) == []


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_document_node_temp_publish_can_cross_max_path(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"source")
    markdown = tmp_path / "source.md"
    markdown.write_text("body\n", encoding="utf-8")
    output = tmp_path / "output"
    remaining = 235 - len(os.path.abspath(output)) - 1
    if remaining < 1:
        pytest.skip("pytest temp root is too long for the private-child boundary")
    output /= "x" * remaining
    assert len(os.path.abspath(output)) == 235

    result = OutputFinalizer().finalize(
        "task.node.long-path",
        [
            _artifact(
                markdown,
                artifact_id="main",
                suggested_name="source.md",
                media_type="text/markdown",
                primary=True,
            )
        ],
        OutputPolicy(output_dir=str(output), overwrite_mode="error"),
        input_path=str(source),
    )

    assert result.success is True, result.diagnostics
    primary = Path(result.artifacts[0].staging_path)
    assert len(os.path.abspath(primary)) >= 260
    assert OutputFinalizer._io_path(primary).read_text(encoding="utf-8") == "body\n"


def test_markdown_rejects_exact_output_file_path(tmp_path: Path) -> None:
    source = tmp_path / "report.docx"
    source.write_bytes(b"source")
    markdown = tmp_path / "report.md"
    markdown.write_text("body", encoding="utf-8")

    try:
        OutputFinalizer().finalize(
            "task.node.path",
            [
                _artifact(
                    markdown,
                    artifact_id="main",
                    suggested_name="report.md",
                    media_type="text/markdown",
                    primary=True,
                )
            ],
            OutputPolicy(output_path=str(tmp_path / "leaf.md")),
            input_path=str(source),
        )
    except ValueError as exc:
        assert "output parent directory" in str(exc)
    else:
        raise AssertionError("Markdown output_path must be rejected")


def test_explicit_in_place_markdown_transform_remains_a_file_update(tmp_path: Path) -> None:
    source = tmp_path / "report.md"
    source.write_text("old\n", encoding="utf-8")
    staging = tmp_path / "staging.md"
    staging.write_text("new\n", encoding="utf-8")

    result = OutputFinalizer().finalize(
        "task.node.in-place",
        [
            _artifact(
                staging,
                artifact_id="main",
                suggested_name="report.md",
                media_type="text/markdown",
                primary=True,
            )
        ],
        OutputPolicy(output_path=str(source), overwrite_mode="overwrite"),
        input_path=str(source),
    )

    assert result.success is True
    assert source.read_text(encoding="utf-8") == "new\n"
    assert [artifact.staging_path for artifact in result.artifacts] == [str(source)]


def test_document_node_collision_policy_applies_to_the_complete_root(
    tmp_path: Path,
    frozen_node_clock: None,
) -> None:
    source = tmp_path / "report.docx"
    source.write_bytes(b"source")
    output = tmp_path / "output"
    finalizer = OutputFinalizer()

    def publish(body: str, mode: str):
        staging = tmp_path / f"staging-{mode}-{len(body)}.md"
        staging.write_text(body, encoding="utf-8")
        return finalizer.finalize(
            f"task.node.{mode}",
            [
                _artifact(
                    staging,
                    artifact_id="main",
                    suggested_name="report.md",
                    media_type="text/markdown",
                    primary=True,
                )
            ],
            OutputPolicy(output_dir=str(output), overwrite_mode=mode),
            input_path=str(source),
        )

    first = publish("first\n", "error")
    rejected = publish("rejected\n", "error")
    renamed = publish("renamed\n", "rename")
    overwritten = publish("overwritten\n", "overwrite")

    assert first.success is True
    assert rejected.success is False
    assert rejected.error is not None and rejected.error.diagnostic_code == "DOCUMENT_NODE_EXISTS"
    first_root = Path(first.metrics.extra["document_node_root"])
    renamed_root = Path(renamed.metrics.extra["document_node_root"])
    assert renamed.success is True and renamed_root != first_root
    assert renamed_root.name.endswith("_001_fromDocx")
    assert (renamed_root / f"{renamed_root.name}.md").read_text(encoding="utf-8") == "renamed\n"
    assert overwritten.success is False
    assert overwritten.error is not None
    assert overwritten.error.diagnostic_code == "DOCUMENT_NODE_POLICY_UNSUPPORTED"
    assert (first_root / f"{first_root.name}.md").read_text(encoding="utf-8") == "first\n"
    reused = publish("first\n", "skip")
    assert reused.success and reused.metrics.output_bytes == 0
    assert reused.artifacts[0].metadata["document_node_reused"] is True
    mismatch = publish("changed\n", "skip")
    assert mismatch.error is not None
    assert not mismatch.success and mismatch.error.diagnostic_code == "DOCUMENT_NODE_SKIP_MISMATCH"
    assert (first_root / f"{first_root.name}.md").read_text(encoding="utf-8") == "first\n"


def test_overwrite_refuses_an_unowned_directory(tmp_path: Path, frozen_node_clock: None) -> None:
    source = tmp_path / "report.docx"
    source.write_bytes(b"source")
    staging = tmp_path / "report.md"
    staging.write_text("new\n", encoding="utf-8")
    output = tmp_path / "output"
    unowned = output / "report_20260820_120000_fromDocx"
    unowned.mkdir(parents=True)
    sentinel = unowned / "keep.txt"
    sentinel.write_text("keep\n", encoding="utf-8")

    result = OutputFinalizer().finalize(
        "task.node.unowned",
        [
            _artifact(
                staging,
                artifact_id="main",
                suggested_name="report.md",
                media_type="text/markdown",
                primary=True,
            )
        ],
        OutputPolicy(output_dir=str(output), overwrite_mode="overwrite"),
        input_path=str(source),
    )

    assert result.success is False
    assert sentinel.read_text(encoding="utf-8") == "keep\n"
    assert not (unowned / "docwen-node.json").exists()


@pytest.mark.parametrize("extra", ["file", "directory", "missing"])
def test_skip_requires_the_complete_current_output_tree(tmp_path, frozen_node_clock, extra):
    staging = tmp_path / "source.md"
    staging.write_bytes(b"body\n")
    output = tmp_path / "out"
    artifact = _artifact(
        staging, artifact_id="main", suggested_name="source.md", media_type="text/markdown", primary=True
    )
    finalizer = OutputFinalizer()
    policy = OutputPolicy(output_dir=str(output), overwrite_mode="skip")
    first = finalizer.finalize("first", [artifact], policy, input_path=str(staging))
    assert first.success
    root = Path(first.artifacts[0].staging_path).parent
    if extra == "file":
        (root / "user.txt").write_bytes(b"keep")
    elif extra == "directory":
        (root / "user-folder").mkdir()
    else:
        Path(first.artifacts[0].staging_path).unlink()
    before = sorted(str(path.relative_to(root)) for path in root.rglob("*"))
    second = finalizer.finalize("second", [artifact], policy, input_path=str(staging))
    assert not second.success and second.artifacts == []
    assert second.error is not None
    assert second.error.diagnostic_code == "DOCUMENT_NODE_SKIP_MISMATCH"
    assert sorted(str(path.relative_to(root)) for path in root.rglob("*")) == before
    assert not list(output.glob(".__docwen-node-*"))


def test_racing_empty_directory_is_not_replaced(tmp_path, monkeypatch):
    staging = tmp_path / "source.md"
    staging.write_bytes(b"body")
    output = tmp_path / "out"
    publish = OutputFinalizer._publish_directory_no_clobber
    raced = []

    def race(source, destination):
        path = Path(destination)
        path.mkdir()
        raced.append(path)
        publish(source, destination)

    monkeypatch.setattr(OutputFinalizer, "_publish_directory_no_clobber", staticmethod(race))
    result = OutputFinalizer().finalize(
        "race",
        [_artifact(staging, artifact_id="main", suggested_name="source.md", media_type="text/markdown", primary=True)],
        OutputPolicy(output_dir=str(output)),
        input_path=str(staging),
    )
    assert not result.success and result.artifacts == []
    assert raced[0].is_dir() and list(raced[0].iterdir()) == []
    assert not list(output.glob(".__docwen-node-*"))


def test_reused_output_survives_cleanup_failure_with_visible_warning(tmp_path, frozen_node_clock, monkeypatch):
    staging = tmp_path / "source.md"
    staging.write_bytes(b"body")
    output = tmp_path / "out"
    artifact = _artifact(
        staging, artifact_id="main", suggested_name="source.md", media_type="text/markdown", primary=True
    )
    finalizer = OutputFinalizer()
    policy = OutputPolicy(output_dir=str(output), overwrite_mode="skip")
    assert finalizer.finalize("first", [artifact], policy, input_path=str(staging)).success

    def fail_cleanup(*args, **kwargs):
        raise PermissionError("temporary file is busy")

    monkeypatch.setattr("docwen_runtime.output.finalizer.shutil.rmtree", fail_cleanup)
    result = finalizer.finalize("again", [artifact], policy, input_path=str(staging))
    assert result.success and result.metrics.output_bytes == 0
    assert Path(result.artifacts[0].staging_path).read_bytes() == b"body"
    assert any(item.code == "FINALIZER_CLEANUP_FAILED" and item.level == "warning" for item in result.diagnostics)


def test_owner_scoped_resources_publish_with_each_document_and_bundle(tmp_path, frozen_node_clock):
    from docwen_application.bundle_mapping import build_bundle_draft
    from docwen_core.models import validate_artifact_bundle_draft

    stage = tmp_path / "stage"
    stage.mkdir()
    main = stage / "main.md"
    child = stage / "child.md"
    image = stage / "shared.png"
    original = stage / "data.csv"
    original_md = stage / "sample.md"
    main.write_text("[annex](annex.md)\n![main](shared.png#main)\n", encoding="utf-8")
    child.write_text("![child](shared.png?size=2#child)\n[data](data.csv)\n[original](sample.md)\n", encoding="utf-8")
    image.write_bytes(b"same original image bytes")
    original.write_bytes(b"name,value\nA,2\n")
    original_md.write_bytes(b"[original resource bytes](shared.png)\r\n")
    artifacts = [
        _artifact(main, artifact_id="main", suggested_name="main.md", media_type="text/markdown", primary=True),
        _artifact(
            child,
            artifact_id="child",
            suggested_name="annex.md",
            media_type="text/markdown",
            ownership=ArtifactOwnership("attachment_of", "main", "attachment", 0, "附件"),
        ),
        _artifact(
            image,
            artifact_id="main-image",
            suggested_name="shared.png",
            media_type="image/png",
            ownership=ArtifactOwnership("resource_of", "main", "image", 0),
        ),
        _artifact(
            image,
            artifact_id="child-image",
            suggested_name="shared.png",
            media_type="image/png",
            ownership=ArtifactOwnership("resource_of", "child", "image", 0),
        ),
        _artifact(
            original,
            artifact_id="original",
            suggested_name="data.csv",
            media_type="text/csv",
            ownership=ArtifactOwnership("resource_of", "child", "original", 1),
        ),
        _artifact(
            original_md,
            artifact_id="original-md",
            suggested_name="sample.md",
            media_type="text/markdown",
            ownership=ArtifactOwnership("resource_of", "child", "original", 2),
        ),
    ]
    output = tmp_path / "out"
    roots = []
    for _ in range(2):
        result = OutputFinalizer().finalize(
            "owner-test",
            list(reversed(artifacts)),
            OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
            input_path=str(tmp_path / "notice.docx"),
        )
        assert result.success, result.error
        by_id = {item.artifact_id: item for item in result.artifacts}
        main_path = Path(by_id["main"].staging_path)
        child_path = Path(by_id["child"].staging_path)
        roots.append(main_path.parent)
        assert child_path.parent.parent == main_path.parent
        assert "../" not in child_path.read_text(encoding="utf-8")
        assert "shared.png?size=2#child" in child_path.read_text(encoding="utf-8")
        assert child_path.relative_to(main_path.parent).as_posix() in main_path.read_text(encoding="utf-8")
        assert Path(by_id["main-image"].staging_path).parent == main_path.parent
        for artifact_id in ("child-image", "original"):
            assert Path(by_id[artifact_id].staging_path).parent == child_path.parent
        assert Path(by_id["child-image"].staging_path).read_bytes() == image.read_bytes()
        assert Path(by_id["main-image"].staging_path).read_bytes() == image.read_bytes()
        assert Path(by_id["original"].staging_path).name == "data.csv"
        assert Path(by_id["original-md"].staging_path).is_relative_to(child_path.parent)
        assert Path(by_id["original-md"].staging_path).read_bytes() == original_md.read_bytes()
        for item in result.artifacts:
            assert item.sha256 == hashlib.sha256(Path(item.staging_path).read_bytes()).hexdigest()
        draft = build_bundle_draft(
            profile="document_with_resources", output_media_type="text/markdown", artifacts=result.artifacts
        )
        validate_artifact_bundle_draft(draft)
        owners = {relation.source_artifact_id: relation.target_artifact_id for relation in draft.relations}
        assert owners == {
            "child": "main",
            "main-image": "main",
            "child-image": "child",
            "original": "child",
            "original-md": "child",
        }
    assert roots[0] != roots[1]
    assert image.is_file(), "publishing duplicate byte sources must not consume the shared staging file"


@pytest.mark.parametrize("case", ["attachment_under_fragment", "physical_page"])
def test_invalid_declared_structure_is_rejected_before_publication(tmp_path, case):
    artifacts = []
    for artifact_id in ("main", "fragment", "child"):
        path = tmp_path / f"{artifact_id}.md"
        path.write_text(artifact_id, encoding="utf-8")
        artifacts.append(
            _artifact(
                path,
                artifact_id=artifact_id,
                suggested_name=path.name,
                media_type="text/markdown",
                primary=artifact_id == "main",
            )
        )
    artifacts[1].ownership = ArtifactOwnership("fragment_of", "main", "ocr_text", 0)
    artifacts[2].ownership = (
        ArtifactOwnership("attachment_of", "fragment", "attachment", 0)
        if case == "attachment_under_fragment"
        else ArtifactOwnership("fragment_of", "main", "ocr_page", 1)
    )
    output = tmp_path / "output"
    with pytest.raises(ValueError):
        OutputFinalizer().finalize(
            "invalid-owner", artifacts, OutputPolicy(output_dir=str(output)), input_path=str(tmp_path / "source.docx")
        )
    assert not output.exists() or not list(output.iterdir())


def test_root_collision_keeps_root_named_attachment_resource_in_its_owner_directory(tmp_path, frozen_node_clock):
    from docwen_runtime.output.document_node import plan_document_node_layout

    main = _artifact(
        tmp_path / "main.md", artifact_id="main", suggested_name="main.md", media_type="text/markdown", primary=True
    )
    child = _artifact(
        tmp_path / "child.md",
        artifact_id="child",
        suggested_name="child.md",
        media_type="text/markdown",
        ownership=ArtifactOwnership("attachment_of", "main", "attachment", 0, "Annex"),
    )
    image = _artifact(
        tmp_path / "image.png",
        artifact_id="image",
        suggested_name="notice_20260820_120000_fromDocx.png",
        media_type="image/png",
        ownership=ArtifactOwnership("resource_of", "child", "image", 0),
    )
    plan = plan_document_node_layout(
        task_id="rename-owner", artifacts=[main, child, image], input_path=str(tmp_path / "notice.docx")
    )
    rebased = plan.rebase_root("renamed-result")
    by_id: dict[str, Path] = {}
    for item in rebased.artifacts:
        assert item.logical_path is not None
        by_id[item.artifact_id] = Path(item.logical_path)
    assert by_id["main"].as_posix() == "renamed-result/renamed-result.md"
    assert by_id["image"].parent == by_id["child"].parent
    assert by_id["image"].name == image.suggested_name
