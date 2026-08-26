from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path

import pytest

from docwen_core.models import ArtifactManifest, OutputPolicy
from docwen_runtime.output.finalizer import OutputFinalizer

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


def _artifact(
    path: Path,
    *,
    artifact_id: str,
    suggested_name: str,
    media_type: str,
    primary: bool = False,
    metadata: dict | None = None,
) -> ArtifactManifest:
    return ArtifactManifest(
        artifact_id=artifact_id,
        kind="primary" if primary else "auxiliary",
        staging_path=str(path),
        suggested_name=suggested_name,
        media_type=media_type,
        metadata=metadata or {},
        is_primary=primary,
    )


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
                metadata={"source_kind": "gongwen_attachment"},
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
    manifest = json.loads((root / "docwen-node.json").read_text(encoding="utf-8"))
    assert manifest["schema"] == "docwen.document_node.v1"
    assert manifest["source"]["sha256"]
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
    assert overwritten.success is True
    assert Path(overwritten.metrics.extra["document_node_root"]) == first_root
    assert (first_root / f"{first_root.name}.md").read_text(encoding="utf-8") == "overwritten\n"


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
