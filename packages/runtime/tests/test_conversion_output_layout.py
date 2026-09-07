"""Conversion folders preserve the admitted source and every deliverable."""

import hashlib
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

from docwen_core.models import ArtifactManifest, ConversionIdentity, ConversionRequest, FileRef, OutputPolicy
from docwen_runtime.output.document_node import plan_document_node_layout
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.output.identity import conversion_identity

pytestmark = pytest.mark.integration


def _identity() -> ConversionIdentity:
    return ConversionIdentity.create(
        task_id="conversion",
        source_stem="项目记录",
        source_format="markdown",
        source_name="项目记录.md",
        created_at=datetime(2026, 9, 7, 18, tzinfo=timezone(timedelta(hours=8))),
    )


def test_csv_sheet_label_precedes_shared_timestamp_and_source(tmp_path):
    sheets = ("人员明细", "部门")
    artifacts = []
    for index, sheet in enumerate(sheets):
        path = tmp_path / f"staged-{index}.csv"
        path.write_text("name\nvalue\n", encoding="utf-8")
        artifacts.append(
            ArtifactManifest(
                str(index),
                "primary",
                str(path),
                path.name,
                "text/csv",
                is_primary=index == 0,
                metadata={"sheet_name": sheet, "sheet_index": index},
            )
        )
    output = tmp_path / "out"
    result = OutputFinalizer().finalize(
        "conversion",
        artifacts,
        OutputPolicy(output_dir=str(output)),
        input_path=str(tmp_path / "项目记录.md"),
        group_outputs=True,
        identity=_identity(),
    )
    assert result.success
    root = output / "项目记录_20260907_180000_fromMd"
    assert [Path(a.staging_path).name for a in result.artifacts if a.is_primary] == [
        "项目记录_人员明细_20260907_180000_fromMd.csv"
    ]
    assert {p.name for p in root.iterdir()} == {
        *(f"项目记录_{sheet}_20260907_180000_fromMd.csv" for sheet in sheets),
        "docwen-node.json",
    }
    manifest = json.loads((root / "docwen-node.json").read_text(encoding="utf-8"))
    assert manifest["source"]["name"] == "项目记录.md"


def test_binary_primary_with_markdown_report_keeps_binary_primary_and_collision_name(tmp_path):
    docx = tmp_path / "stage.docx"
    docx.write_bytes(b"docx")
    report = tmp_path / "report.md"
    report.write_text("report", encoding="utf-8")
    artifacts = [
        ArtifactManifest("report", "auxiliary", str(report), "report.md", "text/markdown"),
        ArtifactManifest(
            "word",
            "primary",
            str(docx),
            "document.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            is_primary=True,
        ),
    ]
    plan = plan_document_node_layout(
        task_id="conversion", artifacts=artifacts, input_path="项目记录.md", identity=_identity()
    )
    rebased = plan.rebase_root(_identity().node_name(collision=1))
    primary = next(a for a in rebased.artifacts if a.is_primary)
    assert primary.logical_path == f"{rebased.root_name}/{rebased.root_name}.docx"
    child = next(a for a in rebased.artifacts if a.media_type == "text/markdown")
    assert child.logical_path is not None
    assert Path(child.logical_path).parent.name == Path(child.logical_path).stem


def test_admitted_source_name_and_format_win_over_intermediate(tmp_path):
    ref = FileRef(str(tmp_path / "resolved-document.json"), "markdown", "text", logical_path="Notes/项目记录.md")
    identity = conversion_identity("machine", ref)
    assert identity.source_name == "项目记录.md"
    assert identity.source_stem == "项目记录"
    assert identity.source_tag == "Md"
    original = tmp_path / "通知.wps"
    original.write_bytes(b"original WPS content")
    (tmp_path / "hub.docx").write_bytes(b"different DOCX content")
    ref = FileRef(
        str(tmp_path / "hub.docx"),
        "docx",
        "document",
        metadata={"_docwen_preconversion_source": {"path": str(tmp_path / "通知.wps"), "format": "wps"}},
    )
    identity = conversion_identity("wps", ref)
    assert identity.source_stem == "通知"
    assert identity.source_tag == "Wps"
    assert identity.source_sha256 == hashlib.sha256(original.read_bytes()).hexdigest()
    staged = tmp_path / "stage.md"
    staged.write_text("notification", encoding="utf-8")
    result = OutputFinalizer().finalize(
        "wps",
        [ArtifactManifest("main", "primary", str(staged), "stage.md", "text/markdown", is_primary=True)],
        OutputPolicy(output_dir=str(tmp_path / "out")),
        input_path=ref.path,
        identity=identity,
    )
    assert result.success
    manifest = next(item for item in result.artifacts if item.kind == "manifest")
    source = json.loads(Path(manifest.staging_path).read_text(encoding="utf-8"))["source"]
    assert source["name"] == original.name
    assert source["sha256"] == hashlib.sha256(original.read_bytes()).hexdigest()


def test_provenance_is_reused_only_for_matching_manifest_and_unchanged_output(tmp_path):
    source = tmp_path / "项目记录.md"
    source.write_text("body", encoding="utf-8")
    result = OutputFinalizer().finalize(
        "conversion",
        [ArtifactManifest("primary", "primary", str(source), "stage.md", "text/markdown", is_primary=True)],
        OutputPolicy(output_dir=str(tmp_path / "out")),
        input_path=str(source),
        identity=_identity(),
    )
    output = next(a.staging_path for a in result.artifacts if a.is_primary)
    ref = FileRef(output, "markdown", "text")
    assert conversion_identity("again", ref).source_stem == "项目记录"
    from docwen_core.cancellation import CancellationToken
    from docwen_core.errors import CancellationRequested

    token = CancellationToken()
    token.cancel(reason="user_cancelled")
    with pytest.raises(CancellationRequested):
        conversion_identity("cancelled", ref, cancellation=token)
    Path(output).write_text("edited", encoding="utf-8")
    assert conversion_identity("edited", ref).source_stem == Path(output).stem
    assert (
        conversion_identity("unproven", FileRef(str(tmp_path / "用户_fromMd.md"), "markdown", "text")).source_stem
        == "用户_fromMd"
    )


def test_conversion_identity_survives_request_transport_without_recomputing_time(tmp_path):
    request = ConversionRequest(
        "conversion",
        [FileRef(str(tmp_path / "private.json"), "markdown", "text")],
        "docx",
        conversion_identity=_identity(),
    )
    restored = ConversionRequest.from_dict(request.to_dict())
    assert restored.conversion_identity == request.conversion_identity
    assert restored.source_stem == "项目记录"


def test_csv_sanitized_names_do_not_lose_worksheets_or_steal_a_later_sheet_name(tmp_path):
    sheets = ("A<B", "A>B", "A_B_2")
    artifacts = []
    for index, name in enumerate(sheets):
        staged = tmp_path / f"stage-{index}.csv"
        staged.write_text(name, encoding="utf-8")
        artifacts.append(
            ArtifactManifest(
                str(index),
                "primary",
                str(staged),
                staged.name,
                "text/csv",
                is_primary=index == 0,
                metadata={"sheet_name": name, "sheet_index": index},
            )
        )
    result = OutputFinalizer().finalize(
        "collision",
        artifacts,
        OutputPolicy(output_dir=str(tmp_path / "out")),
        input_path="项目记录.md",
        identity=_identity(),
        group_outputs=True,
    )
    assert result.success
    csv_files = [Path(item.staging_path) for item in result.artifacts if item.media_type == "text/csv"]
    assert {item.read_text(encoding="utf-8") for item in csv_files} == set(sheets)
    assert len({item.name.casefold() for item in csv_files}) == 3
    assert csv_files[2].name == "项目记录_A_B_2_20260907_180000_fromMd.csv"


@pytest.mark.parametrize("fail_audit", [False, True])
def test_optional_audit_is_covered_by_the_same_directory_transaction(tmp_path, monkeypatch, fail_audit):
    from docwen_core.models.result import ConversionResult
    from docwen_runtime.output.manifest import OutputManifestWriter

    source = tmp_path / "项目记录.md"
    source.write_text("content", encoding="utf-8")
    artifacts = [ArtifactManifest("main", "primary", str(source), "stage.md", "text/markdown", is_primary=True)]
    request = ConversionRequest(
        "audit",
        [FileRef(str(source), "markdown", "text")],
        "md",
        config_snapshot={"output": {"manifest": {"save_to_output": True}}},
    )
    audit = OutputManifestWriter.build_for_success(request, ConversionResult("audit", True, artifacts=artifacts))
    assert audit is not None
    if fail_audit:

        def fail(*args, **kwargs):
            raise OSError("audit preparation failed")

        monkeypatch.setattr("docwen_runtime.output.finalizer.stage_node_audit", fail)
    output = tmp_path / "out"
    result = OutputFinalizer().finalize(
        "audit",
        artifacts,
        OutputPolicy(output_dir=str(output)),
        input_path=str(source),
        identity=_identity(),
        audit_document=audit,
    )
    if fail_audit:
        assert not result.success
        assert not list(output.glob("项目记录*"))
        return
    assert result.success
    assert len(list(output.iterdir())) == 1
    node = json.loads(Path(result.artifacts[-1].staging_path).read_text(encoding="utf-8"))
    record = next(item for item in node["artifacts"] if item["role"] == "audit")
    audit_path = output / record["logical_path"]
    assert record["sha256"] == hashlib.sha256(audit_path.read_bytes()).hexdigest()
    assert (
        json.loads(audit_path.read_text(encoding="utf-8"))["artifacts"][0]["name"]
        == "项目记录_20260907_180000_fromMd.md"
    )
