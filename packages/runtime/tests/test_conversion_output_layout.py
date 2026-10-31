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
        identity=_identity(),
    )
    assert result.success
    root = output / "项目记录_20260907_180000_fromMd"
    assert [Path(a.staging_path).name for a in result.artifacts if a.is_primary] == [
        "项目记录_人员明细_20260907_180000_fromMd.csv"
    ]
    assert {p.name for p in root.iterdir()} == {
        *(f"项目记录_{sheet}_20260907_180000_fromMd.csv" for sheet in sheets),
    }
    assert all(
        item.sha256 == hashlib.sha256(Path(item.staging_path).read_bytes()).hexdigest() for item in result.artifacts
    )


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
    assert len(result.artifacts) == 1
    primary = result.artifacts[0]
    assert Path(primary.staging_path).parent.name.startswith("通知_")
    assert Path(primary.staging_path).parent.name.endswith("_fromWps")
    assert primary.sha256 == hashlib.sha256(b"notification").hexdigest()


def test_later_conversion_uses_its_actual_input_name_without_sidecar_provenance(tmp_path):
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
    assert conversion_identity("again", ref).source_stem == Path(output).stem
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
    assert [item.kind for item in result.artifacts] == ["primary", "manifest"]
    audit = result.artifacts[-1]
    audit_path = Path(audit.staging_path)
    assert audit.sha256 == hashlib.sha256(audit_path.read_bytes()).hexdigest()
    assert audit.logical_path == f"{audit_path.parent.name}/{audit_path.name}"
    assert not (audit_path.parent / "docwen-node.json").exists()
    assert (
        json.loads(audit_path.read_text(encoding="utf-8"))["artifacts"][0]["name"]
        == "项目记录_20260907_180000_fromMd.md"
    )


@pytest.mark.parametrize(
    "suffix,media_type,sheet_name",
    [
        (".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", None),
        (".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", None),
        (".csv", "text/csv", "人员"),
    ],
)
def test_single_markdown_conversion_is_a_named_file(tmp_path, suffix, media_type, sheet_name):
    staged = tmp_path / f"generated{suffix}"
    staged.write_bytes(b"single result")
    output = tmp_path / "out"
    artifact = ArtifactManifest(
        "main",
        "primary",
        str(staged),
        staged.name,
        media_type,
        is_primary=True,
        metadata={"sheet_name": sheet_name} if sheet_name else {},
    )
    result = OutputFinalizer().finalize(
        "single",
        [artifact],
        OutputPolicy(output_dir=str(output)),
        input_path="项目记录.md",
        identity=_identity(),
    )
    assert result.success
    name = f"项目记录{'_人员' if sheet_name else ''}_20260907_180000_fromMd{suffix}"
    assert [path.name for path in output.iterdir()] == [name]
    assert (output / name).read_bytes() == b"single result"
    assert result.artifacts[0].logical_path == name
    assert "document_node_schema" not in result.artifacts[0].metadata


def test_single_markdown_conversion_honors_exact_path_and_overwrite(tmp_path):
    staged = tmp_path / "generated.docx"
    staged.write_bytes(b"new")
    output = tmp_path / "chosen.docx"
    output.write_bytes(b"old")
    result = OutputFinalizer().finalize(
        "exact",
        [ArtifactManifest("main", "primary", str(staged), staged.name, is_primary=True)],
        OutputPolicy(output_path=str(output), overwrite_mode="overwrite"),
        input_path="项目记录.md",
        identity=_identity(),
    )
    assert result.success
    assert result.artifacts[0].staging_path == str(output)
    assert output.read_bytes() == b"new"


def test_multiple_csv_requires_parent_not_exact_output(tmp_path):
    artifacts = []
    for index in range(2):
        staged = tmp_path / f"sheet{index}.csv"
        staged.write_bytes(b"cell")
        artifacts.append(
            ArtifactManifest(
                str(index),
                "primary" if index == 0 else "auxiliary",
                str(staged),
                staged.name,
                "text/csv",
                is_primary=index == 0,
            )
        )
    with pytest.raises(ValueError, match="output parent directory"):
        OutputFinalizer().finalize(
            "many",
            artifacts,
            OutputPolicy(output_path=str(tmp_path / "exact.csv")),
            input_path="项目记录.md",
            identity=_identity(),
        )
    assert not (tmp_path / "exact.csv").exists()


def test_retained_intermediate_does_not_turn_one_document_into_a_directory(tmp_path):
    staged = tmp_path / "generated.docx"
    staged.write_bytes(b"document")
    intermediate = tmp_path / "hub.xlsx"
    intermediate.write_bytes(b"private hub")
    output = tmp_path / "out"
    artifacts = [
        ArtifactManifest("main", "primary", str(staged), staged.name, is_primary=True),
        ArtifactManifest(
            "hub", "auxiliary", str(intermediate), intermediate.name, metadata={"source": "preconversion"}
        ),
    ]
    result = OutputFinalizer().finalize(
        "single", artifacts, OutputPolicy(output_dir=str(output)), input_path="项目记录.md", identity=_identity()
    )
    assert result.success
    assert {path.name for path in output.iterdir()} == {"项目记录_20260907_180000_fromMd.docx", "hub.xlsx"}
    assert all(path.is_file() for path in output.iterdir())


def test_markdown_validation_report_keeps_action_naming(tmp_path):
    report = tmp_path / "review.json"
    report.write_bytes(b"{}")
    output = tmp_path / "out"
    result = OutputFinalizer().finalize(
        "review",
        [ArtifactManifest("report", "primary", str(report), report.name, "application/json", is_primary=True)],
        OutputPolicy(output_dir=str(output)),
        input_path="项目记录.md",
        identity=_identity(),
        conversion_output=False,
    )
    assert result.success
    assert (output / "review.json").is_file()


@pytest.mark.parametrize("mode", ["rename", "error", "skip", "overwrite"])
def test_single_conversion_uses_file_collision_policy(tmp_path, mode):
    staged = tmp_path / "generated.docx"
    staged.write_bytes(b"new")
    output = tmp_path / "out"
    output.mkdir()
    original = output / "项目记录_20260907_180000_fromMd.docx"
    original.write_bytes(b"old")
    result = OutputFinalizer().finalize(
        "collision",
        [ArtifactManifest("main", "primary", str(staged), staged.name, is_primary=True)],
        OutputPolicy(output_dir=str(output), overwrite_mode=mode),
        input_path="项目记录.md",
        identity=_identity(),
    )
    if mode == "error":
        assert not result.success
        assert original.read_bytes() == b"old"
    elif mode == "overwrite":
        assert result.success
        assert original.read_bytes() == b"new"
    elif mode == "rename":
        assert result.success
        assert original.read_bytes() == b"old"
        assert Path(result.artifacts[0].staging_path).read_bytes() == b"new"
        assert Path(result.artifacts[0].staging_path).parent == output
    else:
        assert result.success
        assert original.read_bytes() == b"old"


@pytest.mark.parametrize("fail_audit", [False, True])
def test_explicit_audit_groups_single_document_for_atomic_publication(tmp_path, monkeypatch, fail_audit):
    from docwen_core.models.result import ConversionResult
    from docwen_runtime.output.manifest import OutputManifestWriter

    source = tmp_path / "项目记录.md"
    source.write_text("body", encoding="utf-8")
    staged = tmp_path / "generated.docx"
    staged.write_bytes(b"document")
    artifacts = [ArtifactManifest("main", "primary", str(staged), staged.name, is_primary=True)]
    request = ConversionRequest(
        "audit",
        [FileRef(str(source), "markdown", "text")],
        "docx",
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
    root = output / "项目记录_20260907_180000_fromMd"
    if fail_audit:
        assert not result.success
        assert not root.exists()
    else:
        assert result.success
        assert {path.name for path in root.iterdir()} == {"项目记录_20260907_180000_fromMd.docx", "manifest.json"}
        assert [item.kind for item in result.artifacts] == ["primary", "manifest"]
