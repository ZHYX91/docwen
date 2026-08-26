"""Request-scoped output manifest persistence contracts."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.artifact import ARTIFACT_KIND_MANIFEST, ArtifactManifest
from docwen_core.models.conversion_manifest import ConversionManifestContext, PreconversionStep
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.result import ConversionErrorInfo, ConversionResult
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.output.manifest import OutputManifestDocument, OutputManifestWriter, canonical_manifest_bytes

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _isolate_manifest_staging(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    staging_root = tmp_path / "private-temp"
    staging_root.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(staging_root))


def _request(
    tmp_path: Path,
    *,
    mask_input_path: object = True,
    write_artifacts: bool = True,
    save_to_output: object = True,
) -> ConversionRequest:
    source = tmp_path / "private" / "secret.csv"
    source.parent.mkdir(exist_ok=True)
    source.write_text("city,value\n北京,1\n", encoding="utf-8")
    snapshot = {
        "output": {
            "manifest": {
                "save_to_output": save_to_output,
                "mask_input_path": mask_input_path,
            }
        }
    }
    ref = FileRef(path=str(source.resolve()), format="csv", category="spreadsheet")
    context = ConversionManifestContext.from_request_inputs([ref], snapshot).with_step(
        PreconversionStep(
            input_index=0,
            source_format="csv",
            target_format="xlsx",
            status="completed",
            backend="isolated-test",
        )
    )
    return ConversionRequest(
        request_id="unsafe/../task-id",
        input_refs=[ref],
        target_format="md",
        output_policy=OutputPolicy(
            output_dir=str(tmp_path / "output"),
            write_artifacts=write_artifacts,
        ),
        config_snapshot=snapshot,
        manifest_context=context,
    )


def _success_result(tmp_path: Path) -> ConversionResult:
    output = tmp_path / "output" / "result.md"
    output.parent.mkdir(exist_ok=True)
    output.write_text("| city | value |\n", encoding="utf-8")
    return ConversionResult(
        task_id="unsafe/../task-id",
        success=True,
        artifacts=[
            ArtifactManifest(
                artifact_id="primary",
                kind="primary",
                staging_path=str(output),
                suggested_name=output.name,
                media_type="text/markdown",
                is_primary=True,
            )
        ],
    )


def test_success_manifest_is_masked_canonical_and_not_self_referential(tmp_path: Path) -> None:
    request = _request(tmp_path)
    result = OutputManifestWriter(OutputFinalizer()).persist(request, _success_result(tmp_path))

    assert result.success is True
    assert [item.kind for item in result.artifacts] == ["primary", ARTIFACT_KIND_MANIFEST]
    manifest_path = Path(result.artifacts[-1].staging_path)
    assert manifest_path.name == "manifest.json"
    raw = manifest_path.read_bytes()
    assert raw.startswith(b"{") and raw.endswith(b"\n")
    assert b"\xef\xbb\xbf" not in raw
    assert str(request.input_refs[0].path).encode() not in raw
    payload = json.loads(raw)
    assert payload["inputs"] == [{"path": "<redacted>/secret.csv", "format": "csv", "category": "spreadsheet"}]
    assert payload["preconversion"][0]["backend"] == "isolated-test"
    assert payload["artifacts"] == [
        {"kind": "primary", "name": "result.md", "media_type": "text/markdown", "is_primary": True}
    ]
    assert all(item["kind"] != ARTIFACT_KIND_MANIFEST for item in payload["artifacts"])

    repeated = OutputManifestWriter(OutputFinalizer()).persist(request, result)
    assert repeated is result
    assert sorted(path.name for path in manifest_path.parent.glob("manifest*.json")) == ["manifest.json"]


def test_runtime_adapter_persists_manifest_after_canonical_admission(tmp_path: Path) -> None:
    from docwen_runtime.adapters import RuntimePortAdapter

    request = _request(tmp_path)
    expected = _success_result(tmp_path)

    class _Manager:
        def execute_single(self, admitted: ConversionRequest, *, on_event: object = None) -> ConversionResult:
            del on_event
            assert admitted.input_refs[0].format == "csv"
            assert admitted.manifest_context is not None
            return expected

        def execute_batch(self, request: ConversionRequest, *, on_event: object = None) -> list[ConversionResult]:
            raise AssertionError((request, on_event))

        def cancel(self, task_id: str) -> None:
            del task_id

        def reserve_cancellation(self, task_id: str) -> None:
            del task_id

        def release_cancellation(self, task_id: str) -> None:
            del task_id

        def cancel_all(self) -> None:
            return None

    manager: Any = _Manager()
    adapter = RuntimePortAdapter(
        manager,
        output_manifest_writer=OutputManifestWriter(OutputFinalizer()),
    )

    result = adapter.execute(request)

    assert result.success is True
    assert result.artifacts[-1].kind == ARTIFACT_KIND_MANIFEST
    assert Path(result.artifacts[-1].staging_path).is_file()


def test_unmasked_policy_exposes_only_the_original_input_path(tmp_path: Path) -> None:
    request = _request(tmp_path, mask_input_path=False)
    result = OutputManifestWriter(OutputFinalizer()).persist(request, _success_result(tmp_path))
    payload = json.loads(Path(result.artifacts[-1].staging_path).read_text(encoding="utf-8"))

    assert payload["inputs"][0]["path"] == request.input_refs[0].path
    serialized = json.dumps(payload, ensure_ascii=False)
    assert "private-temp" not in serialized


def test_resource_first_manifest_context_anchors_and_serializes_only_source(tmp_path: Path) -> None:
    request = _request(tmp_path)
    source = request.input_refs[0]
    resource = tmp_path / "private" / "bibliography.json"
    resource.write_text('{"schema":"docwen.semantic_bibliography.v1","entries":[]}', encoding="utf-8")
    resource_ref = FileRef(
        path=str(resource.resolve()),
        format="resource",
        category="other",
        input_kind="resource",
        input_role="bibliography",
        logical_path="bibliography.json",
        media_type="application/vnd.docwen.semantic-bibliography+json",
    )
    context = ConversionManifestContext.from_request_inputs(
        [resource_ref, source],
        request.config_snapshot,
    )
    request.input_refs = [resource_ref, source]
    request.manifest_context = context

    result = OutputManifestWriter(OutputFinalizer()).persist(request, _success_result(tmp_path))

    manifest_path = Path(result.artifacts[-1].staging_path)
    assert manifest_path.parent == (tmp_path / "output").resolve()
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert payload["inputs"] == [{"path": "<redacted>/secret.csv", "format": "csv", "category": "spreadsheet"}]


def test_failure_manifest_uses_hashed_filename_and_preserves_original_error(tmp_path: Path) -> None:
    request = _request(tmp_path)
    original = ConversionResult(
        task_id="../../unsafe-task",
        success=False,
        error=ConversionErrorInfo(
            error_type="conversion_failed",
            message=f"private path: {request.input_refs[0].path}",
            diagnostic_code="PLUGIN_FAILED",
        ),
    )

    persisted = OutputManifestWriter(OutputFinalizer()).persist(request, original)

    assert persisted.success is False
    assert persisted.error is original.error
    manifest = Path(persisted.artifacts[-1].staging_path)
    assert manifest.parent == (tmp_path / "output").resolve()
    assert manifest.name.startswith("manifest_failed_")
    assert ".." not in manifest.name and "/" not in manifest.name
    raw = manifest.read_text(encoding="utf-8")
    assert request.input_refs[0].path not in raw
    assert json.loads(raw)["error"] == {
        "type": "conversion_failed",
        "diagnostic_code": "PLUGIN_FAILED",
    }


def test_deep_temp_root_does_not_drop_success_or_failure_manifest(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    deep_temp = tmp_path / "deep-temp"
    prospective_lock = deep_temp / "docwen-output-finalizer-locks" / f"{'0' * 64}.lock"
    while len(str(prospective_lock)) < 270:
        deep_temp /= "nested-temp"
        prospective_lock = deep_temp / "docwen-output-finalizer-locks" / f"{'0' * 64}.lock"
    deep_temp.mkdir(parents=True)
    monkeypatch.setattr(tempfile, "tempdir", str(deep_temp))

    request = _request(tmp_path)
    writer = OutputManifestWriter(OutputFinalizer())
    success = writer.persist(request, _success_result(tmp_path))
    failure = writer.persist(
        request,
        ConversionResult(
            task_id="deep-temp-failure",
            success=False,
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message="failed",
                diagnostic_code="PLUGIN_FAILED",
            ),
        ),
    )

    assert prospective_lock.is_absolute()
    assert len(str(prospective_lock)) >= 270
    assert [item.kind for item in success.artifacts] == ["primary", ARTIFACT_KIND_MANIFEST]
    assert failure.artifacts[-1].kind == ARTIFACT_KIND_MANIFEST
    assert Path(failure.artifacts[-1].staging_path).name.startswith("manifest_failed_")
    assert all(item.code != "OUTPUT_MANIFEST_WRITE_FAILED" for item in [*success.diagnostics, *failure.diagnostics])


def test_collision_renames_without_overwriting_existing_manifest(tmp_path: Path) -> None:
    request = _request(tmp_path)
    writer = OutputManifestWriter(OutputFinalizer())
    first = writer.persist(request, _success_result(tmp_path))
    first_path = Path(first.artifacts[-1].staging_path)
    first_bytes = first_path.read_bytes()

    second = writer.persist(request, _success_result(tmp_path))
    second_path = Path(second.artifacts[-1].staging_path)

    assert first_path.name == "manifest.json"
    assert second_path.name == "manifest_001.json"
    assert first_path.read_bytes() == first_bytes
    assert second_path.read_bytes() == first_bytes


@pytest.mark.parametrize(
    ("save_to_output", "write_artifacts"),
    [(False, True), ("true", True), (True, False)],
)
def test_disabled_or_dry_run_policy_writes_nothing(
    tmp_path: Path,
    save_to_output: object,
    write_artifacts: bool,
) -> None:
    request = _request(
        tmp_path,
        save_to_output=save_to_output,
        write_artifacts=write_artifacts,
    )
    original = _success_result(tmp_path)

    persisted = OutputManifestWriter(OutputFinalizer()).persist(request, original)

    assert persisted is original
    assert list((tmp_path / "output").glob("manifest*.json")) == []


def test_cancellation_writes_nothing(tmp_path: Path) -> None:
    request = _request(tmp_path)
    original = ConversionResult(
        task_id="cancelled",
        success=False,
        error=ConversionErrorInfo(error_type="cancelled", message="cancelled"),
    )

    persisted = OutputManifestWriter(OutputFinalizer()).persist(request, original)

    assert persisted is original
    assert list((tmp_path / "output").glob("manifest*.json")) == []


def test_writer_failure_adds_warning_without_changing_terminal_truth(tmp_path: Path) -> None:
    class _FailingFinalizer(OutputFinalizer):
        def resolve_output_dir(self, policy: OutputPolicy, input_path: str = "") -> str:
            del policy, input_path
            raise OSError("blocked")

    request = _request(tmp_path)
    original = _success_result(tmp_path)

    persisted = OutputManifestWriter(_FailingFinalizer()).persist(request, original)

    assert persisted.success is True
    assert persisted.error is None
    assert persisted.artifacts == original.artifacts
    assert persisted.diagnostics[-1].code == "OUTPUT_MANIFEST_WRITE_FAILED"


def test_canonical_bytes_are_repeatable_for_unicode_document() -> None:
    document = OutputManifestDocument(
        request_id_sha256="a" * 64,
        status="success",
        target_format="md",
        action_name="",
        inputs=(("<redacted>/北京.csv", "csv", "spreadsheet"),),
        preconversion=(),
        artifacts=(("primary", "结果.md", "text/markdown", True),),
        diagnostics=(("info", "DONE"),),
        error=None,
    )

    first = canonical_manifest_bytes(document)
    second = canonical_manifest_bytes(document)

    assert first == second
    assert "北京" in first.decode("utf-8")
