"""Focused tests split from test_conversion_service.py."""

from __future__ import annotations

from ._conversion_service_support import (
    CSV_MEDIA_TYPE,
    DOCX_MEDIA_TYPE,
    DOCX_TO_MARKDOWN_CAPABILITY_ID,
    MARKDOWN_MEDIA_TYPE,
    MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID,
    MARKDOWN_VALIDATE_CAPABILITY_ID,
    PDF_MEDIA_TYPE,
    PDF_SPLIT_EVERY_PAGE_CAPABILITY_ID,
    PDF_TO_PNG_CAPABILITY_ID,
    PNG_MEDIA_TYPE,
    PNG_TO_OCR_MARKDOWN_CAPABILITY_ID,
    TIFF_FRAMES_TO_PNG_CAPABILITY_ID,
    TIFF_MEDIA_TYPE,
    XLSX_MEDIA_TYPE,
    XLSX_TO_CSV_CAPABILITY_ID,
    Any,
    ArtifactManifest,
    ConversionPlanRequest,
    ConversionResult,
    ConversionService,
    ConversionServiceError,
    LocalInputHandle,
    Path,
    StagingOutputTarget,
    _Committer,
    _Controller,
    _request,
    filesystem_path,
    hashlib,
    os,
    pytest,
    replace,
)

pytestmark = pytest.mark.integration


@pytest.mark.parametrize(
    ("options", "artifact_kind", "media_type", "metadata"),
    [
        (
            {"recognize_text": False, "preserve_resources": True},
            "auxiliary",
            MARKDOWN_MEDIA_TYPE,
            {"ocr": True},
        ),
        (
            {"recognize_text": True, "preserve_resources": True, "ocr_placement": "main_md"},
            "auxiliary",
            MARKDOWN_MEDIA_TYPE,
            {"ocr": True},
        ),
        (
            {"recognize_text": False, "preserve_resources": False},
            "image",
            "image/png",
            {},
        ),
    ],
)
def test_docx_bundle_rejects_producer_artifacts_that_violate_fidelity_options(
    tmp_path: Path,
    options: dict[str, Any],
    artifact_kind: str,
    media_type: str,
    metadata: dict[str, Any],
) -> None:
    class _DriftingDocxController(_Controller):
        def execute_single(self, request: Any) -> ConversionResult:
            root = Path(request.output_policy.output_dir)
            primary = root / "converted.md"
            extra = root / ("unexpected.md" if artifact_kind == "auxiliary" else "unexpected.png")
            primary.write_text("# Converted\n", encoding="utf-8")
            extra.write_bytes(b"unexpected")
            return ConversionResult(
                task_id=request.request_id,
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="artifact.document",
                        kind="primary",
                        staging_path=str(primary),
                        suggested_name=primary.name,
                        media_type=MARKDOWN_MEDIA_TYPE,
                        is_primary=True,
                    ),
                    ArtifactManifest(
                        artifact_id="artifact.extra",
                        kind=artifact_kind,
                        staging_path=str(extra),
                        suggested_name=extra.name,
                        media_type=media_type,
                        metadata=metadata,
                    ),
                ],
            )

    source = tmp_path / "source.docx"
    source.write_bytes(b"docx fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_DriftingDocxController(), _Committer())
    request = _request(
        source,
        staging,
        capability_id=DOCX_TO_MARKDOWN_CAPABILITY_ID,
        media_type=DOCX_MEDIA_TYPE,
        options=options,
    )

    with pytest.raises(ConversionServiceError) as exc_info:
        service.execute_accepted(service.accept(service.plan(request).plan_id, "task.docx.drift"))

    assert exc_info.value.code == "document_fidelity_option_mismatch"
    assert list(staging.iterdir()) == []


def test_png_ocr_markdown_maps_original_resource_and_ocr_provenance(tmp_path: Path) -> None:
    class _ImageController(_Controller):
        def describe_runtime_capabilities(self) -> dict[str, Any]:
            description = super().describe_runtime_capabilities()
            description["gates"] = [
                {"id": "python.pillow", "available": True},
                {"id": "python.rapidocr", "available": True},
            ]
            route = next(
                item
                for item in description["sources"][0]["routes"]
                if item["id"] == "docwen_plugin_image:image:md:convert"
            )
            route.update(
                {
                    "required_capabilities": ["python.pillow"],
                    "optional_capabilities": ["python.pillow_heif", "python.rapidocr"],
                    "limitations": ["HEIC requires pillow-heif", "OCR requires RapidOCR"],
                }
            )
            return description

        def execute_single(self, request: Any) -> ConversionResult:
            assert request.options == {
                "image_mode": "file",
                "to_md_keep_images": True,
                "to_md_enable_ocr": True,
                "ocr_placement": "image_md",
            }
            root = Path(request.output_policy.output_dir)
            markdown = root / "source.md"
            retained = root / "source.png"
            sidecar = root / "source_ocr.md"
            markdown.write_text("![[source_ocr.md]]\n", encoding="utf-8")
            retained.write_bytes(b"png fixture")
            sidecar.write_text("OCR text\n", encoding="utf-8")
            return ConversionResult(
                task_id=request.request_id,
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="artifact.document",
                        kind="primary",
                        staging_path=str(markdown),
                        suggested_name=markdown.name,
                        media_type=MARKDOWN_MEDIA_TYPE,
                        is_primary=True,
                    ),
                    ArtifactManifest(
                        artifact_id="artifact.original",
                        kind="image",
                        staging_path=str(retained),
                        suggested_name=retained.name,
                        media_type=PNG_MEDIA_TYPE,
                    ),
                    ArtifactManifest(
                        artifact_id="artifact.ocr",
                        kind="auxiliary",
                        staging_path=str(sidecar),
                        suggested_name=sidecar.name,
                        media_type=MARKDOWN_MEDIA_TYPE,
                        metadata={"ocr": True},
                    ),
                ],
            )

    source = tmp_path / "source.png"
    source.write_bytes(b"png input")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_ImageController(), _Committer())

    capability = {item.capability_id: item for item in service.list_capabilities()}[PNG_TO_OCR_MARKDOWN_CAPABILITY_ID]
    assert capability.availability == "available"
    assert capability.dependencies == (
        {"dependency_id": "python.pillow", "required": True, "available": True},
        {"dependency_id": "python.rapidocr", "required": True, "available": True},
    )
    assert capability.limitations == ()

    request = _request(
        source,
        staging,
        capability_id=PNG_TO_OCR_MARKDOWN_CAPABILITY_ID,
        media_type=PNG_MEDIA_TYPE,
    )
    plan = service.plan(request)
    outcome = service.execute_accepted(service.accept(plan.plan_id, "task.image-ocr"))

    assert outcome.bundle is not None
    assert [artifact.kind for artifact in outcome.bundle.artifacts] == ["document", "resource", "fragment"]
    assert [(relation.type, relation.role) for relation in outcome.bundle.relations] == [
        ("resource_of", "original"),
        ("fragment_of", "ocr_text"),
        ("derived_from", "source"),
    ]


@pytest.mark.parametrize(
    (
        "capability_id",
        "input_media_type",
        "target_format",
        "artifact_kind",
        "media_type",
        "metadata_key",
        "output_kind",
        "roles",
    ),
    [
        (
            XLSX_TO_CSV_CAPABILITY_ID,
            XLSX_MEDIA_TYPE,
            "csv",
            "auxiliary",
            CSV_MEDIA_TYPE,
            "sheet_index",
            "resource",
            ["worksheet", "worksheet"],
        ),
        (
            PDF_TO_PNG_CAPABILITY_ID,
            PDF_MEDIA_TYPE,
            "png",
            "image",
            PNG_MEDIA_TYPE,
            "page",
            "resource",
            ["image", "image"],
        ),
        (
            PDF_SPLIT_EVERY_PAGE_CAPABILITY_ID,
            PDF_MEDIA_TYPE,
            "pdf",
            "primary",
            PDF_MEDIA_TYPE,
            "page",
            "document",
            ["section", "section"],
        ),
    ],
)
def test_ordered_peer_outputs_become_resource_entries(
    tmp_path: Path,
    capability_id: str,
    input_media_type: str,
    target_format: str,
    artifact_kind: str,
    media_type: str,
    metadata_key: str,
    output_kind: str,
    roles: list[str],
) -> None:
    class _OrderedController(_Controller):
        def execute_single(self, request: Any) -> ConversionResult:
            assert request.target_format == target_format
            root = Path(request.output_policy.output_dir)
            artifacts = []
            for index in range(2):
                output = root / f"output-{index}.{target_format}"
                output.write_bytes(f"artifact {index}".encode())
                artifacts.append(
                    ArtifactManifest(
                        artifact_id=f"artifact.{index}",
                        kind="primary" if index == 0 and artifact_kind == "auxiliary" else artifact_kind,
                        staging_path=str(output),
                        suggested_name=output.name,
                        media_type=media_type,
                        metadata={metadata_key: index + 1 if metadata_key == "page" else index},
                        is_primary=index == 0,
                    )
                )
            return ConversionResult(task_id=request.request_id, success=True, artifacts=artifacts)

    source = tmp_path / ("source.xlsx" if capability_id == XLSX_TO_CSV_CAPABILITY_ID else "source.pdf")
    source.write_bytes(b"fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_OrderedController(), _Committer())
    request = _request(
        source,
        staging,
        capability_id=capability_id,
        media_type=input_media_type,
    )
    outcome = service.execute_accepted(service.accept(service.plan(request).plan_id, "task.ordered"))

    assert outcome.bundle is not None
    assert [artifact.kind for artifact in outcome.bundle.artifacts] == [output_kind, output_kind]
    assert [entry.role for entry in outcome.bundle.entries] == roles
    assert [entry.ordinal for entry in outcome.bundle.entries] == [0, 1]
    assert [entry.preferred for entry in outcome.bundle.entries] == [True, False]


@pytest.mark.parametrize(
    ("capability_id", "input_media_type", "metadata_key", "entry_role"),
    [
        (MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID, MARKDOWN_MEDIA_TYPE, "table_index", "supplementary"),
        (TIFF_FRAMES_TO_PNG_CAPABILITY_ID, TIFF_MEDIA_TYPE, "page_index", "image"),
    ],
)
def test_markdown_tables_and_tiff_frames_keep_route_defined_order(
    tmp_path: Path,
    capability_id: str,
    input_media_type: str,
    metadata_key: str,
    entry_role: str,
) -> None:
    class _PeerController(_Controller):
        def describe_runtime_capabilities(self) -> dict[str, Any]:
            description = super().describe_runtime_capabilities()
            description["gates"] = [
                {"id": "python.mistune", "available": True},
                {"id": "python.openpyxl", "available": True},
                {"id": "python.pillow", "available": True},
            ]
            for route in description["sources"][0]["routes"]:
                if route["id"] == "docwen_plugin_markdown:markdown:csv:convert":
                    route["required_capabilities"] = ["python.mistune", "python.openpyxl"]
                if route["id"] == "docwen_plugin_image:image:png:convert":
                    route["required_capabilities"] = ["python.pillow"]
            return description

        def execute_single(self, request: Any) -> ConversionResult:
            target_format = "csv" if capability_id == MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID else "png"
            media_type = CSV_MEDIA_TYPE if target_format == "csv" else PNG_MEDIA_TYPE
            root = Path(request.output_policy.output_dir)
            artifacts = []
            for index in range(2):
                output = root / f"output-{index}.{target_format}"
                output.write_bytes(f"output {index}".encode())
                artifacts.append(
                    ArtifactManifest(
                        artifact_id=f"artifact.{index}",
                        kind="primary" if target_format == "csv" or index == 0 else "auxiliary",
                        staging_path=str(output),
                        suggested_name=output.name,
                        media_type=media_type,
                        metadata={metadata_key: index},
                        is_primary=index == 0,
                    )
                )
            return ConversionResult(task_id=request.request_id, success=True, artifacts=artifacts)

    source = tmp_path / ("source.md" if input_media_type == MARKDOWN_MEDIA_TYPE else "source.tiff")
    source.write_bytes(b"fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_PeerController(), _Committer())
    request = _request(
        source,
        staging,
        capability_id=capability_id,
        media_type=input_media_type,
    )
    outcome = service.execute_accepted(service.accept(service.plan(request).plan_id, f"task.{metadata_key}"))

    assert outcome.bundle is not None
    assert [artifact.kind for artifact in outcome.bundle.artifacts] == ["resource", "resource"]
    assert [(entry.role, entry.ordinal, entry.preferred) for entry in outcome.bundle.entries] == [
        (entry_role, 0, True),
        (entry_role, 1, False),
    ]


def test_plan_rejects_integrity_mismatch_and_nonempty_staging(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("body", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())

    bad = _request(source, staging)
    bad = replace(bad, inputs=(replace(bad.inputs[0], sha256="0" * 64), bad.inputs[1]))
    with pytest.raises(ConversionServiceError, match="declared input") as mismatch:
        service.plan(bad)
    assert mismatch.value.code == "input_integrity_mismatch"

    (staging / "foreign.txt").write_text("occupied", encoding="utf-8")
    with pytest.raises(ConversionServiceError, match="must be empty") as occupied:
        service.plan(_request(source, staging))
    assert occupied.value.code == "staging_root_not_empty"


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_plan_and_accept_support_ordinary_windows_long_paths(tmp_path: Path) -> None:
    long_root = tmp_path
    while len(str(long_root / "source.md")) < 280:
        long_root /= "long-machine-plan-segment-0123456789"
    source = long_root / "source.md"
    staging = long_root / "staging"
    io_source = filesystem_path(source, force_extended=True)
    io_source.parent.mkdir(parents=True)
    payload = b"# long Machine input\n"
    io_source.write_bytes(payload)
    filesystem_path(staging, force_extended=True).mkdir()
    request = ConversionPlanRequest(
        capability_id=MARKDOWN_VALIDATE_CAPABILITY_ID,
        inputs=(
            LocalInputHandle(
                input_id="input.long",
                path=str(source),
                kind="document",
                role="source",
                logical_path="documents/source.md",
                media_type=MARKDOWN_MEDIA_TYPE,
                size_bytes=len(payload),
                sha256=hashlib.sha256(payload).hexdigest(),
            ),
        ),
        output=StagingOutputTarget(staging_root=str(staging)),
    )
    service = ConversionService(_Controller(), _Committer())

    plan = service.plan(request)
    task_id = service.accept(plan.plan_id, "task.long-input")

    assert task_id == "task.long-input"
    assert not request.inputs[0].path.startswith("\\\\?\\")


def test_plan_rejects_symlink_input_without_reading_target(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("sentinel", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    request = _request(source, staging)
    target = Path(request.inputs[0].path)
    target_bytes = target.read_bytes()
    link = tmp_path / "neutral-link.json"
    try:
        link.symlink_to(target)
    except OSError as exc:
        pytest.skip(f"file symlink unavailable: {exc}")
    service = ConversionService(_Controller(), _Committer())

    with pytest.raises(ConversionServiceError) as rejected:
        service.plan(
            replace(
                request,
                inputs=(replace(request.inputs[0], path=str(link)), request.inputs[1]),
            )
        )

    assert rejected.value.category == "security"
    assert rejected.value.code == "input_is_link"
    assert target.read_bytes() == target_bytes


def test_execute_rejects_input_replaced_by_link_before_reading(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "source.md"
    source.write_text("before", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _Controller()
    service = ConversionService(controller, _Committer())
    request = _request(source, staging)
    neutral_path = filesystem_path(request.inputs[0].path, force_extended=os.name == "nt")
    task_id = service.accept(service.plan(request).plan_id, "task.link-replaced")
    original_guard = service._path_traverses_link_or_junction

    monkeypatch.setattr(
        service,
        "_path_traverses_link_or_junction",
        lambda path: path == neutral_path or original_guard(path),
    )
    monkeypatch.setattr(
        service,
        "_file_integrity",
        lambda *_args, **_kwargs: pytest.fail("link target must not be read"),
    )

    with pytest.raises(ConversionServiceError) as rejected:
        service.execute_accepted(task_id)

    assert rejected.value.category == "security"
    assert rejected.value.code == "input_is_link"
    assert controller.released == [task_id]
    assert service.cancel(task_id) == "already_terminal"


@pytest.mark.parametrize("changed_index", [0, 1])
def test_execute_rechecks_each_dual_input_after_acceptance(tmp_path: Path, changed_index: int) -> None:
    source = tmp_path / "source.md"
    source.write_text("before", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _Controller()
    service = ConversionService(controller, _Committer())
    request = _request(source, staging)
    task_id = service.accept(service.plan(request).plan_id, "task.changed")

    Path(request.inputs[changed_index].path).write_text("{}", encoding="utf-8")
    with pytest.raises(ConversionServiceError, match="changed after planning") as changed:
        service.execute_accepted(task_id)
    assert changed.value.code == "input_changed_after_plan"
    assert controller.released == [task_id]
    assert service.cancel(task_id) == "already_terminal"


def test_cancel_is_atomic_after_acceptance(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("body", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _Controller()
    service = ConversionService(controller, _Committer())
    task_id = service.accept(service.plan(_request(source, staging)).plan_id, "task.cancel")

    assert service.cancel(task_id) == "cancellation_requested"
    assert controller.cancelled == [task_id]
    assert service.cancel("task.unknown") == "not_found"
