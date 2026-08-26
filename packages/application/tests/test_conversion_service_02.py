"""Focused tests split from test_conversion_service.py."""

from __future__ import annotations

from ._conversion_service_support import (
    _EXPECTED_DOCUMENT_SEMANTICS_MACHINE_LIMITATIONS,
    _EXPECTED_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS,
    DOCX_MEDIA_TYPE,
    DOCX_TO_MARKDOWN_CAPABILITY_ID,
    MARKDOWN_MEDIA_TYPE,
    MARKDOWN_TO_DOCX_CAPABILITY_ID,
    OFD_MEDIA_TYPE,
    OFD_TO_MARKDOWN_CAPABILITY_ID,
    PDF_MEDIA_TYPE,
    PDF_SPLIT_CUSTOM_CAPABILITY_ID,
    PDF_TO_MARKDOWN_CAPABILITY_ID,
    TIFF_MEDIA_TYPE,
    TIFF_TO_MARKDOWN_CAPABILITY_ID,
    XPS_MEDIA_TYPE,
    XPS_TO_MARKDOWN_CAPABILITY_ID,
    Any,
    ArtifactManifest,
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


def test_custom_pdf_partition_publishes_disjoint_section_documents(tmp_path: Path) -> None:
    class _PartitionController(_Controller):
        def execute_single(self, request: Any) -> ConversionResult:
            root = Path(request.output_policy.output_dir)
            artifacts = []
            for index, pages in enumerate(([1, 3], [2, 4])):
                output = root / f"part-{index + 1}.pdf"
                output.write_bytes(f"part {index + 1}".encode())
                artifacts.append(
                    ArtifactManifest(
                        artifact_id=f"artifact.part.{index + 1}",
                        kind="primary",
                        staging_path=str(output),
                        suggested_name=output.name,
                        media_type=PDF_MEDIA_TYPE,
                        metadata={"pages": pages, "split_mode": "custom"},
                        is_primary=index == 0,
                    )
                )
            return ConversionResult(task_id=request.request_id, success=True, artifacts=artifacts)

    source = tmp_path / "source.pdf"
    source.write_bytes(b"pdf")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_PartitionController(), _Committer())
    request = _request(
        source,
        staging,
        capability_id=PDF_SPLIT_CUSTOM_CAPABILITY_ID,
        media_type=PDF_MEDIA_TYPE,
        options={"pages": [1, 3]},
    )

    plan = service.plan(request)
    assert plan.effective_options == {"split_mode": "custom", "pages": [1, 3]}
    outcome = service.execute_accepted(service.accept(plan.plan_id, "task.partition"))

    assert outcome.bundle is not None
    assert [artifact.kind for artifact in outcome.bundle.artifacts] == ["document", "document"]
    assert [(entry.role, entry.ordinal, entry.preferred) for entry in outcome.bundle.entries] == [
        ("section", 0, True),
        ("section", 1, False),
    ]


def test_docx_to_markdown_declares_document_resource_graph_and_preserves_images(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _Controller()
    service = ConversionService(controller, _Committer())

    capabilities = {item.capability_id: item for item in service.list_capabilities()}
    capability = capabilities[DOCX_TO_MARKDOWN_CAPABILITY_ID]
    assert capability.output_media_types == (MARKDOWN_MEDIA_TYPE,)
    assert capability.output_shape.to_dict() == {
        "cardinality": "many",
        "artifact_kinds": ["document", "fragment", "resource"],
        "relation_types": ["fragment_of", "resource_of"],
        "atomic_bundle": True,
    }
    assert capability.limitations == _EXPECTED_DOCUMENT_SEMANTICS_MACHINE_LIMITATIONS
    assert capability.options_schema["required"] == []
    assert capability.options_schema["additionalProperties"] is False
    assert capability.options_schema["properties"]["recognize_text"] == {
        "type": "boolean",
        "default": False,
    }
    assert capability.options_schema["properties"]["preserve_resources"] == {
        "type": "boolean",
        "default": True,
    }
    assert set(capability.options_schema["properties"]) == {
        "recognize_text",
        "preserve_resources",
        "ocr_language",
        "image_mode",
        "ocr_placement",
        "image_link_style",
        "table_merge_strategy",
        "remove_numbering",
        "add_numbering",
        "numbering_scheme",
    }
    assert "to_md_enable_ocr" not in capability.options_schema["properties"]
    assert "to_md_keep_images" not in capability.options_schema["properties"]

    request = _request(
        source,
        staging,
        capability_id=DOCX_TO_MARKDOWN_CAPABILITY_ID,
        media_type=DOCX_MEDIA_TYPE,
    )
    plan = service.plan(request)
    assert plan.limitations == _EXPECTED_DOCUMENT_SEMANTICS_MACHINE_LIMITATIONS
    assert plan.effective_options == {
        "add_numbering": False,
        "image_mode": "file",
        "image_link_style": "wiki_embed",
        "numbering_scheme": "gongwen_standard",
        "ocr_language": "auto",
        "ocr_placement": "main_md",
        "preserve_resources": True,
        "recognize_text": False,
        "remove_numbering": True,
        "table_merge_strategy": "fill",
    }
    outcome = service.execute_accepted(service.accept(plan.plan_id, "task.docx-to-md"))
    assert outcome.bundle is not None
    assert outcome.bundle.artifacts[0].media_type == MARKDOWN_MEDIA_TYPE
    assert controller.requests[-1].options["to_md_enable_ocr"] is False
    assert controller.requests[-1].options["to_md_keep_images"] is True
    assert "recognize_text" not in controller.requests[-1].options
    assert "preserve_resources" not in controller.requests[-1].options


@pytest.mark.parametrize(
    ("capability_id", "media_type", "expected_defaults"),
    [
        (
            PDF_TO_MARKDOWN_CAPABILITY_ID,
            PDF_MEDIA_TYPE,
            {"preserve_resources": True, "recognize_text": False},
        ),
        (
            OFD_TO_MARKDOWN_CAPABILITY_ID,
            OFD_MEDIA_TYPE,
            {"preserve_resources": True, "recognize_text": False},
        ),
        (
            XPS_TO_MARKDOWN_CAPABILITY_ID,
            XPS_MEDIA_TYPE,
            {"preserve_resources": True, "recognize_text": False},
        ),
        (
            TIFF_TO_MARKDOWN_CAPABILITY_ID,
            TIFF_MEDIA_TYPE,
            {"preserve_resources": True, "recognize_text": False},
        ),
    ],
)
def test_physical_page_machine_capabilities_are_consumer_neutral(
    tmp_path: Path,
    capability_id: str,
    media_type: str,
    expected_defaults: dict[str, bool],
) -> None:
    source = tmp_path / f"source-{capability_id.rsplit('.', 2)[1]}"
    source.write_bytes(b"fixture")
    staging = tmp_path / f"staging-{capability_id.rsplit('.', 2)[1]}"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())
    capability = {item.capability_id: item for item in service.list_capabilities()}[capability_id]

    assert capability.output_shape.to_dict() == {
        "cardinality": "many",
        "artifact_kinds": ["document", "fragment", "resource"],
        "relation_types": ["fragment_of", "resource_of"],
        "atomic_bundle": True,
        "relation_payloads": ["page_fragment", "page_resource"],
    }
    serialized = capability.to_dict()
    assert "page_nodes" not in repr(serialized)
    assert "pkwf" not in repr(serialized).lower()
    schema = capability.options_schema
    assert schema["required"] == []
    assert schema["additionalProperties"] is False
    assert schema["properties"]["recognize_text"] == {"type": "boolean", "default": False}
    assert schema["properties"]["preserve_resources"] == {"type": "boolean", "default": True}
    assert "to_md_enable_ocr" not in schema["properties"]
    assert "to_md_keep_images" not in schema["properties"]
    expected_keys = {"recognize_text", "preserve_resources", "ocr_language"}
    if capability_id != TIFF_TO_MARKDOWN_CAPABILITY_ID:
        expected_keys.update({"image_mode", "render_dpi"})
    assert set(schema["properties"]) == expected_keys
    plan = service.plan(_request(source, staging, capability_id=capability_id, media_type=media_type))
    for key, value in expected_defaults.items():
        assert plan.effective_options[key] is value


@pytest.mark.parametrize(("option_name", "producer_value"), [("recognize_text", False), ("preserve_resources", False)])
def test_physical_page_machine_rejects_producer_option_drift(
    tmp_path: Path,
    option_name: str,
    producer_value: bool,
) -> None:
    class _DriftingController(_Controller):
        def execute_single(self, request: Any) -> ConversionResult:
            self.requests.append(request)
            output = Path(request.output_policy.output_dir) / "primary.md"
            output.write_text("# primary\n", encoding="utf-8")
            metadata = {
                "physical_page_count": 1,
                "ocr_enabled": request.options["to_md_enable_ocr"],
                "keep_images": request.options["to_md_keep_images"],
            }
            metadata["ocr_enabled" if option_name == "recognize_text" else "keep_images"] = producer_value
            return ConversionResult(
                task_id=request.request_id,
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id="artifact.primary",
                        kind="primary",
                        staging_path=str(output),
                        suggested_name="primary.md",
                        media_type=MARKDOWN_MEDIA_TYPE,
                        metadata=metadata,
                        is_primary=True,
                    )
                ],
            )

    source = tmp_path / "source.pdf"
    source.write_bytes(b"fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    options = {
        "recognize_text": option_name == "recognize_text",
        "preserve_resources": True,
    }
    service = ConversionService(_DriftingController(), _Committer())
    plan = service.plan(
        _request(
            source,
            staging,
            capability_id=PDF_TO_MARKDOWN_CAPABILITY_ID,
            media_type=PDF_MEDIA_TYPE,
            options=options,
        )
    )

    with pytest.raises(ConversionServiceError) as exc_info:
        service.execute_accepted(service.accept(plan.plan_id, f"task.option-drift.{option_name}"))

    assert exc_info.value.code == "physical_page_option_mismatch"
    assert list(staging.iterdir()) == []


@pytest.mark.parametrize(
    "capability_id,media_type",
    [
        (DOCX_TO_MARKDOWN_CAPABILITY_ID, DOCX_MEDIA_TYPE),
        (PDF_TO_MARKDOWN_CAPABILITY_ID, PDF_MEDIA_TYPE),
    ],
)
@pytest.mark.parametrize("recognize_text", [False, True])
@pytest.mark.parametrize("preserve_resources", [False, True])
def test_document_routes_translate_all_four_public_fidelity_combinations_at_runtime(
    tmp_path: Path,
    capability_id: str,
    media_type: str,
    recognize_text: bool,
    preserve_resources: bool,
) -> None:
    class _CapturingController(_Controller):
        def __init__(self) -> None:
            super().__init__()
            self.prepared_requests: list[Any] = []

        def prepare_execution_cancellation(self, request: Any, *, batch: bool = False) -> object:
            self.prepared_requests.append(request)
            return super().prepare_execution_cancellation(request, batch=batch)

    source = tmp_path / f"source-{capability_id}.bin"
    source.write_bytes(b"fixture")
    staging = tmp_path / f"staging-{capability_id}-{recognize_text}-{preserve_resources}"
    staging.mkdir()
    controller = _CapturingController()
    service = ConversionService(controller, _Committer())
    public_options = {
        "recognize_text": recognize_text,
        "preserve_resources": preserve_resources,
    }

    plan = service.plan(
        _request(
            source,
            staging,
            capability_id=capability_id,
            media_type=media_type,
            options=public_options,
        )
    )
    assert {key: plan.effective_options[key] for key in public_options} == public_options
    service.accept(plan.plan_id, f"task.fidelity.{capability_id}.{recognize_text}.{preserve_resources}")

    runtime_options = controller.prepared_requests[-1].options
    assert runtime_options["to_md_enable_ocr"] is recognize_text
    assert runtime_options["to_md_keep_images"] is preserve_resources
    assert "recognize_text" not in runtime_options
    assert "preserve_resources" not in runtime_options


@pytest.mark.parametrize(
    ("capability_id", "media_type"),
    [
        (DOCX_TO_MARKDOWN_CAPABILITY_ID, DOCX_MEDIA_TYPE),
        (PDF_TO_MARKDOWN_CAPABILITY_ID, PDF_MEDIA_TYPE),
    ],
)
@pytest.mark.parametrize("legacy_key", ["to_md_enable_ocr", "to_md_keep_images"])
def test_document_routes_reject_legacy_fidelity_keys_as_public_machine_options(
    tmp_path: Path,
    capability_id: str,
    media_type: str,
    legacy_key: str,
) -> None:
    source = tmp_path / f"source-{capability_id}.bin"
    source.write_bytes(b"fixture")
    staging = tmp_path / f"staging-{capability_id}-{legacy_key}"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())

    with pytest.raises(ConversionServiceError) as exc_info:
        service.plan(
            _request(
                source,
                staging,
                capability_id=capability_id,
                media_type=media_type,
                options={legacy_key: True},
            )
        )

    assert exc_info.value.code == "unsupported_options"
    assert exc_info.value.details == {"option_keys": [legacy_key]}


def test_capability_discovery_projects_runtime_dependencies_and_fails_closed(tmp_path: Path) -> None:
    class _UnavailableController(_Controller):
        def describe_runtime_capabilities(self) -> dict[str, Any]:
            return {
                "gates": [{"id": "python.docx", "available": False}],
                "sources": [
                    {
                        "routes": [
                            {
                                "id": "docwen_plugin_markdown:markdown:docx:convert",
                                "available": False,
                                "required_capabilities": ["python.docx"],
                                "optional_capabilities": [],
                                "missing_required_capabilities": ["python.docx"],
                                "missing_optional_capabilities": [],
                                "limitations": ["python-docx is required"],
                            }
                        ]
                    }
                ],
            }

    source = tmp_path / "source.md"
    source.write_text("# Source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_UnavailableController(), _Committer())

    capability = {item.capability_id: item for item in service.list_capabilities()}[MARKDOWN_TO_DOCX_CAPABILITY_ID]
    assert capability.availability == "unavailable"
    assert capability.dependencies == ({"dependency_id": "python.docx", "required": True, "available": False},)
    assert capability.limitations == (
        *_EXPECTED_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS,
        {
            "severity": "warning",
            "code": "runtime_route_limitation",
            "message": "python-docx is required",
        },
    )

    with pytest.raises(ConversionServiceError) as unavailable:
        service.plan(_request(source, staging))
    assert unavailable.value.code == "capability_unavailable"
    assert unavailable.value.details == {"missing_required_dependencies": ["python.docx"]}


def test_docx_to_markdown_maps_images_and_ocr_sidecars_without_leaking_technical_kinds(tmp_path: Path) -> None:
    class _MultiArtifactController(_Controller):
        def execute_single(self, request: Any) -> ConversionResult:
            root = Path(request.output_policy.output_dir)
            markdown = root / "converted.md"
            image = root / "image1.png"
            ocr = root / "image1.md"
            markdown.write_text("# Converted\n\n![](image1.png)\n", encoding="utf-8")
            image.write_bytes(b"png fixture")
            ocr.write_text("OCR text\n", encoding="utf-8")
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
                        artifact_id="artifact.image",
                        kind="image",
                        staging_path=str(image),
                        suggested_name=image.name,
                        media_type="image/png",
                    ),
                    ArtifactManifest(
                        artifact_id="artifact.ocr",
                        kind="auxiliary",
                        staging_path=str(ocr),
                        suggested_name=ocr.name,
                        media_type=MARKDOWN_MEDIA_TYPE,
                        metadata={"ocr": True},
                    ),
                ],
            )

    source = tmp_path / "source.docx"
    source.write_bytes(b"docx fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_MultiArtifactController(), _Committer())
    request = _request(
        source,
        staging,
        capability_id=DOCX_TO_MARKDOWN_CAPABILITY_ID,
        media_type=DOCX_MEDIA_TYPE,
        options={
            "recognize_text": True,
            "preserve_resources": True,
            "ocr_placement": "image_md",
        },
    )
    outcome = service.execute_accepted(service.accept(service.plan(request).plan_id, "task.multi"))

    assert outcome.bundle is not None
    assert [artifact.kind for artifact in outcome.bundle.artifacts] == ["document", "resource", "fragment"]
    assert [(relation.type, relation.role, relation.ordinal) for relation in outcome.bundle.relations] == [
        ("resource_of", "image", 0),
        ("fragment_of", "ocr_text", 0),
    ]
