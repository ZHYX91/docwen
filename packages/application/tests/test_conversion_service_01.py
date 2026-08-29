"""Focused tests split from test_conversion_service.py."""

from __future__ import annotations

from ._conversion_service_support import (
    _EXPECTED_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS,
    IMAGES_MERGE_TO_TIFF_CAPABILITY_ID,
    MARKDOWN_MEDIA_TYPE,
    MARKDOWN_NUMBERING_CAPABILITY_ID,
    MARKDOWN_TO_DOCX_CAPABILITY_ID,
    MARKDOWN_TO_XLSX_CAPABILITY_ID,
    MARKDOWN_VALIDATE_CAPABILITY_ID,
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    PDF_MEDIA_TYPE,
    PDF_MERGE_CAPABILITY_ID,
    PNG_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
    SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE,
    XLSX_MEDIA_TYPE,
    XLSX_MERGE_TABLES_CAPABILITY_ID,
    XLSX_TO_MARKDOWN_CAPABILITY_ID,
    ConversionPlanRequest,
    ConversionService,
    ConversionServiceError,
    LocalInputHandle,
    Path,
    StagingOutputTarget,
    _Committer,
    _Controller,
    _refingerprint,
    _request,
    _request_many,
    _sha256,
    canonicalize_numbering_plan,
    hashlib,
    json,
    pytest,
    replace,
)

pytestmark = pytest.mark.integration


def test_capability_and_successful_plan_execute_closed_loop(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# Source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _Controller()
    service = ConversionService(controller, _Committer())

    capability = service.list_capabilities()[0].to_dict()
    assert capability["capability_id"] == MARKDOWN_TO_DOCX_CAPABILITY_ID
    assert capability["input_shape"] == {
        "slots": [
            {
                "role": "neutral_document",
                "kind": "document",
                "media_types": [RESOLVED_DOCUMENT_MEDIA_TYPE],
                "min_items": 1,
                "max_items": 1,
            },
            {
                "role": "numbering_export_plan",
                "kind": "resource",
                "media_types": [NUMBERING_EXPORT_PLAN_MEDIA_TYPE],
                "min_items": 1,
                "max_items": 1,
            },
        ],
        "undeclared_roles": "reject",
    }
    assert capability["output_shape"] == {
        "cardinality": "many",
        "artifact_kinds": ["document", "resource"],
        "relation_types": ["resource_of"],
        "atomic_bundle": True,
    }
    assert capability["options_schema"]["additionalProperties"] is False
    assert capability["limitations"] == list(_EXPECTED_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS)

    plan = service.plan(_request(source, staging))
    assert plan.limitations == _EXPECTED_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS
    task_id = service.accept(plan.plan_id, "task.test")
    outcome = service.execute_accepted(task_id)

    assert outcome.state == "completed"
    assert outcome.bundle is not None
    assert outcome.bundle.task_id == task_id
    assert outcome.bundle.artifacts[0].kind == "document"
    assert outcome.bundle.artifacts[1].kind == "resource"
    assert outcome.bundle.entries[0].role == "primary"
    assert outcome.bundle.relations[0].role == "manifest"
    assert controller.released == [task_id]
    assert service.cancel(task_id) == "already_terminal"


@pytest.mark.parametrize(
    ("retained_index", "expected_code"),
    [
        (1, "docwen.resolved_document.missing"),
        (0, "docwen.numbering_export_plan.missing"),
    ],
)
def test_v4_dual_input_missing_role_fails_without_artifact(
    tmp_path: Path, retained_index: int, expected_code: str
) -> None:
    source = tmp_path / "source.md"
    source.write_text("# Source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    request = _request(source, staging)
    service = ConversionService(_Controller(), _Committer())

    with pytest.raises(ConversionServiceError) as rejected:
        service.plan(replace(request, inputs=(request.inputs[retained_index],)))

    assert rejected.value.code == expected_code
    assert list(staging.iterdir()) == []


def test_v4_invalid_plan_fails_without_artifact(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# Source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    request = _request(source, staging)
    Path(request.inputs[1].path).write_text("{}", encoding="utf-8")
    request = replace(request, inputs=(request.inputs[0], _refingerprint(request.inputs[1])))

    with pytest.raises(ConversionServiceError) as rejected:
        ConversionService(_Controller(), _Committer()).plan(request)

    assert rejected.value.code == "docwen.numbering_export_plan.invalid"
    assert list(staging.iterdir()) == []


def test_v4_unsupported_materialization_is_distinct_and_has_no_artifact(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# Source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    request = _request(source, staging)
    neutral_path = Path(request.inputs[0].path)
    plan_path = Path(request.inputs[1].path)
    document = json.loads(neutral_path.read_text(encoding="utf-8"))
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    line = "# Source"
    target = {
        "source_start": 0,
        "source_end": len(line),
        "source_slice_sha256": hashlib.sha256(line.encode()).hexdigest(),
        "kind": "heading",
        "target_id": "source",
        "heading_level": 1,
        "authored_text": "Source",
    }
    document["document"]["targets"] = [target]
    plan["plan"] = {
        "heading_definitions": [
            {
                "definition_id": "definition-1",
                "levels": [
                    {
                        "level": 1,
                        "start": 100,
                        "number_format": "chinese_lower",
                        "display": [{"counter": {"level": 1, "number_format": "chinese_lower"}}],
                        "suffix": "space",
                        "restart_after_level": None,
                    }
                ],
            }
        ],
        "heading_instances": [{"instance_id": "instance-1", "definition_id": "definition-1", "starts": []}],
        "targets": [
            {
                "source_start": 0,
                "source_end": len(line),
                "kind": "heading",
                "enabled": True,
                "target_id": "source",
                "derived_number": "一百",
                "materialization": {
                    "type": "heading_list",
                    "definition_id": "definition-1",
                    "instance_id": "instance-1",
                    "level": 1,
                },
            }
        ],
    }
    plan_sha = hashlib.sha256(canonicalize_numbering_plan(plan["plan"])).hexdigest()
    plan["plan_sha256"] = plan_sha
    document["plan_sha256"] = plan_sha
    neutral_path.write_text(json.dumps(document, separators=(",", ":")), encoding="utf-8")
    plan_path.write_text(json.dumps(plan, separators=(",", ":")), encoding="utf-8")
    request = replace(request, inputs=tuple(_refingerprint(item) for item in request.inputs))

    with pytest.raises(ConversionServiceError) as rejected:
        ConversionService(_Controller(), _Committer()).plan(request)

    assert rejected.value.category == "unsupported"
    assert rejected.value.code == "docwen.numbering_export_plan.unsupported_materialization"
    assert list(staging.iterdir()) == []


def test_v4_markdown_plan_rejects_third_linked_resource_input(tmp_path: Path) -> None:
    source = tmp_path / "physical-source" / "report.md"
    resource = tmp_path / "physical-resource" / "provided.png"
    source.parent.mkdir()
    resource.parent.mkdir()
    source.write_text("![chart](assets/chart.png)\n", encoding="utf-8")
    resource.write_bytes(b"png fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())

    request = _request(source, staging)
    request = ConversionPlanRequest(
        capability_id=request.capability_id,
        inputs=(
            *request.inputs,
            LocalInputHandle(
                input_id="input.legacy-resource",
                path=str(resource),
                kind="resource",
                role="linked_resource",
                logical_path="document/assets/chart.png",
                media_type=PNG_MEDIA_TYPE,
                size_bytes=resource.stat().st_size,
                sha256=_sha256(resource),
            ),
        ),
        output=request.output,
        options=request.options,
    )

    with pytest.raises(ConversionServiceError) as rejected:
        service.plan(request)
    assert rejected.value.code == "undeclared_input_role"


def test_v4_markdown_plan_rejects_third_bibliography_input(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    bibliography = tmp_path / "bibliography.json"
    source.write_text("# source\n", encoding="utf-8")
    bibliography.write_text('{"schema":"docwen.semantic_bibliography.v1","entries":[]}', encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _Controller()
    service = ConversionService(controller, _Committer())
    request = _request(source, staging)
    typed = LocalInputHandle(
        input_id="input.2",
        path=str(bibliography),
        kind="resource",
        role="bibliography",
        logical_path="bibliography.json",
        media_type=SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE,
        size_bytes=bibliography.stat().st_size,
        sha256=_sha256(bibliography),
    )
    with pytest.raises(ConversionServiceError) as rejected:
        service.plan(replace(request, inputs=(*request.inputs, typed)))
    assert rejected.value.code == "undeclared_input_role"


def test_v4_markdown_plan_rejects_legacy_bibliography_before_cardinality(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    first = tmp_path / "first.json"
    second = tmp_path / "second.json"
    source.write_text("# source\n", encoding="utf-8")
    first.write_bytes(b"{}")
    second.write_bytes(b"{}")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())
    request = _request(source, staging)
    typed = tuple(
        LocalInputHandle(
            input_id=f"input.{index}",
            path=str(path),
            kind="resource",
            role="bibliography",
            logical_path=path.name,
            media_type=SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE,
            size_bytes=path.stat().st_size,
            sha256=_sha256(path),
        )
        for index, path in enumerate((first, second), start=2)
    )

    with pytest.raises(ConversionServiceError) as rejected:
        service.plan(replace(request, inputs=(*request.inputs, *typed)))

    assert rejected.value.code == "undeclared_input_role"


@pytest.mark.parametrize(
    ("replacement", "expected_code"),
    [
        ({"role": "unknown_role", "logical_path": "refs.json"}, "undeclared_input_role"),
        ({"kind": "document", "logical_path": "refs.json"}, "input_slot_kind_mismatch"),
        (
            {
                "kind": "resource",
                "media_type": "image/jpeg",
                "logical_path": "plan.jpg",
            },
            "input_slot_media_type_mismatch",
        ),
        ({"logical_path": "plans/../plan.json"}, "invalid_input_logical_path"),
    ],
)
def test_markdown_plan_rejects_invalid_typed_input_shapes(
    tmp_path: Path,
    replacement: dict[str, str],
    expected_code: str,
) -> None:
    source = tmp_path / "source.md"
    source.write_text("# source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())
    valid = _request(source, staging)
    invalid = replace(valid.inputs[1], **replacement)
    request = ConversionPlanRequest(
        capability_id=MARKDOWN_TO_DOCX_CAPABILITY_ID,
        inputs=(valid.inputs[0], invalid),
        output=StagingOutputTarget(staging_root=str(staging)),
        options={},
    )

    with pytest.raises(ConversionServiceError) as rejected:
        service.plan(request)
    assert rejected.value.code == expected_code


def test_plan_validates_all_roles_before_any_kind_or_media(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())
    valid = _request(source, staging)
    wrong_media = replace(valid.inputs[1], media_type="image/jpeg")
    undeclared = replace(
        wrong_media,
        input_id="input.undeclared",
        role="citation_style",
        logical_path="citation.csl",
    )

    with pytest.raises(ConversionServiceError) as rejected:
        service.plan(
            ConversionPlanRequest(
                capability_id=MARKDOWN_TO_DOCX_CAPABILITY_ID,
                inputs=(valid.inputs[0], wrong_media, undeclared),
                output=StagingOutputTarget(staging_root=str(staging)),
            )
        )

    assert rejected.value.code == "undeclared_input_role"


def test_plan_validates_all_kinds_before_any_media(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())
    valid = _request(source, staging)
    wrong_media = replace(valid.inputs[1], media_type="image/jpeg")
    wrong_kind = replace(valid.inputs[0], kind="resource")

    with pytest.raises(ConversionServiceError) as rejected:
        service.plan(
            ConversionPlanRequest(
                capability_id=MARKDOWN_TO_DOCX_CAPABILITY_ID,
                inputs=(wrong_kind, wrong_media),
                output=StagingOutputTarget(staging_root=str(staging)),
            )
        )

    assert rejected.value.code == "input_slot_kind_mismatch"


def test_markdown_plan_rejects_duplicate_logical_path_and_neutral_cardinality(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())
    valid = _request(source, staging)
    duplicate_path = replace(valid.inputs[1], logical_path=valid.inputs[0].logical_path)
    with pytest.raises(ConversionServiceError) as duplicate:
        service.plan(
            ConversionPlanRequest(
                capability_id=MARKDOWN_TO_DOCX_CAPABILITY_ID,
                inputs=(valid.inputs[0], duplicate_path),
                output=StagingOutputTarget(staging_root=str(staging)),
                options={},
            )
        )
    assert duplicate.value.code == "duplicate_input_logical_path"

    second = replace(valid.inputs[0], input_id="input.neutral-2", logical_path="second.json")
    with pytest.raises(ConversionServiceError) as cardinality:
        service.plan(
            ConversionPlanRequest(
                capability_id=MARKDOWN_TO_DOCX_CAPABILITY_ID,
                inputs=(valid.inputs[0], second, valid.inputs[1]),
                output=StagingOutputTarget(staging_root=str(staging)),
                options={},
            )
        )
    assert cardinality.value.code == "docwen.resolved_document.invalid"


def test_consumer_facing_options_are_discovered_resolved_and_validated(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# Source\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())

    capabilities = {item.capability_id: item for item in service.list_capabilities()}
    schema = capabilities[MARKDOWN_TO_DOCX_CAPABILITY_ID].options_schema
    assert schema["additionalProperties"] is False
    assert set(schema["properties"]) == {
        "locale",
        "template_name",
        "heading_merge_mode",
    }
    plan = service.plan(
        _request(
            source,
            staging,
            options={
                "locale": "en_US",
                "template_name": f"template.docx.{'a' * 64}",
                "heading_merge_mode": "never",
            },
        )
    )
    assert plan.effective_options["template_name"] == f"template.docx.{'a' * 64}"
    assert plan.effective_options["locale"] == "en_US"
    assert plan.effective_options["heading_merge_mode"] == "never"

    other_staging = tmp_path / "other-staging"
    other_staging.mkdir()
    with pytest.raises(ConversionServiceError, match="does not accept") as exc_info:
        service.plan(_request(source, other_staging, options={"add_numbering": True}))
    assert exc_info.value.code == "unsupported_options"

    invalid_locale_staging = tmp_path / "invalid-locale-staging"
    invalid_locale_staging.mkdir()
    with pytest.raises(ConversionServiceError) as invalid_locale:
        service.plan(_request(source, invalid_locale_staging, options={"locale": "en_GB"}))
    assert invalid_locale.value.code == "option_value_invalid"


@pytest.mark.parametrize(
    ("capability_id", "media_type", "suffix", "expected_kind", "expected_role"),
    [
        (MARKDOWN_TO_XLSX_CAPABILITY_ID, MARKDOWN_MEDIA_TYPE, ".md", "document", "primary"),
        (XLSX_TO_MARKDOWN_CAPABILITY_ID, XLSX_MEDIA_TYPE, ".xlsx", "document", "primary"),
        (MARKDOWN_VALIDATE_CAPABILITY_ID, MARKDOWN_MEDIA_TYPE, ".md", "resource", "supplementary"),
        (MARKDOWN_NUMBERING_CAPABILITY_ID, MARKDOWN_MEDIA_TYPE, ".md", "document", "primary"),
    ],
)
def test_new_single_input_capabilities_publish_declared_bundle_semantics(
    tmp_path: Path,
    capability_id: str,
    media_type: str,
    suffix: str,
    expected_kind: str,
    expected_role: str,
) -> None:
    source = tmp_path / f"source{suffix}"
    source.write_bytes(b"fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())

    plan = service.plan(_request(source, staging, capability_id=capability_id, media_type=media_type))
    outcome = service.execute_accepted(service.accept(plan.plan_id, f"task.{capability_id}"))

    assert outcome.bundle is not None
    assert outcome.bundle.artifacts[0].kind == expected_kind
    assert outcome.bundle.entries[0].role == expected_role


@pytest.mark.parametrize(
    ("capability_id", "media_types", "expected_kind", "expected_role"),
    [
        (PDF_MERGE_CAPABILITY_ID, [PDF_MEDIA_TYPE, PDF_MEDIA_TYPE], "document", "primary"),
        (XLSX_MERGE_TABLES_CAPABILITY_ID, [XLSX_MEDIA_TYPE, XLSX_MEDIA_TYPE], "document", "primary"),
        (IMAGES_MERGE_TO_TIFF_CAPABILITY_ID, [PNG_MEDIA_TYPE, "image/jpeg"], "resource", "image"),
    ],
)
def test_aggregate_capabilities_preserve_input_order_and_publish_one_bundle(
    tmp_path: Path,
    capability_id: str,
    media_types: list[str],
    expected_kind: str,
    expected_role: str,
) -> None:
    inputs = [tmp_path / f"input-{index}.bin" for index in range(2)]
    for index, input_path in enumerate(inputs):
        input_path.write_bytes(f"input {index}".encode())
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())

    request = _request_many(
        inputs,
        staging,
        capability_id=capability_id,
        media_types=media_types,
    )
    plan = service.plan(request)
    outcome = service.execute_accepted(service.accept(plan.plan_id, f"task.{capability_id}"))

    assert outcome.bundle is not None
    assert outcome.bundle.artifacts[0].kind == expected_kind
    assert outcome.bundle.entries[0].role == expected_role


def test_aggregate_capability_rejects_too_few_inputs(tmp_path: Path) -> None:
    source = tmp_path / "source.pdf"
    source.write_bytes(b"pdf")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())

    with pytest.raises(ConversionServiceError, match="invalid cardinality") as exc_info:
        service.plan(
            _request(
                source,
                staging,
                capability_id=PDF_MERGE_CAPABILITY_ID,
                media_type=PDF_MEDIA_TYPE,
            )
        )
    assert exc_info.value.code == "input_slot_cardinality_mismatch"
