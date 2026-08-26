"""Focused tests split from test_resolved_numbering_port.py."""

from __future__ import annotations

from ._resolved_numbering_port_support import (
    _ONE_PIXEL_PNG,
    CaptionMaterialization,
    Draft202012Validator,
    HeadingListMaterialization,
    Path,
    ResolvedNumberingPortError,
    _basic_heading_resources,
    _caption_matrix_resources,
    _cross_reference_resources,
    _digest,
    _disabled_heading_scope_resources,
    _embedded_resource,
    _plan_target,
    _resource_occurrence,
    _resources,
    _target,
    base64,
    canonicalize_numbering_plan,
    hashlib,
    json,
    load_resolved_numbering_bytes,
    pytest,
    replace,
    validate_document,
    validate_plan,
    validate_port,
)

pytestmark = pytest.mark.contract


def test_valid_port_authenticates_authored_manual_prefix_and_numeric_source_id() -> None:
    document, plan = _basic_heading_resources()
    port = load_resolved_numbering_bytes(document, plan)
    assert port.document.authored_markdown.startswith("# 2.3 标题")
    assert port.document.targets[0].authored_text == "2.3 标题 ^1-target"
    assert port.document.targets[0].target_id == "1-target"
    assert port.plan.targets[0].derived_number == "1"


def test_cross_reference_token_binds_stable_selector_alias_and_opaque_page_locator() -> None:
    document, plan = _cross_reference_resources(
        "@[[Folder/Other Page#^1-target|Current title]]",
        alias="Current title",
    )

    port = load_resolved_numbering_bytes(document, plan)

    assert port.document.references[0].target_id == "1-target"
    assert port.document.references[0].alias == "Current title"


@pytest.mark.parametrize(
    ("token", "alias"),
    [
        ("@[[Other#^wrong-id|Current title]]", "Current title"),
        ("@[[Other#^1-target|Authored alias]]", "Forged alias"),
    ],
)
def test_cross_reference_stable_selector_and_alias_mutations_fail_closed(
    token: str,
    alias: str,
) -> None:
    document, plan = _cross_reference_resources(token, alias=alias)

    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(document, plan)
    assert error.value.code == "docwen.resolved_document.invalid"


def test_soft_heading_selector_cannot_be_retyped_as_a_caption_reference() -> None:
    document, plan = _cross_reference_resources(
        "@[[Other#Caption]]",
        alias=None,
        target_kind="figure",
        target_id=None,
    )

    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(document, plan)
    assert error.value.code == "docwen.resolved_document.invalid"


def test_disabled_target_is_valid_and_has_no_numbering_objects() -> None:
    source = "Figure: authored caption\n"
    figure = _target(source, "Figure: authored caption", "figure", None)
    plan = {
        "heading_definitions": [],
        "heading_instances": [],
        "targets": [_plan_target(figure, enabled=False, derived_number=None, materialization=None)],
    }
    document_bytes, plan_bytes = _resources(source, [figure], plan)
    port = load_resolved_numbering_bytes(document_bytes, plan_bytes)
    assert port.plan.targets[0].enabled is False
    assert port.plan.targets[0].materialization is None


@pytest.mark.parametrize("materialization_type", ["simple_seq", "chapter_seq"])
@pytest.mark.parametrize("action", ["continue", "reset_to_start", "restart_by_heading_level"])
def test_caption_type_and_action_are_orthogonal(materialization_type: str, action: str) -> None:
    document, plan = _caption_matrix_resources(
        materialization_type,
        action,
        sequence_start=7 if action == "reset_to_start" else 1,
    )
    port = load_resolved_numbering_bytes(document, plan)
    materialization = port.plan.targets[-1].materialization
    assert isinstance(materialization, CaptionMaterialization)
    assert materialization.type == materialization_type
    assert materialization.sequence_action == action


def test_chapter_and_restart_levels_may_differ_without_cache_splitting() -> None:
    document, plan = _caption_matrix_resources(
        "chapter_seq", "restart_by_heading_level", chapter_cache="1", chapter_level=1
    )
    port = load_resolved_numbering_bytes(document, plan)
    materialization = port.plan.targets[-1].materialization
    assert isinstance(materialization, CaptionMaterialization)
    assert materialization.chapter_heading_level == 1
    assert materialization.restart_heading_level == 2
    assert materialization.chapter_cached_number == "1"
    assert port.plan.targets[-1].derived_number == "1-1"


@pytest.mark.parametrize("level", [7, 8, 9])
def test_typed_document_heading_levels_through_nine_are_valid_but_must_match_the_plan(level: int) -> None:
    document_bytes, plan_bytes = _basic_heading_resources()
    port = load_resolved_numbering_bytes(document_bytes, plan_bytes)
    target = replace(port.document.targets[0], heading_level=level)
    document = replace(port.document, targets=(target,))
    document_envelope = replace(port.document_envelope, document=document)

    validate_document(document, port.source_sha256, "docwen.resolved_document.invalid")
    with pytest.raises(ResolvedNumberingPortError):
        validate_port(document_envelope, port.plan_envelope)


@pytest.mark.parametrize("level", [7, 8, 9])
def test_typed_heading_materialization_must_match_the_document_target(level: int) -> None:
    document_bytes, plan_bytes = _basic_heading_resources()
    port = load_resolved_numbering_bytes(document_bytes, plan_bytes)
    target = port.plan.targets[0]
    assert isinstance(target.materialization, HeadingListMaterialization)
    materialization = replace(target.materialization, level=level)
    plan = replace(port.plan, targets=(replace(target, materialization=materialization),))

    with pytest.raises(ResolvedNumberingPortError):
        validate_plan(plan)


@pytest.mark.parametrize("level", [7, 8, 9])
@pytest.mark.parametrize("binding", ["chapter", "restart"])
def test_typed_caption_heading_bindings_through_nine_pass_plan_validation(
    level: int,
    binding: str,
) -> None:
    if binding == "chapter":
        document_bytes, plan_bytes = _caption_matrix_resources("chapter_seq", "continue")
    else:
        document_bytes, plan_bytes = _caption_matrix_resources("simple_seq", "restart_by_heading_level")
    port = load_resolved_numbering_bytes(document_bytes, plan_bytes)
    target = port.plan.targets[-1]
    assert isinstance(target.materialization, CaptionMaterialization)
    if binding == "chapter":
        materialization = replace(
            target.materialization,
            chapter_heading_level=level,
            chapter_heading_style=f"heading_{level}",
        )
    else:
        materialization = replace(
            target.materialization,
            restart_heading_level=level,
            restart_heading_style=f"heading_{level}",
        )
    plan = replace(
        port.plan,
        targets=(*port.plan.targets[:-1], replace(target, materialization=materialization)),
    )

    validate_plan(plan)
    # This fixture contains only Heading 1/2, so the full port still rejects
    # the binding for the independent reason that no matching Heading 7-9
    # precedes the caption.
    with pytest.raises(ResolvedNumberingPortError, match="no preceding Heading"):
        validate_port(port.document_envelope, replace(port.plan_envelope, plan=plan))


def test_typed_heading_level_ten_fails_the_document_boundary() -> None:
    document_bytes, plan_bytes = _basic_heading_resources()
    port = load_resolved_numbering_bytes(document_bytes, plan_bytes)
    target = replace(port.document.targets[0], heading_level=10)
    document = replace(port.document, targets=(target,))

    with pytest.raises(ResolvedNumberingPortError):
        validate_document(document, port.source_sha256, "docwen.resolved_document.invalid")


@pytest.mark.parametrize("binding", ["chapter", "restart"])
def test_typed_caption_heading_binding_level_ten_fails(binding: str) -> None:
    if binding == "chapter":
        document_bytes, plan_bytes = _caption_matrix_resources("chapter_seq", "continue")
    else:
        document_bytes, plan_bytes = _caption_matrix_resources("simple_seq", "restart_by_heading_level")
    port = load_resolved_numbering_bytes(document_bytes, plan_bytes)
    target = port.plan.targets[-1]
    assert isinstance(target.materialization, CaptionMaterialization)
    if binding == "chapter":
        materialization = replace(
            target.materialization,
            chapter_heading_level=10,
            chapter_heading_style="heading_10",
        )
    else:
        materialization = replace(
            target.materialization,
            restart_heading_level=10,
            restart_heading_style="heading_10",
        )
    plan = replace(
        port.plan,
        targets=(*port.plan.targets[:-1], replace(target, materialization=materialization)),
    )

    with pytest.raises(ResolvedNumberingPortError):
        validate_plan(plan)


def test_restart_scope_uses_disabled_heading_style_occurrence() -> None:
    document, plan = _disabled_heading_scope_resources(chapter_seq=False)

    port = load_resolved_numbering_bytes(document, plan)

    second_caption = port.plan.targets[-1].materialization
    assert isinstance(second_caption, CaptionMaterialization)
    assert second_caption.sequence_cached_number == "1"


def test_chapter_seq_does_not_skip_disabled_latest_heading_for_earlier_cache() -> None:
    document, plan = _disabled_heading_scope_resources(chapter_seq=True)

    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(document, plan)

    assert error.value.code == "docwen.numbering_export_plan.invalid"


@pytest.mark.parametrize(
    ("number_format", "accepted_value", "accepted_text", "rejected_value", "rejected_text"),
    [
        ("chinese_lower", 99, "九十九", 100, "一百"),
        ("roman_upper", 3999, "MMMCMXCIX", 4000, "MMMM"),
        ("arabic_circled", 50, "㊿", 51, "(51)"),
    ],
)
def test_unproven_host_counter_boundaries_fail_unsupported(
    number_format: str,
    accepted_value: int,
    accepted_text: str,
    rejected_value: int,
    rejected_text: str,
) -> None:
    load_resolved_numbering_bytes(
        *_basic_heading_resources(number_format=number_format, start=accepted_value, derived=accepted_text)
    )
    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(
            *_basic_heading_resources(number_format=number_format, start=rejected_value, derived=rejected_text)
        )
    assert error.value.code == "docwen.numbering_export_plan.unsupported_materialization"


@pytest.mark.parametrize("target_id", ["_bad", "bad/id", "a" * 129])
def test_source_target_id_grammar_rejects_legacy_or_oversize_forms(target_id: str) -> None:
    document_bytes, plan_bytes = _basic_heading_resources()
    document = json.loads(document_bytes)
    plan = json.loads(plan_bytes)
    document["document"]["targets"][0]["target_id"] = target_id
    plan["plan"]["targets"][0]["target_id"] = target_id
    plan["plan_sha256"] = hashlib.sha256(canonicalize_numbering_plan(plan["plan"])).hexdigest()
    document["plan_sha256"] = plan["plan_sha256"]
    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(
            json.dumps(document, ensure_ascii=False, separators=(",", ":")).encode(),
            json.dumps(plan, ensure_ascii=False, separators=(",", ":")).encode(),
        )
    assert error.value.code == "docwen.resolved_document.invalid"


@pytest.mark.parametrize("mutation", ["source_sha", "slice_sha", "plan_sha", "pointer", "duplicate_key"])
def test_authenticated_dual_input_mutations_fail_closed(mutation: str) -> None:
    document_bytes, plan_bytes = _basic_heading_resources()
    document = json.loads(document_bytes)
    plan = json.loads(plan_bytes)
    if mutation == "source_sha":
        document["source_sha256"] = "0" * 64
    elif mutation == "slice_sha":
        document["document"]["targets"][0]["source_slice_sha256"] = "0" * 64
    elif mutation == "plan_sha":
        plan["plan_sha256"] = "0" * 64
    elif mutation == "pointer":
        document["input_id"] = "other"
    else:
        duplicate = document_bytes.replace(
            b'"schema":"docwen.resolved_document.v1",',
            b'"schema":"docwen.resolved_document.v1","schema":"duplicate",',
            1,
        )
        with pytest.raises(ResolvedNumberingPortError):
            load_resolved_numbering_bytes(duplicate, plan_bytes)
        return
    with pytest.raises(ResolvedNumberingPortError):
        load_resolved_numbering_bytes(
            json.dumps(document, ensure_ascii=False, separators=(",", ":")).encode(),
            json.dumps(plan, ensure_ascii=False, separators=(",", ":")).encode(),
        )


def test_closed_schemas_accept_exact_six_combo_resources() -> None:
    root = Path(__file__).parents[3]
    document_schema = json.loads(
        (root / "contracts/schemas/docwen.resolved_document.v1.schema.json").read_text(encoding="utf-8")
    )
    plan_schema = json.loads(
        (root / "contracts/schemas/docwen.numbering_export_plan.v1.schema.json").read_text(encoding="utf-8")
    )
    Draft202012Validator.check_schema(document_schema)
    Draft202012Validator.check_schema(plan_schema)
    document_validator = Draft202012Validator(document_schema)
    plan_validator = Draft202012Validator(plan_schema)
    for materialization_type in ("simple_seq", "chapter_seq"):
        for action in ("continue", "reset_to_start", "restart_by_heading_level"):
            document, plan = _caption_matrix_resources(materialization_type, action)
            document_validator.validate(json.loads(document))
            plan_validator.validate(json.loads(plan))


def test_cache_component_mutation_cannot_be_rescued_by_ambiguous_total() -> None:
    document_bytes, plan_bytes = _caption_matrix_resources(
        "chapter_seq", "reset_to_start", chapter_cache="1-1", chapter_level=2, sequence_start=3
    )
    document = json.loads(document_bytes)
    plan = json.loads(plan_bytes)
    materialization = plan["plan"]["targets"][-1]["materialization"]
    materialization["chapter_cached_number"] = "1"
    materialization["sequence_cached_number"] = "1-3"
    plan_sha = hashlib.sha256(canonicalize_numbering_plan(plan["plan"])).hexdigest()
    plan["plan_sha256"] = plan_sha
    document["plan_sha256"] = plan_sha
    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(
            json.dumps(document, ensure_ascii=False, separators=(",", ":")).encode(),
            json.dumps(plan, ensure_ascii=False, separators=(",", ":")).encode(),
        )
    assert error.value.code == "docwen.numbering_export_plan.invalid"


def test_unknown_or_control_materialization_text_fails_without_ooxml_mutation() -> None:
    document_bytes, plan_bytes = _caption_matrix_resources("chapter_seq", "continue")
    document = json.loads(document_bytes)
    plan = json.loads(plan_bytes)
    plan["plan"]["targets"][-1]["materialization"]["chapter_separator"] = "\u0001"
    plan_sha = hashlib.sha256(canonicalize_numbering_plan(plan["plan"])).hexdigest()
    plan["plan_sha256"] = plan_sha
    document["plan_sha256"] = plan_sha
    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(
            json.dumps(document, ensure_ascii=False, separators=(",", ":")).encode(),
            json.dumps(plan, ensure_ascii=False, separators=(",", ":")).encode(),
        )
    assert error.value.code == "docwen.numbering_export_plan.unsupported_materialization"


def test_range_bound_resources_allow_same_locator_to_resolve_differently() -> None:
    token = "![[same.png]]"
    source = f"{token}\n{token}\n"
    plan_value: dict[str, object] = {
        "heading_definitions": [],
        "heading_instances": [],
        "targets": [],
    }
    document_bytes, plan_bytes = _resources(source, [], plan_value)
    document = json.loads(document_bytes)
    first = _resource_occurrence(source, token, "resource-a")
    second = _resource_occurrence(source, token, "resource-b", start=first["source_end"])
    document["document"]["resource_occurrences"] = [first, second]
    document["document"]["resources"] = [
        _embedded_resource("resource-a"),
        _embedded_resource("resource-b"),
    ]

    port = load_resolved_numbering_bytes(json.dumps(document, separators=(",", ":")).encode(), plan_bytes)

    assert [item.resource_id for item in port.document.resource_occurrences] == [
        "resource-a",
        "resource-b",
    ]
    assert port.document.resources[0].content == _ONE_PIXEL_PNG


@pytest.mark.parametrize(
    "mutation",
    ["token_range", "locator", "missing", "unused", "wrong_magic", "empty"],
)
def test_embedded_resource_mutations_fail_before_materialization(mutation: str) -> None:
    token = "![[same.png]]"
    source = f"{token}\n"
    plan_value: dict[str, object] = {
        "heading_definitions": [],
        "heading_instances": [],
        "targets": [],
    }
    document_bytes, plan_bytes = _resources(source, [], plan_value)
    document = json.loads(document_bytes)
    occurrence = _resource_occurrence(source, token, "resource-a")
    resource = _embedded_resource("resource-a")
    document["document"]["resource_occurrences"] = [occurrence]
    document["document"]["resources"] = [resource]
    if mutation == "token_range":
        occurrence["source_end"] -= 1
    elif mutation == "locator":
        occurrence["authored_locator"] = "other.png"
    elif mutation == "missing":
        occurrence["resource_id"] = "resource-missing"
    elif mutation == "unused":
        document["document"]["resources"].append(_embedded_resource("resource-b"))
    elif mutation == "wrong_magic":
        resource.update(_embedded_resource("resource-a", b"not a png"))
    else:
        resource.update(_embedded_resource("resource-a", b""))

    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(json.dumps(document, separators=(",", ":")).encode(), plan_bytes)
    assert error.value.code == "docwen.resolved_document.invalid"


def test_closed_bibliography_and_both_citation_forms_are_admitted() -> None:
    narrative = "@fig-legacy"
    parenthetical = "[@smith; @wang]"
    source = f"See {narrative} and {parenthetical}.\n"
    plan_value: dict[str, object] = {
        "heading_definitions": [],
        "heading_instances": [],
        "targets": [],
    }
    document_bytes, plan_bytes = _resources(source, [], plan_value)
    document = json.loads(document_bytes)
    bibliography = json.dumps(
        {
            "schema": "docwen.semantic_bibliography.v1",
            "entries": [
                {"item_id": "record-smith", "runs": [{"text": "Smith (2026)."}]},
                {"item_id": "record-wang", "runs": [{"text": "Wang (2025)."}]},
            ],
        },
        separators=(",", ":"),
    ).encode()
    document["document"]["resources"] = [
        {
            "resource_id": "bibliography-1",
            "role": "bibliography",
            "media_type": "application/vnd.docwen.semantic-bibliography+json",
            "size_bytes": len(bibliography),
            "sha256": hashlib.sha256(bibliography).hexdigest(),
            "content_base64": base64.b64encode(bibliography).decode("ascii"),
        }
    ]

    def citation(token: str, form: str, cluster_id: str, keys: list[str]):
        start = source.index(token)
        return {
            "source_start": start,
            "source_end": start + len(token),
            "source_slice_sha256": _digest(token),
            "authored_token": token,
            "form": form,
            "cluster_id": cluster_id,
            "items": [
                {
                    "citation_key": key,
                    "record_id": f"record:{key}",
                    "record_sha256": _digest(f"record:{key}"),
                    "presentation": key.title(),
                }
                for key in keys
            ],
            "cached_result": "; ".join(key.title() for key in keys),
        }

    document["document"]["citations"] = [
        citation(narrative, "narrative", "cluster-a", ["fig-legacy"]),
        citation(parenthetical, "parenthetical", "cluster-b", ["smith", "wang"]),
    ]
    document["document"]["citations"][0]["items"][0]["record_id"] = "reference-record:98"
    document["document"]["citations"][1]["items"][0]["record_id"] = "98.numeric-leading-record"
    document["document"]["citations"][1]["items"][1]["record_id"] = "record:" + ("long-" * 40) + "end"

    port = load_resolved_numbering_bytes(json.dumps(document, separators=(",", ":")).encode(), plan_bytes)
    assert port.document.citations[0].items[0].citation_key == "fig-legacy"
    assert port.document.citations[0].items[0].record_id == "reference-record:98"
    assert port.document.citations[1].items[0].record_id.startswith("98.")
    assert len(port.document.citations[1].items[1].record_id) > 200
    assert port.document.resources[0].role == "bibliography"


@pytest.mark.parametrize("mutation", ["citation_order", "duplicate_key", "bad_bibliography"])
def test_resolved_citation_and_bibliography_mutations_fail_closed(mutation: str) -> None:
    source = "[@smith; @wang]\n"
    plan_value: dict[str, object] = {
        "heading_definitions": [],
        "heading_instances": [],
        "targets": [],
    }
    document_bytes, plan_bytes = _resources(source, [], plan_value)
    document = json.loads(document_bytes)
    token = source.rstrip()
    start = source.index(token)
    items = [
        {
            "citation_key": key,
            "record_id": f"record:{key}",
            "record_sha256": _digest(key),
            "presentation": key,
        }
        for key in ("smith", "wang")
    ]
    document["document"]["citations"] = [
        {
            "source_start": start,
            "source_end": start + len(token),
            "source_slice_sha256": _digest(token),
            "authored_token": token,
            "form": "parenthetical",
            "cluster_id": "cluster-a",
            "items": items,
            "cached_result": "Smith; Wang",
        }
    ]
    if mutation == "citation_order":
        items.reverse()
    elif mutation == "duplicate_key":
        items[1] = dict(items[0])
    else:
        bibliography = b'{"schema":"docwen.semantic_bibliography.v1","entries":{},"extra":true}'
        document["document"]["resources"] = [
            {
                "resource_id": "bibliography-1",
                "role": "bibliography",
                "media_type": "application/vnd.docwen.semantic-bibliography+json",
                "size_bytes": len(bibliography),
                "sha256": hashlib.sha256(bibliography).hexdigest(),
                "content_base64": base64.b64encode(bibliography).decode("ascii"),
            }
        ]

    with pytest.raises(ResolvedNumberingPortError) as error:
        load_resolved_numbering_bytes(json.dumps(document, separators=(",", ":")).encode(), plan_bytes)
    assert error.value.code == "docwen.resolved_document.invalid"
