from __future__ import annotations

import copy
import io
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

import pytest
from scripts.release import build_v4_package_input as producer
from scripts.release import v4_package_input_ooxml_identity as identity
from tests.test_repo import test_build_v4_package_input as base
from tests.test_repo import v4_package_input_test_support as support

pytestmark = pytest.mark.contract

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _xml_mutation(part: str, operation: Callable[[ElementTree.Element], None]) -> Callable[[dict[str, bytes]], None]:
    def mutate(parts: dict[str, bytes]) -> None:
        root = ElementTree.fromstring(parts[part])
        operation(root)
        parts[part] = ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    return mutate


def _remove_first(root: ElementTree.Element, tag: str) -> None:
    parents = {child: parent for parent in root.iter() for child in parent}
    item = next(root.iter(tag))
    parents[item].remove(item)


def _drop_member(name: str) -> Callable[[dict[str, bytes]], None]:
    return lambda parts: parts.pop(name)


def _wrong_num_id(root: ElementTree.Element) -> None:
    next(root.iter(f"{_W}numId")).set(f"{_W}val", "999")


def _duplicate_bookmark_name(root: ElementTree.Element) -> None:
    starts = list(root.iter(f"{_W}bookmarkStart"))
    starts[1].set(f"{_W}name", str(starts[0].get(f"{_W}name")))


def _wrong_field_cache(root: ElementTree.Element) -> None:
    for item in root.iter(f"{_W}t"):
        if item.text == "1":
            item.text = "999"
            return
    raise AssertionError("cached field result missing")


def _wrong_seq_instruction(root: ElementTree.Element) -> None:
    for item in root.iter(f"{_W}instrText"):
        if (item.text or "").startswith(" SEQ "):
            item.text = f"{item.text} \\r 999"
            return
    raise AssertionError("SEQ instruction missing")


def _drop_styles_relationship(root: ElementTree.Element) -> None:
    for item in root:
        if (item.get("Type") or "").endswith("/styles"):
            root.remove(item)
            return
    raise AssertionError("styles relationship missing")


def _move_heading_after_caption_targets(root: ElementTree.Element) -> None:
    body = root.find(f"{_W}body")
    assert body is not None
    target_carriers = [
        child
        for child in body
        if child.tag == f"{_W}sdt"
        and (tag := child.find(f"{_W}sdtPr/{_W}tag")) is not None
        and (tag.get(f"{_W}val") or "").startswith("docwen-target-v1:")
    ]
    heading = next(
        carrier
        for carrier in target_carriers
        if (style := carrier.find(f".//{_W}pPr/{_W}pStyle")) is not None and style.get(f"{_W}val") == "Heading1"
    )
    body.remove(heading)
    body.insert(list(body).index(target_carriers[-1]) + 1, heading)


def _add_managed_style_alias(root: ElementTree.Element) -> None:
    heading = next(item for item in root.findall(f"{_W}style") if item.get(f"{_W}styleId") == "Heading1")
    ElementTree.SubElement(heading, f"{_W}aliases", {f"{_W}val": "Spoof Heading"})


def _add_unmanaged_style_alias(root: ElementTree.Element) -> None:
    style = ElementTree.SubElement(
        root,
        f"{_W}style",
        {f"{_W}type": "paragraph", f"{_W}styleId": "TemplateStyle"},
    )
    ElementTree.SubElement(style, f"{_W}name", {f"{_W}val": "Template Style"})
    ElementTree.SubElement(style, f"{_W}aliases", {f"{_W}val": "Template Alias"})


@pytest.mark.parametrize(
    ("mutation", "expected_code"),
    [
        (_drop_member("[Content_Types].xml"), "required_part_missing"),
        (_drop_member("_rels/.rels"), "required_part_missing"),
        (_drop_member("word/numbering.xml"), "relationship_target_missing"),
        (
            _xml_mutation("word/document.xml", lambda root: _remove_first(root, f"{_W}numPr")),
            "heading_numpr_missing",
        ),
        (_xml_mutation("word/document.xml", _wrong_num_id), "heading_num_missing"),
        (
            _xml_mutation("word/document.xml", lambda root: _remove_first(root, f"{_W}bookmarkEnd")),
            "bookmark_range_unbalanced",
        ),
        (_xml_mutation("word/document.xml", _duplicate_bookmark_name), "bookmark_start_invalid"),
        (_xml_mutation("word/document.xml", _wrong_field_cache), "field_inventory_or_cache_mismatch"),
        (_xml_mutation("word/document.xml", _wrong_seq_instruction), "field_inventory_or_cache_mismatch"),
        (
            _xml_mutation("word/_rels/document.xml.rels", _drop_styles_relationship),
            "styles_relationship_invalid",
        ),
    ],
)
def test_resolved_ooxml_semantic_mutations_fail_closed(
    tmp_path: Path,
    mutation: Callable[[dict[str, bytes]], None],
    expected_code: str,
) -> None:
    fixture = base._fixture(tmp_path)
    harness = producer._load_harness(fixture.docwen)
    case = next(item for item in harness.cases if item.case_id == "rich-semantics-composite")
    payload = support.rewrite_docx(base._docx(case), mutation)
    with pytest.raises(producer.V4PackageInputError, match=expected_code):
        producer._inspect_docx(payload, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)


def test_duplicate_opc_member_fails_closed(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    case = next(
        item for item in producer._load_harness(fixture.docwen).cases if item.case_id == "rich-semantics-composite"
    )
    source = zipfile.ZipFile(io.BytesIO(base._docx(case)))
    target = io.BytesIO()
    with zipfile.ZipFile(target, "w") as archive:
        for item in source.infolist():
            archive.writestr(item.filename, source.read(item))
        with pytest.warns(UserWarning, match="Duplicate name"):
            archive.writestr("word/document.xml", source.read("word/document.xml"))
    source.close()
    with pytest.raises(producer.V4PackageInputError, match="duplicate_zip_member"):
        producer._inspect_docx(
            target.getvalue(),
            case.expected_ooxml,
            case.neutral_envelope,
            case.plan_envelope,
        )


def test_plan_target_inventory_and_disabled_field_mutations_fail_closed(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    harness = producer._load_harness(fixture.docwen)
    rich = next(item for item in harness.cases if item.case_id == "rich-semantics-composite")
    mutated_plan = copy.deepcopy(rich.plan_envelope)
    mutated_plan["plan"]["targets"][0]["target_id"] = "different-heading"
    with pytest.raises(producer.V4PackageInputError, match="target_inventory_mismatch"):
        producer._inspect_docx(
            base._docx(rich),
            rich.expected_ooxml,
            rich.neutral_envelope,
            mutated_plan,
        )

    disabled_neutral, disabled_plan, disabled_expected = support.case_semantics(
        "numbering-figure-off",
        base.SOURCE,
    )

    def inject_field(root: ElementTree.Element) -> None:
        paragraph = next(item for item in root.iter(f"{_W}p") if item.find(f"{_W}pPr/{_W}pStyle") is not None)
        support._field(paragraph, " SEQ Figure \\* ARABIC ", "1")

    payload = support.rewrite_docx(
        support.build_docx(disabled_neutral, disabled_plan),
        _xml_mutation("word/document.xml", inject_field),
    )
    with pytest.raises(producer.V4PackageInputError, match="field_inventory_or_cache_mismatch"):
        producer._inspect_docx(
            payload,
            disabled_expected,
            disabled_neutral,
            disabled_plan,
        )


def _idless_semantics(kind: str, *, enabled: bool = True) -> tuple[dict[str, Any], dict[str, Any], dict[str, int]]:
    suffix = kind.removesuffix("_block")
    neutral, plan, expected = support.case_semantics(
        f"numbering-{suffix}-{'on' if enabled else 'off'}",
        base.SOURCE,
    )
    neutral["document"]["targets"][0]["target_id"] = None
    neutral["document"]["references"] = []
    plan["plan"]["targets"][0]["target_id"] = None
    plan_sha = support._sha(support._json_bytes(plan["plan"]))
    neutral["plan_sha256"] = plan_sha
    plan["plan_sha256"] = plan_sha
    expected["bookmarkCount"] = 0
    expected["refFieldCount"] = 0
    return neutral, plan, expected


@pytest.mark.parametrize("kind", ["heading", "figure", "table", "equation", "code_block"])
def test_enabled_idless_targets_are_proven_by_typed_style_and_source_order(kind: str) -> None:
    neutral, plan, expected = _idless_semantics(kind)
    result = producer._inspect_docx(support.build_docx(neutral, plan), expected, neutral, plan)
    assert result["bookmarkCount"] == 0
    assert result["seqFieldCount"] == (0 if kind == "heading" else 1)


def test_disabled_idless_caption_requires_derived_occurrence_map_and_wrapper() -> None:
    neutral, plan, expected = _idless_semantics("figure", enabled=False)
    assert producer._inspect_docx(support.build_docx(neutral, plan), expected, neutral, plan)["seqFieldCount"] == 0


def test_heading_merge_suffix_preserves_exact_authored_prefix() -> None:
    neutral, plan, expected = _idless_semantics("heading")

    def append_merge(root: ElementTree.Element) -> None:
        paragraph = next(item for item in root.iter(f"{_W}p") if item.find(f"{_W}pPr/{_W}numPr") is not None)
        support._text(paragraph, "Merged body.")

    payload = support.rewrite_docx(support.build_docx(neutral, plan), _xml_mutation("word/document.xml", append_merge))
    producer._inspect_docx(payload, expected, neutral, plan)


def _two_disabled_idless_captions() -> tuple[dict[str, Any], dict[str, Any], dict[str, int]]:
    neutral, plan, _expected = support.case_semantics("rich-semantics-composite", base.SOURCE)
    neutral["document"]["targets"] = neutral["document"]["targets"][1:3]
    neutral["document"]["references"] = []
    neutral["document"]["citations"] = []
    plan["plan"]["targets"] = plan["plan"]["targets"][1:3]
    plan["plan"]["heading_definitions"] = []
    plan["plan"]["heading_instances"] = []
    for target, planned in zip(neutral["document"]["targets"], plan["plan"]["targets"], strict=True):
        target["target_id"] = None
        planned.update({"target_id": None, "enabled": False, "derived_number": None, "materialization": None})
    plan_sha = support._sha(support._json_bytes(plan["plan"]))
    neutral["plan_sha256"] = plan_sha
    plan["plan_sha256"] = plan_sha
    return (
        neutral,
        plan,
        {
            "abstractNumCount": 0,
            "numCount": 0,
            "bookmarkCount": 0,
            "seqFieldCount": 0,
            "styleRefFieldCount": 0,
            "refFieldCount": 0,
            "citationFieldCount": 0,
        },
    )


def test_occurrence_tag_exchange_and_extra_managed_candidate_fail_closed() -> None:
    neutral, plan, expected = _two_disabled_idless_captions()
    original = support.build_docx(neutral, plan)

    def exchange_tags(root: ElementTree.Element) -> None:
        tags = [
            item
            for item in root.iter(f"{_W}tag")
            if (item.get(f"{_W}val") or "").startswith("docwen-numbering-occurrence-v1:")
        ]
        left, right = tags[0].get(f"{_W}val"), tags[1].get(f"{_W}val")
        tags[0].set(f"{_W}val", str(right))
        tags[1].set(f"{_W}val", str(left))

    exchanged = support.rewrite_docx(original, _xml_mutation("word/document.xml", exchange_tags))
    with pytest.raises(producer.V4PackageInputError, match="disabled_idless_caption_carrier_invalid"):
        producer._inspect_docx(exchanged, expected, neutral, plan)

    def duplicate_caption(root: ElementTree.Element) -> None:
        body = root.find(f"{_W}body")
        paragraph = next(
            item
            for item in root.iter(f"{_W}p")
            if (style := item.find(f"{_W}pPr/{_W}pStyle")) is not None
            and (style.get(f"{_W}val") or "").startswith("Caption")
        )
        assert body is not None
        body.append(copy.deepcopy(paragraph))

    ambiguous = support.rewrite_docx(original, _xml_mutation("word/document.xml", duplicate_caption))
    with pytest.raises(producer.V4PackageInputError, match="managed_target_paragraph_inventory_invalid"):
        producer._inspect_docx(ambiguous, expected, neutral, plan)


def test_global_five_kind_source_order_rejects_heading_moved_after_captions(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    case = next(
        item for item in producer._load_harness(fixture.docwen).cases if item.case_id == "rich-semantics-composite"
    )
    payload = support.rewrite_docx(
        base._docx(case),
        _xml_mutation("word/document.xml", _move_heading_after_caption_targets),
    )
    with pytest.raises(producer.V4PackageInputError, match="target_kind_style_mismatch"):
        producer._inspect_docx(payload, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)


def test_managed_style_alias_and_noncanonical_owned_custom_xml_fail_closed(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    case = next(
        item for item in producer._load_harness(fixture.docwen).cases if item.case_id == "rich-semantics-composite"
    )
    original = base._docx(case)
    aliased = support.rewrite_docx(original, _xml_mutation("word/styles.xml", _add_managed_style_alias))
    with pytest.raises(producer.V4PackageInputError, match="managed_style_alias_rejected"):
        producer._inspect_docx(aliased, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)

    def add_evil_owned_signal(parts: dict[str, bytes]) -> None:
        parts["customXml/evil.xml"] = f'<evil xmlns="{support.CAPTION_STYLE_NAMESPACE}"/>'.encode()

    evil = support.rewrite_docx(original, add_evil_owned_signal)
    with pytest.raises(
        producer.V4PackageInputError,
        match="owned_custom_xml_signal_in_noncanonical_part",
    ):
        producer._inspect_docx(evil, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)


def test_unmanaged_template_style_alias_is_allowed(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    case = next(
        item for item in producer._load_harness(fixture.docwen).cases if item.case_id == "rich-semantics-composite"
    )
    payload = support.rewrite_docx(
        base._docx(case),
        _xml_mutation("word/styles.xml", _add_unmanaged_style_alias),
    )
    producer._inspect_docx(payload, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)


def test_extra_owned_carrier_and_spoofed_owned_bookmark_fail_closed(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    case = next(
        item for item in producer._load_harness(fixture.docwen).cases if item.case_id == "rich-semantics-composite"
    )
    original = base._docx(case)

    def extra_carrier(root: ElementTree.Element) -> None:
        body = root.find(f"{_W}body")
        assert body is not None
        support._sdt(body, f"docwen-target-v1:{'f' * 32}")

    payload = support.rewrite_docx(original, _xml_mutation("word/document.xml", extra_carrier))
    with pytest.raises(producer.V4PackageInputError, match="target_sdt_inventory_invalid"):
        producer._inspect_docx(payload, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)

    def spoof_bookmark(root: ElementTree.Element) -> None:
        body = root.find(f"{_W}body")
        assert body is not None
        support._bookmark(support._sub(body, "p"), 999, "DW_T_spoof")

    payload = support.rewrite_docx(original, _xml_mutation("word/document.xml", spoof_bookmark))
    with pytest.raises(producer.V4PackageInputError, match="owned_bookmark_identity_noncanonical"):
        producer._inspect_docx(payload, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)


def test_unrelated_template_bookmark_and_page_field_are_allowed_but_owned_ref_is_not(tmp_path: Path) -> None:
    fixture = base._fixture(tmp_path)
    case = next(
        item for item in producer._load_harness(fixture.docwen).cases if item.case_id == "rich-semantics-composite"
    )
    original = base._docx(case)

    def template_signals(root: ElementTree.Element) -> None:
        body = root.find(f"{_W}body")
        assert body is not None
        paragraph = support._sub(body, "p")
        support._bookmark(paragraph, 999, "TemplateBookmark")
        support._field(paragraph, " PAGE ", "7")

    payload = support.rewrite_docx(original, _xml_mutation("word/document.xml", template_signals))
    producer._inspect_docx(payload, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)

    owned_name = identity.target_identity("heading", "heading-one")[1]

    def spoof_ref(root: ElementTree.Element) -> None:
        body = root.find(f"{_W}body")
        assert body is not None
        support._field(support._sub(body, "p"), f" REF {owned_name} \\h ", "1")

    payload = support.rewrite_docx(original, _xml_mutation("word/document.xml", spoof_ref))
    with pytest.raises(producer.V4PackageInputError, match="owned_field_outside_managed_carrier"):
        producer._inspect_docx(payload, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)
