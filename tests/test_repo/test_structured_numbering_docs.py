"""Guards for the v4 resolved-numbering DocWen boundary."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
SPECS = ROOT / "docs" / "specs"


def _read(name: str) -> str:
    return (SPECS / name).read_text(encoding="utf-8")


def _normalized(name: str) -> str:
    return " ".join(_read(name).split())


def test_only_derived_numbering_is_authoritative_and_markdown_is_immutable() -> None:
    plan = _normalized("structured-numbering-phases.md")
    markdown = _normalized("markdown-compatibility.md")

    for text in (plan, markdown):
        assert "`## 2.3 标题`" in text
        assert "`2.3 标题`" in text
        assert "zero Markdown" in text
        assert "`第二章`" in text and "`一、`" in text

    assert "There is no v4 cleanup command and no authored/manual-number source" in plan
    assert "derived prefix and the complete authored title may both be visible" in markdown
    for declaration in ("`Figure:`", "`Table:`", "`Equation:`", "`Code:`"):
        assert declaration in plan


def test_resolved_plan_keeps_numbering_rules_out_of_docwen() -> None:
    plan = _normalized("structured-numbering-phases.md")
    machine = _normalized("machine-protocol-v1.md")

    for kind in ("`heading`", "`figure`", "`table`", "`equation`", "`code_block`"):
        assert kind in plan
    for rule in (
        "enablement",
        "format",
        "localized label",
        "start/reset",
        "document-wide versus Heading scope",
        "chapter-number inclusion",
        "separator",
    ):
        assert rule in plan

    assert "provider-neutral resolved document plus a resolved numbering/export plan" in plan
    assert "Workspace paths, Node IDs, consumer types, resolver instructions" in plan
    assert "DocWen does not repair or recompute the plan" in plan
    assert "Machine options may not carry the plan as an opaque JSON string" in plan
    assert "provider-neutral resolved document plus a separately versioned" in machine


def test_enable_disable_and_template_empty_are_one_materialization_state() -> None:
    plan = _normalized("structured-numbering-phases.md")
    markdown = _normalized("markdown-compatibility.md")

    assert "v1 has no separate state in which a number is hidden" in plan
    assert "Heading level whose selected template is empty is disabled" in plan
    assert "docwen.markdown.cross_reference.unnumbered_target" in plan
    assert "ordinary `[[...]]` remains a navigation link" in plan
    assert "does not display Alias without a number" in markdown

    for label in (
        "Heading enabled",
        "Heading disabled",
        "Heading-level template empty",
        "Figure enabled",
        "Figure disabled",
        "Table enabled",
        "Table disabled",
        "Equation enabled",
        "Equation disabled",
        "Code enabled",
        "Code disabled",
    ):
        assert f"| {label} |" in _read("structured-numbering-phases.md")


def test_docwen_materializes_only_proved_docx_semantics() -> None:
    plan = _normalized("structured-numbering-phases.md")
    markdown = _normalized("markdown-compatibility.md")
    styles = _normalized("templates-and-styles.md")

    assert "`word/numbering.xml`" in plan
    assert "caption `SEQ`" in plan
    assert r"`REF <bookmark> \n \h`" in plan
    assert "zero-width bookmark at the start of the caption paragraph" in plan
    assert "valid `numbering.xml` plus paragraph numbering bindings" in plan
    assert "visible prefix without sufficient Word proof remains authored text" in plan
    assert "calling provider or consumer decides the final Markdown representation" in plan
    assert "disabled caption" in styles and "contain no `SEQ`" in styles
    assert "An unnumbered semantic reference fails before rendering" in markdown


def test_machine_freezes_exact_numbering_inputs_and_distinct_plan_failures() -> None:
    machine = _normalized("machine-protocol-v1.md")

    for token in (
        "`neutral_document`",
        "`application/vnd.docwen.resolved-document+json`",
        "`docwen.resolved_document.v1`",
        "`urn:docwen:schema:resolved-document:v1`",
        "`numbering_export_plan`",
        "`application/vnd.docwen.numbering-export-plan+json`",
        "`docwen.numbering_export_plan.v1`",
        "`urn:docwen:schema:numbering-export-plan:v1`",
    ):
        assert token in machine
    assert machine.count("exactly 1") >= 2
    assert "same `input_id`, `source_sha256`, and `plan_sha256`" in machine
    assert "RFC 8785 canonical UTF-8 bytes of the closed `plan` member only" in machine
    for code in (
        "`docwen.numbering_export_plan.missing`",
        "`docwen.numbering_export_plan.invalid`",
        "`docwen.numbering_export_plan.unsupported_materialization`",
    ):
        assert code in machine
    assert "distinct from a valid disabled target" in machine


def test_closed_portable_heading_and_caption_materialization_is_frozen() -> None:
    plan = _normalized("structured-numbering-phases.md")

    for number_format in (
        "`chinese_lower`",
        "`chinese_upper`",
        "`arabic_half`",
        "`arabic_full`",
        "`arabic_circled`",
        "`letter_upper`",
        "`letter_lower`",
        "`roman_upper`",
        "`roman_lower`",
    ):
        assert number_format in plan
    assert '`{"counter":{"level":N,"number_format":"..."}}`' in plan
    assert "Heading paragraphs contain no separate cached-number run" in plan
    for action in ("`continue|reset_to_start|restart_by_heading_level`", "`simple_seq|chapter_seq`"):
        assert action in plan
    assert r"` SEQ <counter> \r <start_value> \* <switch> `" in plan
    assert r"` SEQ <counter> \s N \* <switch> `" in plan
    assert '` STYLEREF "<resolved-chapter-heading-N-name>"' in plan
    assert "all six closed combinations" in plan
    assert "may differ" in plan
    assert "chapter_cached_number + chapter_separator + sequence_cached_number" in plan
    assert "never splits a composite number" in plan
    assert "heading restart with any start other than 1" in plan
    assert "two independent sequence scopes" in plan
    assert "Update Fields in Word/WPS/LibreOffice" in plan


def test_exact_two_port_embeds_closed_resolved_dependencies() -> None:
    plan = _normalized("structured-numbering-phases.md")
    machine = _normalized("machine-protocol-v1.md")

    assert "authored_markdown,targets,references,resource_occurrences,citations,resources" in plan
    assert "source_start,source_end,source_slice_sha256,authored_token,authored_locator,resource_id" in plan
    assert "same locator at different ranges may resolve to different resource IDs" in plan
    assert "resource_id,role,media_type,size_bytes,sha256,content_base64" in plan
    assert "application/vnd.docwen.semantic-bibliography+json" in plan
    assert "Decoded resources total at most 6,000,000 bytes" in plan
    assert "performs no base-path" in plan
    assert "global locator replacement" in plan
    assert "or filesystem resolution" in plan
    assert "resource_occurrences,citations,resources" in machine
    assert "accepts no independent bibliography or `citation_style` slot" in machine
    assert "may accept one already-presented bibliography resource" not in machine
    assert "The upstream owner chooses profiles, counters, enablement" in machine
    assert "materializes only resolved Heading list" in machine


def test_resolved_citations_keep_provider_identity_behind_closed_word_addresses() -> None:
    markdown = _normalized("markdown-compatibility.md")
    plan = _normalized("structured-numbering-phases.md")
    golden = _normalized("golden-regression-suite.md")

    for token in (
        "https://docwen.dev/schema/document-citation-item-map/v1",
        "https://docwen.dev/schema/document-citation-occurrence-map/v1",
        "DWCIT_<digest32>",
        "_DWC_<digest35>",
        "docwen-citation-item-map-v1\\0<source_sha256>\\0<record_id>\\0<record_sha256>\\0<presentation_sha256>",
        "docwen-citation-item-ref-v1\\0<citation_key>\\0<word_tag>\\0<item_sha256>",
        "docwen-citation-occurrence-map-v1\\0<source_sha256>\\0<source_start>\\0<source_end>",
        "word_tag,record_id,record_sha256,presentation_base64,presentation_sha256,sha256",
        "citation_key,word_tag,item_sha256,sha256",
        "exactly one item map and exactly one non-empty occurrence map",
        "orphan, dangling, or unused records fail closed",
        'w:fldLock="true"',
        "with no `w:dirty`",
    ):
        assert token in markdown
    assert "exactly 38 ASCII characters" in markdown
    assert "bookmark is exactly 40 ASCII characters" in markdown
    assert "full or truncated item-tag, SDT-tag, bookmark-name" in markdown
    assert "reference-record:98" in plan
    assert "never rejected, truncated, or replaced by its authored key" in plan
    assert "key remains a lookup key" in golden
    for host in ("Word", "WPS", "LibreOffice"):
        assert host in markdown and host in golden


def test_disabled_idless_caption_has_executable_occurrence_authority() -> None:
    markdown = _normalized("markdown-compatibility.md")

    assert "https://docwen.dev/schema/document-numbering-occurrence-map/v1" in markdown
    assert "documentNumberingOccurrenceMap" in markdown
    assert "docwen-numbering-occurrence-v1:<digest32>" in markdown
    assert (
        "docwen-numbering-occurrence-map-v1\\0<source_sha256>\\0<source_start>\\0<source_end>"
        "\\0<kind>\\0false\\0\\0\\0<plan_sha256>"
    ) in markdown
    assert (
        "tag,source_sha256,source_start,source_end,kind,enabled,target_id,derived_number,plan_sha256,sha256" in markdown
    )
    assert "independently allocated canonical custom-XML trio" in markdown
    assert "exactly the caption paragraph and one matching logical object" in markdown
    assert "It creates no target, hidden ID, bookmark, `SEQ`, `REF`, or number" in markdown


def test_unnumbered_provider_mapping_is_one_to_one() -> None:
    markdown = _normalized("markdown-compatibility.md")
    machine = _normalized("machine-protocol-v1.md")

    for text in (markdown, machine):
        assert "`interop.cross_reference.unnumbered_target`" in text
        assert "`docwen.markdown.cross_reference.unnumbered_target`" in text
        assert "one-to-one" in text
    assert "must never be coerced into `unnumbered_target`" in markdown


def test_source_authoring_options_are_not_resolved_plan_inputs() -> None:
    machine = _normalized("machine-protocol-v1.md")
    capabilities = " ".join((ROOT / "docs" / "capabilities.md").read_text(encoding="utf-8").split())

    assert "final v4 DOCX-to-Markdown capability has exactly these seven properties" in machine
    assert "required=[]` and `additionalProperties=false" in machine
    for option in ("`remove_numbering`", "`add_numbering`", "`numbering_scheme`"):
        assert option in machine
        assert option in capabilities
    assert "are not inputs to the resolved-plan Conversion Port" in machine
    assert "在 resolved route 中全部拒绝" in capabilities
    readme = " ".join((ROOT / "docs" / "README.md").read_text(encoding="utf-8").split())
    assert "resolved-plan capability rejects" in readme
    assert "guarded by an explicit negative capability gate" in readme


def test_numbering_acceptance_keeps_evidence_layers_separate() -> None:
    plan = _read("structured-numbering-phases.md")
    golden = _read("golden-regression-suite.md")
    machine = _read("machine-protocol-v1.md")

    for layer in (
        "`source_oracle`",
        "`packaged`",
        "`headless_ooxml`",
        "`roundtrip`",
        "`word_host`",
        "`wps_host`",
        "`libreoffice_host`",
    ):
        assert layer in plan
        assert layer in machine
    assert "current schema/URN/media identities" in golden
    assert "Toggling each kind on→off→on produces a byte-identical Markdown SHA-256" in golden
    assert "headless XML or another host" in plan
    assert "Upstream-provider write/read observations remain provider evidence" in machine
