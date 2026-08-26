"""Source-oracle gates for the v4 fenced-source occurrence carrier."""

from __future__ import annotations

import base64
import hashlib

import pytest

from docwen_core.docx_semantics_v3 import fenced_source_identity_from_mapping_v3
from docwen_plugin_markdown.document_semantics_v3 import analyze_markdown_semantics_v3
from docwen_plugin_markdown.document_semantics_v3_fenced_source import (
    fenced_source_info_insertion_offset_v3,
    recover_fenced_logical_body_v3,
)

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("source", "expected_info", "expected_prefixes", "expected_body", "closing_state"),
    [
        (
            "  ```rust  exact\r\n  alpha\r\n beta\n    gamma\r\n  ```  \n",
            "rust  exact",
            ("  ", " ", "  "),
            "alpha\r\nbeta\n  gamma\r\n",
            "present",
        ),
        (
            "> ```mermaid\r\n> graph TD\r\n> ```\r\n",
            "mermaid",
            ("> ",),
            "graph TD\r\n",
            "present",
        ),
        (
            "- ```query\n  tag:#x\n  ```\n",
            "query",
            ("  ",),
            "tag:#x\n",
            "present",
        ),
        ("```view\nbody", "view", ("",), "body", "omitted_eof"),
    ],
)
def test_fenced_projection_binds_exact_framing_and_logical_body(
    source: str,
    expected_info: str,
    expected_prefixes: tuple[str, ...],
    expected_body: str,
    closing_state: str,
) -> None:
    analysis = analyze_markdown_semantics_v3(source, input_id="fenced.md")

    assert not analysis.has_errors
    [record] = analysis.projection["fenced_sources"]
    identity = fenced_source_identity_from_mapping_v3(record)
    prefixes = base64.b64decode(record["body_prefixes_b64"], validate=True).split(b"\0")

    assert identity.source_start == 0
    assert identity.source_end == len(source)
    assert base64.b64decode(record["info_b64"], validate=True).decode() == expected_info
    assert tuple(item.decode() for item in prefixes) == expected_prefixes
    assert record["closing_state"] == closing_state
    assert recover_fenced_logical_body_v3(source, record) == expected_body
    assert record["body_sha256"] == hashlib.sha256(expected_body.encode()).hexdigest()
    assert fenced_source_info_insertion_offset_v3(source, record) == (
        len(base64.b64decode(record["opening_prefix_b64"]).decode()) + record["opening_length"] + len(expected_info)
    )


def test_fenced_projection_is_independent_of_caption_and_ordinary_anchor_ownership() -> None:
    source = "Code: Example ^code-target\n\n```rust\nfn main() {}\n```\n\n^raw-code\n"
    analysis = analyze_markdown_semantics_v3(source, input_id="captioned.md")

    assert not analysis.has_errors
    assert [(target["id"], target["kind"]) for target in analysis.projection["targets"]] == [
        ("code-target", "code_block")
    ]
    assert [(anchor["id"], anchor["block_kind"]) for anchor in analysis.projection["anchors"]] == [
        ("raw-code", "code_block")
    ]
    [record] = analysis.projection["fenced_sources"]
    assert not {
        "source_id",
        "target_id",
        "bookmark_name",
        "anchor_id",
        "kind",
        "block_kind",
    }.intersection(record)
    assert recover_fenced_logical_body_v3(source, record) == "fn main() {}\n"


@pytest.mark.parametrize(
    "mutation",
    [
        lambda source, _record: source.replace("body", "evil"),
        lambda source, record: source[: record["source_start"]] + "x" + source[record["source_start"] :],
    ],
)
def test_fenced_runtime_proof_rejects_source_tamper(mutation) -> None:
    source = "```rust\nbody\n```\n"
    [record] = analyze_markdown_semantics_v3(source, input_id="tamper.md").projection["fenced_sources"]

    with pytest.raises(ValueError, match=r"source|hash|bound"):
        recover_fenced_logical_body_v3(mutation(source, record), record)


def test_fenced_projection_is_occurrence_bound_for_identical_blocks() -> None:
    source = "```rust\nsame\n```\n\n```rust\nsame\n```\n"
    records = analyze_markdown_semantics_v3(source, input_id="duplicates.md").projection["fenced_sources"]

    assert len(records) == 2
    assert records[0]["body_sha256"] == records[1]["body_sha256"]
    assert records[0]["block_sha256"] == records[1]["block_sha256"]
    assert records[0]["identity_sha256"] != records[1]["identity_sha256"]
    assert records[0]["tag"] != records[1]["tag"]


def test_four_space_indented_quote_marker_is_not_recursively_reparsed_as_quote() -> None:
    source = "    > literal code-like text\n"

    analysis = analyze_markdown_semantics_v3(source, input_id="indented-quote.md")

    assert not analysis.has_errors
    assert analysis.projection["fenced_sources"] == []


def test_backtick_in_backtick_fence_info_is_not_projected_as_a_fenced_source() -> None:
    source = "```rust`invalid\nbody\n```\n"

    analysis = analyze_markdown_semantics_v3(source, input_id="invalid-opener.md")

    assert not analysis.has_errors
    # The invalid first line stays ordinary paragraph text.  Its later bare
    # backticks are independently a valid empty fence opener at EOF.
    [record] = analysis.projection["fenced_sources"]
    assert record["source_start"] == source.rindex("```")
    assert record["info_b64"] == ""
    assert record["closing_state"] == "omitted_eof"


@pytest.mark.parametrize(
    ("source", "expected_range"),
    [
        ("```rust\nbody\n``` ^bad-code", {"start": 17, "end": 26}),
        ("```rust\nbody\n``` ^bad-code\n", {"start": 17, "end": 26}),
        ("```rust\r\nbody\r\n``` ^bad-code\r\n", {"start": 19, "end": 28}),
    ],
)
def test_eof_pseudo_closer_never_becomes_an_omitted_eof_carrier(
    source: str,
    expected_range: dict[str, int],
) -> None:
    analysis = analyze_markdown_semantics_v3(source, input_id="pseudo-closer-eof.md")

    assert analysis.has_errors
    assert analysis.projection["fenced_sources"] == []
    assert [(item["code"], item["range"]) for item in analysis.diagnostics] == [
        ("docwen.markdown.anchor.invalid_id", expected_range),
    ]


@pytest.mark.parametrize(
    "source",
    [
        "```rust\rbody\r```\r",
        "```rust\nbody\u2028mutated\n```\n",
        "```rust\nbody\u2029mutated\n```\n",
    ],
)
def test_unsupported_line_separators_are_rejected_before_fenced_projection(source: str) -> None:
    with pytest.raises(
        ValueError,
        match=r"Markdown semantics v3 accepts only LF and CRLF line endings \(unsupported separator at offset \d+\)",
    ):
        analyze_markdown_semantics_v3(source, input_id="unsupported-eol.md")
