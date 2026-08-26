"""Focused source-oracle tests for ``docwen.markdown_semantics.v3``."""

from __future__ import annotations

import hashlib

import pytest

from docwen_plugin_markdown.document_semantics_v3 import (
    ExternalCitationResolution,
    ExternalReferenceResolution,
    analyze_markdown_semantics_v3,
    is_resource_less_image_carrier_v3,
    render_caption_declaration,
    render_citation_token,
    render_cross_reference_token,
    select_safe_fence,
)

pytestmark = pytest.mark.contract


def _analyze(source: str, **kwargs):
    return analyze_markdown_semantics_v3(source, input_id="fixture.md", **kwargs)


@pytest.mark.parametrize("level", [7, 8, 9])
def test_docwen_extended_heading_levels_are_semantic_targets(level: int) -> None:
    source = f"{'#' * level} Deep heading ^deep-{level}\n"

    analysis = _analyze(source)

    assert not analysis.has_errors
    [target] = analysis.projection["targets"]
    assert (target["kind"], target["heading_level"], target["id"]) == (
        "heading",
        level,
        f"deep-{level}",
    )


def test_untyped_ids_are_owned_by_structure_and_legacy_bare_at_is_citation() -> None:
    source = """# Introduction ^section-a

Figure: System architecture ^visual-a

![[image.png]] ^fig-legacy

See [[Paper#^fig-legacy]], ![[Paper#^fig-legacy]], @[[#^visual-a|System architecture]], and @fig-legacy.
"""
    analysis = _analyze(source)

    assert not analysis.has_errors
    assert [(item["kind"], item.get("id")) for item in analysis.projection["targets"]] == [
        ("heading", "section-a"),
        ("figure", "visual-a"),
    ]
    assert analysis.projection["anchors"] == [
        {
            "id": "fig-legacy",
            "block_kind": "image",
            "placement": "inline",
            "range": {
                "start": source.index("^fig-legacy"),
                "end": source.index("^fig-legacy") + len("^fig-legacy"),
            },
            "block_range": {
                "start": source.index("![[image.png]]"),
                "end": source.index("![[image.png]]") + len("![[image.png]] ^fig-legacy\n"),
            },
            "container_path": [],
        }
    ]
    assert [item["kind"] for item in analysis.projection["links"]] == ["embed", "link", "embed"]
    assert analysis.projection["references"][0]["resolution_status"] == "resolved"
    assert analysis.projection["references"][0]["resolved_kind"] == "figure"
    assert analysis.projection["citations"][0]["items"][0] == {
        "key": "fig-legacy",
        "key_range": {
            "start": source.rindex("fig-legacy"),
            "end": source.rindex("fig-legacy") + len("fig-legacy"),
        },
        "resolution_status": "unresolved",
    }
    assert "{#" not in "".join(analysis.authored_tokens())


def test_four_caption_kinds_and_idless_targets_are_numbered() -> None:
    source = """Figure: Figure text

![image](image.png)

Table: Table text ^shared-name

| A |
|---|
| 1 |

Equation: Equation text

$$x=1$$

Code: Code text ^code-name

```python
print(1)
```
"""
    analysis = _analyze(source)

    assert not analysis.has_errors
    assert [(item["kind"], item.get("id"), item["number"]) for item in analysis.projection["targets"]] == [
        ("figure", None, "1"),
        ("table", "shared-name", "1"),
        ("equation", None, "1"),
        ("code_block", "code-name", "1"),
    ]
    assert {item["source_keyword"] for item in analysis.projection["targets"]} == {
        "Figure",
        "Table",
        "Equation",
        "Code",
    }


def test_resource_less_image_carrier_preserves_figure_and_ordinary_image_ownership() -> None:
    source = """Figure: Recovered illustration ^figure-owner

![image omitted]()

![image omitted]() ^ordinary-owner
"""

    analysis = _analyze(source)

    assert is_resource_less_image_carrier_v3("![image omitted]()")
    assert not analysis.has_errors
    [figure] = analysis.projection["targets"]
    assert (figure["kind"], figure["id"], figure["title"]) == (
        "figure",
        "figure-owner",
        "Recovered illustration",
    )
    assert source[figure["object_range"]["start"] : figure["object_range"]["end"]] == "![image omitted]()\n"
    [anchor] = analysis.projection["anchors"]
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == (
        "ordinary-owner",
        "image",
        "inline",
    )
    assert source[anchor["block_range"]["start"] : anchor["block_range"]["end"]] == (
        "![image omitted]() ^ordinary-owner\n"
    )
    assert analysis.projection["source"]["sha256"] == hashlib.sha256(source.encode("utf-8")).hexdigest()


@pytest.mark.parametrize(
    ("near_miss", "ordinary_image"),
    [
        ("![Image omitted]()", False),
        ("![image omitted ]()", False),
        ("![image-omitted]()", False),
        ("![image omitted]( )", True),
        ("![image omitted](missing.png)", True),
    ],
)
def test_only_exact_empty_destination_carrier_has_resource_less_image_semantics(
    near_miss: str,
    ordinary_image: bool,
) -> None:
    source = f"Figure: Recovered illustration ^figure-owner\n\n{near_miss}\n"

    analysis = _analyze(source)

    assert not is_resource_less_image_carrier_v3(near_miss)
    assert analysis.has_errors is (not ordinary_image)
    if ordinary_image:
        assert analysis.projection["targets"][0]["kind"] == "figure"
    else:
        assert analysis.projection["targets"] == []
        assert analysis.diagnostics[0]["code"] == "docwen.markdown.caption.object_mismatch"


def test_structured_post_block_anchors_bind_complete_blocks() -> None:
    source = """- one
- two

^whole-list

> quote
> continuation
^quote-anchor

> [!note] Callout
> body

^callout-anchor

| A |
|---|
| 1 |
^table-anchor

$$
x=1
$$

^equation-anchor

```python
print(1)
```
^code-anchor

```mermaid
graph TD
```

^mermaid-anchor
"""
    analysis = _analyze(source)

    assert not analysis.has_errors
    assert [(item["id"], item["block_kind"], item["placement"]) for item in analysis.projection["anchors"]] == [
        ("whole-list", "list", "post_block"),
        ("quote-anchor", "block_quote", "post_block"),
        ("callout-anchor", "callout", "post_block"),
        ("table-anchor", "table", "post_block"),
        ("equation-anchor", "equation", "post_block"),
        ("code-anchor", "code_block", "post_block"),
        ("mermaid-anchor", "fenced_block", "post_block"),
    ]
    for item in analysis.projection["anchors"]:
        block = source[item["block_range"]["start"] : item["block_range"]["end"]]
        assert item["id"] not in block


def test_every_list_item_position_can_own_an_inline_anchor() -> None:
    source = "- first ^first-item\n- middle ^middle-item\n- last ^last-item\n"

    analysis = _analyze(source)

    assert not analysis.has_errors
    assert [(item["id"], item["block_kind"], item["placement"]) for item in analysis.projection["anchors"]] == [
        ("first-item", "list_item", "inline"),
        ("middle-item", "list_item", "inline"),
        ("last-item", "list_item", "inline"),
    ]


def test_nested_commonmark_containers_preserve_semantic_ownership_and_ranges() -> None:
    source = """> # 标题😀 ^heading-id
>
> Table: 数据 @citation [[Page#Heading]] ^table-target
>
> | A |
> |---|
> | 1 |
>
> ^raw-table
"""

    analysis = _analyze(source)

    assert not analysis.has_errors
    assert [(item["kind"], item.get("id")) for item in analysis.projection["targets"]] == [
        ("heading", "heading-id"),
        ("table", "table-target"),
    ]
    assert analysis.projection["targets"][0]["title"] == "标题😀"
    heading_id_start = source.index("^heading-id")
    assert analysis.projection["targets"][0]["id_range"] == {
        "start": heading_id_start,
        "end": heading_id_start + len("^heading-id"),
    }
    [anchor] = analysis.projection["anchors"]
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == (
        "raw-table",
        "table",
        "post_block",
    )
    assert source[anchor["range"]["start"] : anchor["range"]["end"]] == "^raw-table"
    assert source[anchor["block_range"]["start"] : anchor["block_range"]["end"]].startswith("> | A |")
    assert [item["raw"] for item in analysis.projection["links"]] == ["[[Page#Heading]]"]
    assert [item["raw"] for item in analysis.projection["citations"]] == ["@citation"]


def test_nested_list_and_callout_anchors_bind_at_their_exact_sibling_depth() -> None:
    source = """- parent
  - first ^first-item
  - second

  ^nested-list

> [!NOTE] Nested source
> ```mermaid
> graph TD
> ```
>
> ^raw-graph
"""

    analysis = _analyze(source)

    assert not analysis.has_errors
    assert [(item["id"], item["block_kind"], item["placement"]) for item in analysis.projection["anchors"]] == [
        ("first-item", "list_item", "inline"),
        ("nested-list", "list", "post_block"),
        ("raw-graph", "fenced_block", "post_block"),
    ]
    for anchor in analysis.projection["anchors"]:
        assert source[anchor["range"]["start"] : anchor["range"]["end"]] == f"^{anchor['id']}"


def test_nested_paragraph_anchor_is_not_upgraded_to_its_quote_container() -> None:
    source = "> 普通段落😀 ^inside\n"

    analysis = _analyze(source)

    assert not analysis.has_errors
    [anchor] = analysis.projection["anchors"]
    assert (anchor["id"], anchor["block_kind"], anchor["placement"]) == (
        "inside",
        "paragraph",
        "inline",
    )
    assert source[anchor["range"]["start"] : anchor["range"]["end"]] == "^inside"


def test_nested_structured_inline_anchor_reports_one_exact_diagnostic() -> None:
    source = "> | A |\n> |---|\n> | 1 | ^bad-table\n"

    analysis = _analyze(source)

    [diagnostic] = analysis.diagnostics
    assert diagnostic["code"] == "docwen.markdown.anchor.invalid_id"
    assert source[diagnostic["range"]["start"] : diagnostic["range"]["end"]] == "^bad-table"


def test_single_list_item_can_contain_a_heading_semantic_target() -> None:
    source = "- # Nested heading ^nested-heading\n"

    analysis = _analyze(source)

    assert not analysis.has_errors
    [heading] = analysis.projection["targets"]
    assert (heading["kind"], heading["id"], heading["title"]) == (
        "heading",
        "nested-heading",
        "Nested heading",
    )
    assert source[heading["id_range"]["start"] : heading["id_range"]["end"]] == "^nested-heading"


def test_yaml_front_matter_is_masked_but_diagnostic_identity_is_full_source() -> None:
    source = '\ufeff---\ntitle: "@yaml [[Page#Heading]] ^bad_id"\n---\n\n正文😀 ^bad_id\n'

    analysis = _analyze(source)

    [diagnostic] = analysis.diagnostics
    body_token_start = source.rindex("^bad_id")
    assert diagnostic["range"] == {"start": body_token_start, "end": body_token_start + len("^bad_id")}
    assert diagnostic["source"]["sha256"] == hashlib.sha256(source.encode("utf-8")).hexdigest()
    assert analysis.projection["links"] == []
    assert analysis.projection["citations"] == []


def test_stable_soft_and_neutral_external_resolution_stay_separate() -> None:
    source = """# Methods ^methods-id

## Sample

@[[#Methods#Sample|Old sample]]
@[[Other#^external-id|External title]]
@[[Other#Parent#Child]]
@external-cite
"""
    document_sha = hashlib.sha256(b"external document").hexdigest()
    citation_sha = hashlib.sha256(b"external citation record").hexdigest()
    analysis = _analyze(
        source,
        external_references=(
            ExternalReferenceResolution(
                page_locator="Other",
                selector_kind="stable_id",
                target_id="external-id",
                resolved_document_id="document:stable-1",
                resolved_document_sha256=document_sha,
                resolved_kind="table",
                cached_number="4",
                current_title="External title",
            ),
            ExternalReferenceResolution(
                page_locator="Other",
                selector_kind="heading_path",
                heading_path=("Parent", "Child"),
                resolved_document_id="document:stable-1",
                resolved_document_sha256=document_sha,
                resolved_kind="heading",
                cached_number="2.3",
                current_title="Child",
            ),
        ),
        external_citations=(
            ExternalCitationResolution(
                key="external-cite",
                record_id="reference-record:98",
                record_sha256=citation_sha,
                presentation="External, 2026",
            ),
        ),
    )

    assert [item["resolution_status"] for item in analysis.projection["references"]] == [
        "resolved",
        "resolved",
        "resolved",
    ]
    assert analysis.projection["references"][0]["resolved_kind"] == "heading"
    assert analysis.projection["references"][1]["resolved_kind"] == "table"
    assert analysis.projection["references"][2]["heading_path"] == ["Parent", "Child"]
    citation = analysis.projection["citations"][0]["items"][0]
    assert citation["key"] == "external-cite"
    assert citation["resolved_record_id"] == "reference-record:98"
    assert citation["resolved_record_sha256"] == citation_sha
    assert citation["key"] != citation["resolved_record_id"]
    assert [(item["severity"], item["code"]) for item in analysis.diagnostics] == [
        ("warning", "docwen.markdown.cross_reference.alias_stale")
    ]


def test_soft_heading_resolution_is_unique_or_fails_closed() -> None:
    source = """# One
## Repeated
# Two
## Repeated

@[[#Repeated]] @[[#One#Repeated]] @[[#Missing]] @[[Page#Repeated]]
"""
    analysis = _analyze(source)

    assert [item["resolution_status"] for item in analysis.projection["references"]] == [
        "ambiguous",
        "resolved",
        "missing",
        "external_unresolved",
    ]
    assert [item["code"] for item in analysis.diagnostics] == [
        "docwen.markdown.cross_reference.ambiguous",
        "docwen.markdown.cross_reference.missing",
    ]
    ambiguous = analysis.diagnostics[0]
    assert len(ambiguous["related_ranges"]) == 2


@pytest.mark.parametrize("invalid_id", ["bad.id", "bad/id", "bad_id", "x" * 129])
def test_invalid_anchor_has_exact_authenticated_range(invalid_id: str) -> None:
    source = f"Paragraph ^{invalid_id}\n"
    analysis = _analyze(source)

    assert analysis.has_errors
    [diagnostic] = analysis.diagnostics
    token = f"^{invalid_id}"
    assert diagnostic["code"] == "docwen.markdown.anchor.invalid_id"
    assert diagnostic["range"] == {"start": source.index(token), "end": source.index(token) + len(token)}
    assert diagnostic["source"] == {
        "input_id": "fixture.md",
        "sha256": hashlib.sha256(source.encode()).hexdigest(),
        "encoding": "utf-8",
        "coordinate_system": "unicode_code_point",
        "offset_base": 0,
        "range_end": "exclusive",
    }


def test_duplicate_kind_mismatch_and_nonsemantic_reference_are_fail_closed() -> None:
    source = """# First ^same

Paragraph ^same

Figure: Wrong object ^caption-id

| A |
|---|
| 1 |

![[image.png]] ^raw-image

@[[#^raw-image]]
"""
    analysis = _analyze(source)

    assert analysis.has_errors
    assert [item["code"] for item in analysis.diagnostics] == [
        "docwen.markdown.anchor.duplicate",
        "docwen.markdown.caption.object_mismatch",
        "docwen.markdown.cross_reference.non_semantic_target",
    ]
    duplicate = analysis.diagnostics[0]
    assert duplicate["related_ranges"] == [
        {
            "start": source.index("^same"),
            "end": source.index("^same") + len("^same"),
        }
    ]
    assert analysis.projection["references"][0]["resolution_status"] == "non_semantic"


def test_dangling_empty_caption_and_empty_equation_diagnostics_are_exact() -> None:
    source = """^dangling

Figure:

![image](image.png)

Equation:

$$x=1$$
"""
    equation_start = source.index("Equation")
    analysis = _analyze(source, semantic_id_replacements={equation_start: "energy-id"})

    assert [item["code"] for item in analysis.diagnostics] == [
        "docwen.markdown.anchor.dangling",
        "docwen.markdown.caption.content_required",
        "docwen.markdown.caption.empty_equation_target_required",
    ]
    equation = analysis.diagnostics[-1]
    assert equation["fixes"][0]["fix_id"] == "docwen.markdown.fix.add_semantic_id"
    assert equation["fixes"][0]["edits"][0]["range"]["start"] == equation["fixes"][0]["edits"][0]["range"]["end"]
    assert equation["fixes"][0]["edits"][0]["replacement"] == " ^energy-id"


def test_closing_fence_and_closing_math_suffixes_are_invalid_not_post_block_anchors() -> None:
    source = """```text
body
``` ^bad-code

$$
x=1
$$ ^bad-equation
"""
    analysis = _analyze(source)

    assert analysis.projection["anchors"] == []
    assert analysis.projection["fenced_sources"] == []
    assert [(item["code"], source[item["range"]["start"] : item["range"]["end"]]) for item in analysis.diagnostics] == [
        ("docwen.markdown.anchor.invalid_id", "^bad-code"),
        ("docwen.markdown.anchor.invalid_id", "^bad-equation"),
    ]


@pytest.mark.parametrize(
    "source",
    [
        "| A |\n|---|\n| 1 | ^bad-table\n",
        "> first\n> second ^bad-quote\n",
        "> [!NOTE]\n> body ^bad-callout\n",
    ],
)
def test_structured_inline_anchor_diagnostic_range_selects_authored_token(source: str) -> None:
    analysis = _analyze(source)

    [diagnostic] = analysis.diagnostics
    selected = source[diagnostic["range"]["start"] : diagnostic["range"]["end"]]
    assert diagnostic["code"] == "docwen.markdown.anchor.invalid_id"
    assert selected.startswith("^bad-")


def test_post_block_anchor_depth_mismatch_is_dangling() -> None:
    source = "| A |\n|---|\n| 1 |\n  ^wrong-depth\n"
    analysis = _analyze(source)

    assert analysis.projection["anchors"] == []
    assert [item["code"] for item in analysis.diagnostics] == ["docwen.markdown.anchor.dangling"]


def test_post_block_anchor_accepts_multiple_blank_lines_and_eof_without_newline() -> None:
    source = "| A |\n|---|\n| 1 |\n\n\n^table-at-eof"
    analysis = _analyze(source)

    assert not analysis.has_errors
    assert [(item["id"], item["block_kind"]) for item in analysis.projection["anchors"]] == [("table-at-eof", "table")]


def test_intervening_content_prevents_searching_back_for_an_anchor_owner() -> None:
    source = "| A |\n|---|\n| 1 |\n\nIntervening paragraph\n\n^not-table\n"
    analysis = _analyze(source)

    assert analysis.projection["anchors"] == []
    assert [item["code"] for item in analysis.diagnostics] == ["docwen.markdown.anchor.dangling"]


def test_tokens_in_code_urls_html_and_email_are_not_semantics() -> None:
    source = """`@inline-code` and person@example.com and https://example.test/@url-key

<span data-key="@html-key">text</span>

```text
@code-key @[[#^code-target]]
```

Visible @citation-key.
"""
    analysis = _analyze(source)

    assert [item["raw"] for item in analysis.projection["citations"]] == ["@citation-key"]
    assert analysis.projection["references"] == []


def test_normal_parser_does_not_recognize_legacy_grammar() -> None:
    source = """Figure: Legacy {#fig-legacy}

![image](image.png)

Listing: Legacy listing {#lst-legacy}

```text
body
```

List: Never reserved {#list-legacy}

@fig-legacy
"""
    analysis = _analyze(source)

    assert analysis.projection["targets"] == []
    assert analysis.projection["references"] == []
    assert [item["items"][0]["key"] for item in analysis.projection["citations"]] == ["fig-legacy"]


def test_writer_helpers_preserve_token_ownership_and_choose_safe_fences() -> None:
    assert render_cross_reference_token(selector_kind="stable_id", target_id="x-1") == "@[[#^x-1]]"
    assert (
        render_cross_reference_token(
            selector_kind="heading_path",
            page_locator="Paper",
            heading_path=("Methods", "Sample"),
            alias="Current sample",
        )
        == "@[[Paper#Methods#Sample|Current sample]]"
    )
    assert render_citation_token(("fig-legacy",), parenthetical=False) == "@fig-legacy"
    assert render_citation_token(("smith", "wang"), parenthetical=True) == "[@smith; @wang]"
    assert render_caption_declaration("code_block", "Example", target_id="any-name") == "Code: Example ^any-name"
    assert render_caption_declaration("equation", "", target_id="energy") == "Equation: ^energy"
    assert select_safe_fence("```\n~~~~") == "````"
    assert select_safe_fence("````\n~~~") == "~~~~"
    assert select_safe_fence("literal ````` inline") == "```"
