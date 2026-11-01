"""Source ownership must agree across navigation and link policy consumers."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import (
    extract_block_by_id,
    extract_section_by_heading,
    markdown_heading_ranges,
    markdown_source_projection,
    process_markdown_links,
)

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "literal",
    [
        "---\ntitle: |\n  # Target\n  EXAMPLE_ONLY ^target\n---\n",
        "<!--\n# Target\nEXAMPLE_ONLY ^target\n-->\n",
        "<pre>\n# Target\nEXAMPLE_ONLY ^target\n</pre>\n",
        "%%\n# Target\nEXAMPLE_ONLY ^target\n%%\n",
        "````markdown\n# Target\n```\nEXAMPLE_ONLY ^target\n````\n",
        "    # Target\n    EXAMPLE_ONLY ^target\n",
        "> ```markdown\n> # Target\n> EXAMPLE_ONLY ^target\n> ```\n",
        "- ```markdown\n  # Target\n  EXAMPLE_ONLY ^target\n  ```\n",
        "`` EXAMPLE_ONLY ^target\n``\n",
        "%%\n```markdown\n# Target\nEXAMPLE_ONLY ^target\n%%\n",
        "<!--\n--># Target\n\n",
        "<!--\n-->EXAMPLE_ONLY ^target\n\n",
    ],
)
def test_literal_anchors_never_override_a_real_target(literal: str) -> None:
    source = literal + "\n# Target\n\nREAL_PARAGRAPH ^target\n\n# End\nOTHER\n"
    assert extract_section_by_heading(source, "Target") == "# Target\n\nREAL_PARAGRAPH ^target"
    assert extract_block_by_id(source, "target") == "REAL_PARAGRAPH"


def test_literal_next_heading_does_not_truncate_an_embedded_section() -> None:
    expected = "# Target\n\nBEFORE\n\n````md\n# Example\n```\n# Still example\n````\n\nAFTER"
    assert extract_section_by_heading(expected + "\n\n# End\nOTHER", "Target") == expected


def test_duplicate_heading_titles_keep_the_first_real_heading_contract() -> None:
    assert extract_section_by_heading("# Target\nONE\n\n# Target\nTWO", "Target") == "# Target\nONE"


def test_duplicate_block_ids_do_not_select_an_ambiguous_owner() -> None:
    assert extract_block_by_id("ONE ^target\n\nTWO ^target", "target") is None


@pytest.mark.parametrize("marker", ["bad_id", "bad.id", "bad/id", "x" * 129])
def test_invalid_block_id_does_not_gain_navigation_authority(marker: str) -> None:
    assert extract_block_by_id(f"body ^{marker}", marker) is None


@pytest.mark.parametrize(
    "source, expected",
    [
        ("Target\n======\n\nBODY\n\nNext\n====\nOTHER", "Target\n======\n\nBODY"),
        ("## **Target** ##\n\nBODY\n\n## Next\nOTHER", "## **Target** ##\n\nBODY"),
    ],
)
def test_heading_lookup_uses_supported_heading_syntax(source: str, expected: str) -> None:
    assert extract_section_by_heading(source, "Target") == expected


@pytest.mark.parametrize(
    "prefix",
    [
        "- ```text\n  literal\n\n",
        "- parent\n  - ```text\n    literal\n\n",
        "> - ```text\n>   literal\n\n",
        "1. ````text\n   ```\n   literal\n\n",
        "- ```text\n  literal\n  ```\n\n",
        "- parent\n  ```text\n  literal\n\n",
        "- > ```text\n  > literal\n\n",
        "- - ```text\n    literal\n\n",
    ],
)
@pytest.mark.parametrize("kind", ["markdown", "wiki"])
@pytest.mark.parametrize("mode", ["remove", "extract_text", "keep"])
def test_list_fence_ownership_ends_before_the_next_top_level_link(
    prefix: str, kind: str, mode: str, tmp_path: Path
) -> None:
    link = "[Shown](https://example.test/out)" if kind == "markdown" else "[[Other|Shown]]"
    expected = {"remove": "", "extract_text": "Shown", "keep": link}[mode]
    policy = replace(LinkRuntimeConfig(), **{f"non_embed_{kind}_mode": mode})
    actual = process_markdown_links(prefix + link + "\n", str(tmp_path / "source.md"), link_config=policy)
    assert actual == prefix + expected + "\n"


def test_sibling_list_item_does_not_extend_the_previous_unclosed_fence(tmp_path: Path) -> None:
    source = "- ```text\n  [literal](https://example.test/code)\n- [Shown](https://example.test/out)\n"
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
    )
    assert result == "- ```text\n  [literal](https://example.test/code)\n- \n"


def test_full_source_projection_preserves_unicode_coordinates_and_eols() -> None:
    source = "\ufeff---\r\ntitle: ignored\r\n---\r\n\n``１２３\n😀``\r\n\r\nVisible 己\n"
    projection = markdown_source_projection(source)
    assert len(projection) == len(source)
    assert [(index, char) for index, char in enumerate(projection) if char in "\r\n"] == [
        (index, char) for index, char in enumerate(source) if char in "\r\n"
    ]
    assert projection.index("Visible 己") == source.index("Visible 己")
    assert "ignored" not in projection and "１２３" not in projection and "😀" not in projection


def test_unmatched_inline_ticks_do_not_cross_a_new_heading_block() -> None:
    source = "Example ``\n# Target\nREAL `` text\n\n# Next\nOTHER"
    assert extract_section_by_heading(source, "Target") == "# Target\nREAL `` text"


def test_a_comment_closer_cannot_manufacture_a_same_line_fence() -> None:
    source = "%%\n```example\n%%``` remains prose\n# Target\nREAL\n"
    assert extract_section_by_heading(source, "Target") == "# Target\nREAL"


def test_heading_ranges_respect_the_extended_heading_switch_and_raw_coordinates() -> None:
    source = "\ufeff#tag\r\n\r\n# One ###\r\n\r\n####### Seven\n######## Eight\n######### Nine\n########## Text\n"
    default = markdown_heading_ranges(source, extended_headings=False)
    enabled = markdown_heading_ranges(source, extended_headings=True)
    assert [(item.level, source[item.text_start : item.text_end]) for item in default] == [(1, "One")]
    assert [(item.level, source[item.text_start : item.text_end]) for item in enabled] == [
        (1, "One"),
        (7, "Seven"),
        (8, "Eight"),
        (9, "Nine"),
    ]
    assert default[0].text_start == source.index("One")
    assert default[0].text_end == source.index("One") + 3


def test_a_standalone_anchor_can_own_a_complete_fenced_block_with_internal_blank_lines() -> None:
    code = "````text\nEXAMPLE ^fake\n\n```\nSTILL_CODE\n````"
    source = code + "\n\n^target\n\nOTHER"
    assert extract_block_by_id(source, "target") == code
    assert extract_block_by_id(source, "fake") is None


@pytest.mark.parametrize(
    "block",
    [
        "> quoted\n> continued",
        "> [!note]\n> callout\n>\n> continuation",
        "- first\n- second",
        "1. first\n\n2. second\n   continuation",
        "- parent\n  - child\n- sibling",
        "| Head | Value |\n| --- | --- |\n| one | two |",
    ],
)
def test_standalone_anchor_preserves_a_supported_complete_block(block: str) -> None:
    source = "PROLOGUE\n\n" + block + "\n\n^outer\n\nEPILOGUE"
    assert extract_block_by_id(source, "outer") == block


@pytest.mark.parametrize("boundary", ["---", "| Head |\n| --- |\n| Cell |", "> quote", "- item"])
def test_paragraph_anchor_does_not_capture_an_adjacent_structural_block(boundary: str) -> None:
    source = "BEFORE ^target\n" + boundary + "\n\nAFTER"
    if boundary == "---":
        # This is an actual Setext heading, where an inline ordinary anchor is
        # not a valid simple-paragraph owner.
        assert extract_block_by_id(source, "target") is None
    else:
        assert extract_block_by_id(source, "target") == "BEFORE"


def test_inline_list_item_anchor_keeps_its_descendants_but_not_a_sibling() -> None:
    source = "- BEFORE\n- TARGET ^target\n  - CHILD\n- AFTER"
    assert extract_block_by_id(source, "target") == "- TARGET\n  - CHILD"


@pytest.mark.parametrize("middle", ["> Continued", "Continued"])
def test_setext_heading_lookup_keeps_container_prefixes_out_of_the_title(middle: str) -> None:
    source = f"> Title\n{middle}\n> =====\n> BODY\n\n# Next\nOTHER"
    assert extract_section_by_heading(source, "Title Continued") == f"> Title\n{middle}\n> =====\n> BODY"


def test_inline_code_heading_text_is_not_reinterpreted_as_emphasis() -> None:
    source = "# `*literal*`\nCODE\n\n# *literal*\nEMPHASIS\n"
    assert extract_section_by_heading(source, "*literal*") == "# `*literal*`\nCODE"
    assert extract_section_by_heading(source, "literal") == "# *literal*\nEMPHASIS"


def test_cr_only_fences_end_before_a_following_heading_and_link(tmp_path: Path) -> None:
    source = "```\rcode\r```\r\r# Real\r[Visible](https://example.test)\r"
    assert extract_section_by_heading(source, "Real") == "# Real\r[Visible](https://example.test)"
    assert (
        process_markdown_links(
            source,
            str(tmp_path / "source.md"),
            link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
        )
        == "```\rcode\r```\r\r# Real\r\r"
    )


def test_heading_after_an_initial_list_fence_has_a_standalone_extraction() -> None:
    source = "10. ````text\n    # EXAMPLE\n    ````\n    # Real\n    BODY\n\n# Next\nOTHER"
    assert extract_section_by_heading(source, "Real") == "# Real\nBODY"


def test_bom_heading_has_original_source_coordinates() -> None:
    source = "\ufeff   # Real\r\n"
    [heading] = markdown_heading_ranges(source)
    assert (heading.text_start, heading.text_end) == (source.index("Real"), source.index("Real") + 4)


@pytest.mark.parametrize(
    "source",
    [
        "[Example ^target](https://example.test)\n",
        "![Example ^target](image.png)\n",
        "Example ^target `literal`\n",
        "Example ^target <!-- comment -->\n",
    ],
)
def test_projection_cannot_manufacture_a_source_block_suffix(source: str) -> None:
    assert extract_block_by_id(source, "target") is None


def test_a_visible_link_label_does_not_make_a_real_block_id_ambiguous_or_truncate_its_paragraph() -> None:
    source = "[Example ^target](https://example.test)\nREAL ^target\n"
    assert extract_block_by_id(source, "target") == "[Example ^target](https://example.test)\nREAL"


@pytest.mark.parametrize("level", [6, 7, 8, 9, 10])
def test_default_heading_dialect_does_not_expose_a_literal_code_span(level: int, tmp_path: Path) -> None:
    source = f"`start\r\n{'#' * level} １２３\r\nend`\r\n"
    expected = source if level == 6 else "".join(char if char in "\r\n" else " " for char in source)
    assert markdown_source_projection(source) == expected
    link = "[Visible](https://example.test)"
    linked_source = source.replace("１２３", "１２３ " + link)
    assert process_markdown_links(
        linked_source,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
    ) == (linked_source.replace(link, "") if level == 6 else linked_source)


@pytest.mark.parametrize("extended_headings", [False, True])
@pytest.mark.parametrize("level", [6, 7, 8, 9, 10])
def test_explicit_heading_dialect_controls_block_and_inline_boundaries(
    extended_headings: bool, level: int, tmp_path: Path
) -> None:
    active_heading = level == 6 or (extended_headings and level <= 9)
    source = f"`start\r\n{'#' * level} １２３\r\nend`\r\n"
    expected = source if active_heading else "".join(char if char in "\r\n" else " " for char in source)
    assert markdown_source_projection(source, extended_headings=extended_headings) == expected
    link = "[Visible](https://example.test)"
    linked_source = source.replace("１２３", "１２３ " + link)
    assert process_markdown_links(
        linked_source,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
        extended_headings=extended_headings,
    ) == (linked_source.replace(link, "") if active_heading else linked_source)

    following_indent = f"{'#' * level} Title\r\n    １２３\r\n"
    assert markdown_source_projection(following_indent, extended_headings=extended_headings) == (
        following_indent.replace("    １２３", "       ") if active_heading else following_indent
    )


@pytest.mark.parametrize("extended_headings", [False, True])
@pytest.mark.parametrize("level", [6, 7, 8, 9, 10])
def test_precision_embedding_uses_the_frozen_heading_dialect(
    extended_headings: bool, level: int, tmp_path: Path
) -> None:
    source = f"`start\n{'#' * level} Target\nBODY\nend`\n"
    (tmp_path / "child.md").write_text(source, encoding="utf-8")
    reference = "![[child.md#Target]]"
    actual = process_markdown_links(
        reference,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), file_not_found_mode="keep"),
        extended_headings=extended_headings,
    )
    expected = f"{'#' * level} Target\nBODY\nend`" if level == 6 or (extended_headings and level <= 9) else reference
    assert actual == expected


def test_direct_heading_lookup_retains_its_h9_default() -> None:
    source = "######### Target\nBODY\n"
    assert extract_section_by_heading(source, "Target") == source.rstrip("\n")
    assert extract_section_by_heading(source, "Target", extended_headings=False) is None


@pytest.mark.parametrize("extended_headings", [False, True])
@pytest.mark.parametrize("level", [7, 8, 9, 10])
def test_disabled_extended_atx_can_still_be_a_plain_setext_title(extended_headings: bool, level: int) -> None:
    title = f"{'#' * level} Title"
    source = title + "\n-----\nBODY\n"
    [heading] = markdown_heading_ranges(source, extended_headings=extended_headings)
    if extended_headings and level <= 9:
        assert (heading.level, heading.style, source[heading.text_start : heading.text_end]) == (level, "atx", "Title")
    else:
        assert (heading.level, heading.style, source[heading.text_start : heading.text_end]) == (2, "setext", title)
