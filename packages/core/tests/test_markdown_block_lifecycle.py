"""One paragraph/container lifecycle owns heading and literal source ranges."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import (
    extract_block_by_id,
    extract_section_by_heading,
    markdown_heading_ranges,
    markdown_source_context_lines,
    markdown_source_projection,
    process_markdown_links,
)
from docwen_core.links._source_document import markdown_heading_text

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("opener", ["```text", "<pre>"])
@pytest.mark.parametrize(
    ("intro", "indent"),
    [("- text\nlazy\n", "  "), ("- parent\n  - text\nlazy\n", "    ")],
)
def test_lazy_paragraph_keeps_its_list_owner_until_a_literal_really_exits(
    opener: str, intro: str, indent: str, tmp_path: Path
) -> None:
    prefix = intro + indent + opener + "\n" + indent + "[hidden](https://example.test/in)\n\n"
    source = prefix + "# Real\nREAL ^target\n\n[out](https://example.test) [[Outside]]\n"
    assert extract_section_by_heading(source, "Real") == (
        "# Real\nREAL ^target\n\n[out](https://example.test) [[Outside]]"
    )
    assert extract_block_by_id(source, "target") == "REAL"
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove", non_embed_wiki_mode="remove"),
    )
    assert result == prefix + "# Real\nREAL ^target\n\n \n"


@pytest.mark.parametrize("marker", ["0.", "2.", "9)", "123."])
def test_other_ordered_starts_continue_a_paragraph_and_its_inline_scope(marker: str, tmp_path: Path) -> None:
    source = f"intro\n{marker} # Fake\n\n# Real\nBody\n"
    headings = markdown_heading_ranges(source)
    assert [(h.level, source[h.text_start : h.text_end]) for h in headings] == [(1, "Real")]
    literal = f"start `open\n{marker} [literal](https://example.test)\nclose`\n"
    result = process_markdown_links(
        literal,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
    )
    assert result == literal
    assert "literal" not in markdown_source_projection(literal)


def test_one_ordered_start_and_real_list_siblings_remain_block_boundaries() -> None:
    source = "intro\n1. # One\n2. # Two\n3. # Three\n\n# Real\n"
    assert [(h.level, source[h.text_start : h.text_end]) for h in markdown_heading_ranges(source)] == [
        (1, "One"),
        (1, "Two"),
        (1, "Three"),
        (1, "Real"),
    ]


def test_setext_closes_its_inline_scope_without_splitting_multiline_title(tmp_path: Path) -> None:
    source = "start `open\n-\n[literal](https://example.test) close`\n"
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
    )
    assert result == "start `open\n-\n close`\n"
    projection = markdown_source_projection(source)
    assert "literal" in projection
    complete = "start `open\n[literal](https://example.test) close`\n===\n"
    assert (
        process_markdown_links(
            complete,
            str(tmp_path / "source.md"),
            link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
        )
        == complete
    )
    assert "literal" not in markdown_source_projection(complete)


@pytest.mark.parametrize("rule", ["***", "* * *", "___", "_ _ _", "---"])
def test_thematic_break_is_not_setext_paragraph_content(rule: str) -> None:
    source = f"{rule}\n===\n\n# Real\nBody\n"
    assert [(h.level, source[h.text_start : h.text_end]) for h in markdown_heading_ranges(source)] == [(1, "Real")]


@pytest.mark.parametrize("marker", ["0.", "2.", "3)"])
def test_noninterrupting_ordered_text_can_belong_to_multiline_setext(marker: str) -> None:
    source = f"foo\n{marker} bar\n===\n"
    [heading] = markdown_heading_ranges(source)
    assert (heading.level, heading.style, heading.start, heading.end) == (1, "setext", 0, len(source))
    assert [source[start:end] for start, end in heading.text_ranges] == ["foo", f"{marker} bar"]


@pytest.mark.parametrize("level", range(1, 10))
@pytest.mark.parametrize("prefix", ["", "> ", "- "])
@pytest.mark.parametrize("closing", ["", " ###"])
def test_empty_atx_owns_no_closing_hashes(level: int, prefix: str, closing: str) -> None:
    source = "\ufeff" + prefix + "#" * level + closing + "\r\n"
    [heading] = markdown_heading_ranges(source, extended_headings=True)
    assert (heading.level, heading.style) == (level, "atx")
    assert source[heading.text_start : heading.text_end] == ""
    assert [source[start:end] for start, end in heading.text_ranges] == [""]


@pytest.mark.parametrize("autolink", ["https://example.test", "person@example.test"])
def test_navigation_keeps_autolink_visible_text_and_selects_the_actual_title(autolink: str) -> None:
    source = f"# A <{autolink}>\nFIRST\n\n# A\nSECOND\n"
    assert extract_section_by_heading(source, "A") == "# A\nSECOND"
    assert extract_section_by_heading(source, f"A {autolink}") == f"# A <{autolink}>\nFIRST"


@pytest.mark.parametrize(
    ("source", "visible"),
    [
        (r"A \<https://example.test>", "A <https://example.test>"),
        ("A `<https://example.test>`", "A <https://example.test>"),
        ("A <em>Label</em>", "A Label"),
        ("A <not-a-valid tag=> tail", "A <not-a-valid tag=> tail"),
        ("A <br> B", "A   B"),
    ],
)
def test_heading_angle_text_follows_inline_atom_ownership(source: str, visible: str) -> None:
    assert markdown_heading_text(source) == visible


@pytest.mark.parametrize("opener", ["```text", "<pre>"])
@pytest.mark.parametrize("separated", [False, True])
def test_list_literal_dedent_requires_an_actual_item_exit(opener: str, separated: bool, tmp_path: Path) -> None:
    prefix = "- " + opener + "\n  literal\n" + ("\n" if separated else "")
    outgoing = "[visible](https://example.test)"
    source = prefix + outgoing + "\n"
    result = process_markdown_links(
        source,
        str(tmp_path / "source.md"),
        link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
    )
    assert result == prefix + ("" if separated else outgoing) + "\n"
    assert ("visible" in markdown_source_projection(source)) is separated
    rows = markdown_source_context_lines(source)
    assert rows[1].block_kind == "literal"
    assert rows[-1].block_kind == ("paragraph" if separated else "literal")
    assert not rows[1].paragraph_continuation


@pytest.mark.parametrize("separator", ["-", "==="])
def test_setext_eligibility_distinguishes_unmarked_quote_and_list_lines(separator: str) -> None:
    quote = "> title\n" + separator + "\n"
    item = "- title\n" + separator + "\n"
    assert markdown_heading_ranges(quote) == []
    headings = markdown_heading_ranges(item)
    if separator == "-":
        # A new item at the outer indentation ends the old item first.
        assert headings == []
    else:
        assert [(h.style, h.level, item[h.text_start : h.text_end]) for h in headings] == [("setext", 1, "title")]


@pytest.mark.parametrize("level", range(1, 7))
def test_atx_after_unmarked_quote_line_retains_one_container_identity(level: int) -> None:
    source = "> intro\n" + "#" * level + " title\ncontinued\n"
    rows = markdown_source_context_lines(source)
    assert rows[0].containers == rows[1].containers == rows[2].containers
    assert [row.block_kind for row in rows] == ["paragraph", "atx_heading", "paragraph"]


@pytest.mark.parametrize(
    "source, visible",
    [
        ("A <https://example.test/?x=&amp;y>", "A https://example.test/?x=&amp;y"),
        (r"A <https://example.test/\*>", r"A https://example.test/\*"),
        ("A <br class='break'>B", "A  B"),
    ],
)
def test_autolink_text_is_literal_but_real_html_still_has_visible_semantics(source: str, visible: str) -> None:
    assert markdown_heading_text(source) == visible


@pytest.mark.parametrize(
    "source, kinds",
    [
        ("$$x + y$$\n", ["math_block"]),
        ("\ufeff$$\r\nx + y\r\n$$\r\n", ["math_block", "literal", "literal"]),
        ("intro\n$$\nx + y\n$$\n", ["paragraph", "math_block", "literal", "literal"]),
        ("    $$x + y$$\n", ["indented_code"]),
        ("```text\n$$x + y$$\n```\n", ["fenced_code", "literal", "literal"]),
        ("<pre>\n$$x + y$$\n</pre>\n", ["html_block", "literal", "literal"]),
        ("$$\nx + y\n", ["paragraph", "paragraph"]),
    ],
)
def test_context_exposes_only_inventory_proven_math_openers(source: str, kinds: list[str]) -> None:
    rows = markdown_source_context_lines(source)
    assert [row.block_kind for row in rows] == kinds
    assert all(not row.paragraph_continuation for row in rows if row.block_kind == "math_block")


@pytest.mark.parametrize("prefix, continuation", [("> ", "> "), ("- ", "  "), ("> - ", ">   "), ("- > ", "  > ")])
@pytest.mark.parametrize("multiline", [False, True])
def test_container_math_keeps_its_owner_and_hides_formula_source(
    prefix: str, continuation: str, multiline: bool, tmp_path: Path
) -> None:
    formula = "[hidden](https://example.test/math) ^fake"
    body = prefix + (
        "$$\r\n" + continuation + formula + "\r\n" + continuation + "$$" if multiline else "$$" + formula + "$$"
    )
    source = "\ufeff" + body + "\r\n\r\n# Real\r\n[visible](https://example.test/out) ^real\r\n"
    rows = markdown_source_context_lines(source)
    assert rows[0].block_kind == "math_block"
    if multiline:
        assert [row.block_kind for row in rows[1:3]] == ["literal", "literal"]
        assert rows[0].containers == rows[1].containers == rows[2].containers
    projection = markdown_source_projection(source)
    assert len(projection) == len(source)
    assert [(index, char) for index, char in enumerate(projection) if char in "\r\n"] == [
        (index, char) for index, char in enumerate(source) if char in "\r\n"
    ]
    assert "fake" not in projection and "hidden" not in projection
    assert extract_block_by_id(source, "fake") is None
    assert "visible" in projection and "real" in projection
    assert (
        process_markdown_links(
            source,
            str(tmp_path / "source.md"),
            link_config=replace(LinkRuntimeConfig(), non_embed_markdown_mode="remove"),
        )
        == "\ufeff" + body + "\r\n\r\n# Real\r\n ^real\r\n"
    )


def test_math_closer_outside_an_exited_list_does_not_authenticate_the_opener() -> None:
    source = "- $$\n  formula\n\n# Real\n$$\n"
    assert not any(row.block_kind == "math_block" for row in markdown_source_context_lines(source))
    assert extract_section_by_heading(source, "Real") == "# Real\n$$"
