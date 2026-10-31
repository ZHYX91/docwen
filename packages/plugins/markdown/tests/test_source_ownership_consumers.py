"""Numbering and embedding must preserve source ownership in real artifacts."""

from __future__ import annotations

from pathlib import Path

import pytest
from tests.support.numbering import repository_numbering_registry

from docwen_core.text.heading_numbering import compile_clean_rules_from_data
from docwen_plugin_markdown.numbering.converter import MdNumberingProcessor

from ._link_processing_routes_support import _convert_docx, _link_config
from .conftest import make_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("operation", ["add", "remove"])
def test_numbering_changes_only_the_visible_heading_source_range(operation: str, tmp_path: Path) -> None:
    prefix = (
        "\ufeff---\r\ndescription: |\r\n  # 1. metadata literal\r\n---\r\n\r\n"
        "<!--\r\n# 1. hidden literal\r\n-->\r\n\r\n"
        "<pre>\n# 1. raw HTML\n</pre>\n\n"
        "%%\n# 1. Obsidian comment\n%%\n\n"
        "- ````markdown\n  # 1. fenced literal\n  ```\n  # 2. still literal\n  ````\n\n"
        "Ordinary body\n\n    # 1. indented literal\n\n#tag\n\n########## Not a heading\n\n"
    )
    title = "Visible heading" if operation == "add" else "1. Visible heading"
    original = prefix + "  #\t" + title + "  ### \r\n"
    source = tmp_path / "number.md"
    source.write_bytes(original.encode("utf-8"))
    context, _workspace = make_context(
        str(source),
        target_format="md",
        action_name="process_md_numbering",
        options={
            "add_numbering": operation == "add",
            "remove_numbering": operation == "remove",
            "numbering_scheme": "hierarchical_standard",
        },
        numbering_registry=repository_numbering_registry(),
        heading_cleanup_rules=compile_clean_rules_from_data(
            [{"id": "decimal", "enabled": True, "pattern": r"^\d+\.\s*", "level": 1}]
        ),
    )
    result = MdNumberingProcessor().convert(context)
    assert result.success, result.error
    expected_title = "1 Visible heading" if operation == "add" else "Visible heading"
    assert Path(result.artifacts[0].staging_path).read_bytes() == (
        prefix + "  #\t" + expected_title + "  ### \r\n"
    ).encode("utf-8")
    assert source.read_bytes() == original.encode("utf-8")


@pytest.mark.parametrize("anchor", ["Target", "^target"])
def test_embedding_publishes_the_real_target_once_not_a_code_example(anchor: str, tmp_path: Path) -> None:
    child = tmp_path / "child.md"
    child.write_text(
        "````markdown\n# Target\nEXAMPLE_ONLY ^target\n````\n\n# Target\n\nREAL_PARAGRAPH ^target\n\n# End\nOTHER\n",
        encoding="utf-8",
    )
    source = tmp_path / "source.md"
    source.write_text(f"![[child.md#{anchor}]]\n", encoding="utf-8")
    result = _convert_docx(source, _link_config())
    assert result.text.count("REAL_PARAGRAPH") == 1
    assert "EXAMPLE_ONLY" not in result.text
    assert "OTHER" not in result.text


@pytest.mark.parametrize("anchor", ["Target", "^target"])
def test_missing_real_target_keeps_the_requested_reference(anchor: str, tmp_path: Path) -> None:
    (tmp_path / "child.md").write_text("```md\n# Target\nEXAMPLE_ONLY ^target\n```\n", encoding="utf-8")
    source = tmp_path / "source.md"
    reference = f"![[child.md#{anchor}]]"
    source.write_text(reference + "\n", encoding="utf-8")
    result = _convert_docx(source, _link_config(file_not_found_mode="keep"))
    assert reference in result.text
    assert "EXAMPLE_ONLY" not in result.text


@pytest.mark.parametrize(
    "mode, expected_text, expected_links", [("remove", "", 0), ("extract_text", "Shown", 0), ("hyperlink", "Shown", 1)]
)
def test_link_policy_after_unclosed_list_fence_reaches_the_saved_docx(
    mode: str, expected_text: str, expected_links: int, tmp_path: Path
) -> None:
    source = tmp_path / "links.md"
    source.write_text("- ```text\n  literal\n\n[Shown](https://example.test/out)\n", encoding="utf-8")
    result = _convert_docx(source, _link_config(markdown_mode=mode))
    assert "literal" in result.text
    assert result.text.count("Shown") == int(bool(expected_text))
    assert len(result.hyperlink_targets) == expected_links


@pytest.mark.parametrize("extended", [False, True])
def test_numbering_extension_policy_is_applied_to_the_authored_source(extended: bool, tmp_path: Path) -> None:
    source = tmp_path / "extended.md"
    source.write_text(
        "# 1. First\n####### 7. Seven\n######## 8. Eight\n######### 9. Nine\n########## 10. Text\n", encoding="utf-8"
    )
    context, _workspace = make_context(
        str(source),
        target_format="md",
        action_name="process_md_numbering",
        options={"remove_numbering": True, "add_numbering": False},
        config_values={"conversion": {"markdown_extensions": {"input": {"extended_headings": extended}}}},
        heading_cleanup_rules=compile_clean_rules_from_data(
            [{"id": "decimal", "enabled": True, "pattern": r"^\d+\.\s*", "level": 1}]
        ),
    )
    result = MdNumberingProcessor().convert(context)
    assert result.success, result.error
    expected = (
        "# First\n"
        + (
            "####### Seven\n######## Eight\n######### Nine\n"
            if extended
            else "####### 7. Seven\n######## 8. Eight\n######### 9. Nine\n"
        )
        + "########## 10. Text\n"
    )
    assert Path(result.artifacts[0].staging_path).read_text(encoding="utf-8") == expected


@pytest.mark.parametrize(
    "title", ["`Literal`", "``literal ` and #``", "$x + y$", "![Alt](image.png)", "[Title](https://example.test)"]
)
def test_numbering_uses_original_separator_when_the_heading_is_an_inline_literal(title: str, tmp_path: Path) -> None:
    source = tmp_path / "literal-heading.md"
    source.write_bytes(("#\t" + title + " ###\r\n").encode("utf-8"))
    context, _workspace = make_context(
        str(source),
        target_format="md",
        action_name="process_md_numbering",
        options={"remove_numbering": False, "add_numbering": True, "numbering_scheme": "hierarchical_standard"},
        numbering_registry=repository_numbering_registry(),
    )
    result = MdNumberingProcessor().convert(context)
    assert result.success, result.error
    assert Path(result.artifacts[0].staging_path).read_bytes() == ("#\t1 " + title + " ###\r\n").encode("utf-8")


@pytest.mark.parametrize(
    "markdown",
    [
        "`\n# Heading\n[Visible](https://example.test)\n`\n",
        "| A | B |\n| --- | --- |\n| ` | [Visible](https://example.test) ` |\n",
    ],
)
def test_link_removal_never_pairs_inline_code_across_a_block_or_table_cell(markdown: str, tmp_path: Path) -> None:
    source = tmp_path / "scopes.md"
    source.write_text(markdown, encoding="utf-8")
    result = _convert_docx(source, _link_config(markdown_mode="remove"))
    assert result.hyperlink_targets == frozenset()
    assert "Visible" not in result.text
    assert "`" in result.text


def test_numbering_keeps_list_context_when_the_first_child_is_fenced_code(tmp_path: Path) -> None:
    prefix = "10. ````text\n    # 9. EXAMPLE\n    ````\n"
    source = tmp_path / "list-heading.md"
    source.write_bytes((prefix + "    # Real\r\n").encode("utf-8"))
    context, _workspace = make_context(
        str(source),
        target_format="md",
        action_name="process_md_numbering",
        options={"remove_numbering": False, "add_numbering": True, "numbering_scheme": "hierarchical_standard"},
        numbering_registry=repository_numbering_registry(),
    )
    result = MdNumberingProcessor().convert(context)
    assert result.success, result.error
    assert Path(result.artifacts[0].staging_path).read_bytes() == (prefix + "    # 1 Real\r\n").encode("utf-8")


@pytest.mark.parametrize(
    "fake",
    [
        "[EXAMPLE_ONLY ^target](https://example.test)",
        "![EXAMPLE_ONLY ^target](missing.png)",
        "EXAMPLE_ONLY ^target `literal`",
    ],
)
def test_projection_only_suffixes_do_not_override_a_real_embedded_block(fake: str, tmp_path: Path) -> None:
    (tmp_path / "child.md").write_text(fake + "\n\nREAL_PARAGRAPH ^target\n", encoding="utf-8")
    source = tmp_path / "source.md"
    source.write_text("![[child.md#^target]]\n", encoding="utf-8")
    result = _convert_docx(source, _link_config())
    assert result.text.count("REAL_PARAGRAPH") == 1
    assert "EXAMPLE_ONLY" not in result.text


@pytest.mark.parametrize("extended_headings", [None, False, True])
@pytest.mark.parametrize("level", [6, 7, 8, 9, 10])
def test_source_numbering_uses_the_frozen_heading_dialect_inside_literal_candidates(
    extended_headings: bool | None, level: int, tmp_path: Path
) -> None:
    original = f"`start\r\n{'#' * level} 1. Title\r\nend`\r\n"
    raw = original.encode("utf-8")
    source = tmp_path / "numbering-heading-dialect.md"
    source.write_bytes(raw)
    config_values = (
        None
        if extended_headings is None
        else {"conversion": {"markdown_extensions": {"input": {"extended_headings": extended_headings}}}}
    )
    context, _workspace = make_context(
        str(source),
        target_format="md",
        action_name="process_md_numbering",
        options={"remove_numbering": True, "add_numbering": False},
        config_values=config_values,
        heading_cleanup_rules=compile_clean_rules_from_data(
            [{"id": "decimal", "enabled": True, "pattern": r"^\d+\.\s*", "level": 1}]
        ),
    )
    result = MdNumberingProcessor().convert(context)
    assert result.success, result.error
    expected = original.replace("1. ", "") if level == 6 or (extended_headings and level <= 9) else original
    assert Path(result.artifacts[0].staging_path).read_bytes() == expected.encode("utf-8")
    assert source.read_bytes() == raw
