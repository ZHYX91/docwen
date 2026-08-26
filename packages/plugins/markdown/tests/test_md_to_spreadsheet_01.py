"""Focused tests split from test_md_to_spreadsheet.py."""

from __future__ import annotations

from ._md_to_spreadsheet_support import (
    LinkRuntimeConfig,
    MdToCsvConverter,
    MdToXlsxConverter,
    _required_range_boundaries,
    make_context,
    make_table_safe,
    parse_raw_md_tables,
    pytest,
    spreadsheet_converter,
    write_temp_md,
)


@pytest.mark.contract
def test_required_range_boundaries_returns_complete_integer_geometry() -> None:
    assert _required_range_boundaries("B3:D7") == (2, 3, 4, 7)


@pytest.mark.contract
def test_raw_table_parser_requires_matching_strict_fence_closer() -> None:
    content = (
        "````markdown\n"
        "| Hidden | Table |\n"
        "| --- | --- |\n"
        "| one | two |\n"
        "```not-a-closing-fence\n"
        "| Also | Hidden |\n"
        "| --- | --- |\n"
        "| three | four |\n"
        "```\n"
        "````\n"
        "    | Indented | Code |\n"
        "    | --- | --- |\n"
        "    | seven | eight |\n"
        "\t| Tab | Code |\n"
        "\t| --- | --- |\n"
        "\t| nine | ten |\n"
        "> ~~~\n"
        "> | Quoted | Code |\n"
        "> | --- | --- |\n"
        "> | eleven | twelve |\n"
        "> ~~~\n"
        "| Real | Table |\n"
        "| --- | --- |\n"
        "| five | six |\n"
    )

    assert parse_raw_md_tables(content) == [{"headers": ["Real", "Table"], "rows": [["five", "six"]]}]


@pytest.mark.contract
def test_authenticated_table_break_token_is_repeatably_restored() -> None:
    safe_cell = make_table_safe("\nA\n\nB\n")
    content = f"| Value |\n| --- |\n| {safe_cell} |\n"

    first = parse_raw_md_tables(content)
    second = parse_raw_md_tables(content)

    assert first == [{"headers": ["Value"], "rows": [["A\nB"]]}]
    assert second == first


@pytest.mark.parametrize(
    ("converter_type", "target_format"),
    [(MdToXlsxConverter, "xlsx"), (MdToCsvConverter, "csv")],
)
@pytest.mark.contract
def test_spreadsheet_routes_forward_request_link_policy(
    monkeypatch: pytest.MonkeyPatch,
    converter_type: type,
    target_format: str,
) -> None:
    """Both spreadsheet routes consume the request snapshot, not global state."""
    observed: dict[str, object] = {}

    def _observe_link_processing(text: str, source_file_path: str, **kwargs: object) -> str:
        observed.update(kwargs)
        observed["source_file_path"] = source_file_path
        return text

    monkeypatch.setattr(spreadsheet_converter, "process_markdown_links", _observe_link_processing)
    md_path = write_temp_md("| A | B |\n| --- | --- |\n| value | policy |\n")
    context, workspace = make_context(
        md_path,
        target_format=target_format,
        config_values={
            "link": {
                "non_embed_links": {
                    "wiki_mode": "extract_text",
                    "markdown_mode": "remove",
                    "auto_link_bare_url": True,
                },
                "embed_links": {
                    "wiki_image_mode": "keep",
                    "markdown_image_mode": "extract_text",
                    "md_file_mode": "remove",
                },
                "embedding": {"max_depth": 7},
                "path_resolution": {"search_dirs": ["vault"]},
                "error_handling": {
                    "file_not_found": "keep",
                    "detect_circular": False,
                    "circular_reference": "ignore",
                    "max_depth_reached": "keep",
                },
            }
        },
    )

    result = converter_type().convert(context)

    assert result.success is True, result.error
    assert observed["source_file_path"] == md_path
    assert observed["target_format"] == target_format
    assert observed["table_safe"] is True
    assert observed["temp_dir"] == str(workspace.staging_dir)
    link_config = observed["link_config"]
    assert isinstance(link_config, LinkRuntimeConfig)
    assert link_config.max_depth == 7
    assert link_config.non_embed_wiki_mode == "extract_text"
    assert link_config.non_embed_markdown_mode == "remove"
    assert link_config.auto_link_bare_url is True
    assert link_config.embed_wiki_image_mode == "keep"
    assert link_config.embed_markdown_image_mode == "extract_text"
    assert link_config.embed_md_file_mode == "remove"
    assert link_config.search_dirs == ("vault",)
    assert link_config.detect_circular is False
    assert link_config.file_not_found_mode == "keep"
    assert link_config.circular_reference_mode == "ignore"
    assert link_config.max_depth_reached_mode == "keep"
