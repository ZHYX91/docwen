"""Focused tests split from test_link_processing_routes.py."""

from __future__ import annotations

from ._link_processing_routes_support import (
    _TINY_PNG,
    MdToCsvConverter,
    MdToXlsxConverter,
    Path,
    Workbook,
    _convert_docx,
    _link_config,
    csv,
    load_workbook,
    make_context,
    pytest,
)

pytestmark = pytest.mark.contract


def test_docx_strict_and_indented_code_never_materialize_links_or_images(
    tmp_path: Path,
) -> None:
    image = tmp_path / "secret.png"
    image.write_bytes(_TINY_PNG)
    source = tmp_path / "strict-code.md"
    source.write_text(
        "````markdown\n"
        "[inside](https://inside.example)\n"
        "```not-a-closing-fence\n"
        f"{{{{IMAGE:{image.as_posix()}}}}}\n"
        "https://bare-inside.example\n"
        "````\n\n"
        "    [indented](https://indented.example) ![image](secret.png)\n",
        encoding="utf-8",
    )

    observation = _convert_docx(
        source,
        _link_config(
            markdown_mode="remove",
            markdown_image_mode="remove",
            auto_link_bare_url=True,
        ),
    )

    assert not observation.hyperlink_targets
    assert not observation.media_names
    assert "inside.example" in observation.text
    assert "indented.example" in observation.text


def test_docx_remove_policy_preserves_literal_unscoped_image_marker(
    tmp_path: Path,
) -> None:
    image = tmp_path / "literal.png"
    image.write_bytes(_TINY_PNG)
    marker = f"{{{{IMAGE:{image.as_posix()}}}}}"
    source = tmp_path / "literal-marker-remove.md"
    source.write_text(
        f"Literal {marker}\n\n![removed](literal.png)",
        encoding="utf-8",
    )

    observation = _convert_docx(
        source,
        _link_config(markdown_image_mode="remove"),
    )

    assert marker in observation.text
    assert "removed" not in observation.text
    assert not observation.media_names


def test_docx_nested_literal_marker_stays_text_while_generated_image_embeds(
    tmp_path: Path,
) -> None:
    image = tmp_path / "nested.png"
    image.write_bytes(_TINY_PNG)
    marker = f"{{{{IMAGE:{image.as_posix()}}}}}"
    child = tmp_path / "child.md"
    child.write_text(
        f"Nested literal {marker}\n\n![[nested.png]]",
        encoding="utf-8",
    )
    source = tmp_path / "nested-marker.md"
    source.write_text("![[child.md]]", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert marker in observation.text
    assert len(observation.media_names) == 1
    assert "<w:drawing" in observation.document_xml


def test_docx_nested_markdown_keeps_child_source_context_for_assets_and_links(
    tmp_path: Path,
) -> None:
    nested = tmp_path / "nested"
    nested.mkdir()
    image = nested / "pixel.png"
    image.write_bytes(_TINY_PNG)
    wiki_target = nested / "wiki-target.md"
    wiki_target.write_text("# Wiki target\n", encoding="utf-8")
    markdown_target = nested / "markdown(target).md"
    markdown_target.write_text("# Markdown target\n", encoding="utf-8")
    child = nested / "child.md"
    child.write_text(
        "![child image](pixel.png)\n\n[[wiki-target.md|Wiki target]]\n\n[Markdown target](markdown\\(target\\).md)\n",
        encoding="utf-8",
    )
    source = tmp_path / "parent.md"
    source.write_text("![[nested/child.md]]", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert len(observation.media_names) == 1
    assert "File not found" not in observation.text
    assert "Wiki target" in observation.text
    assert "Markdown target" in observation.text
    assert observation.hyperlink_targets == {
        wiki_target.resolve().as_uri(),
        markdown_target.resolve().as_uri(),
    }


def test_invalid_backtick_info_string_does_not_hide_internal_image_marker(
    tmp_path: Path,
) -> None:
    from docwen_plugin_markdown.preprocessor import materialize_image_placeholders

    image = tmp_path / "visible.png"
    text = f"```bad`info\n{{{{IMAGE:{image.as_posix()}}}}}\n"

    result = materialize_image_placeholders(text)

    assert "```bad`info" in result
    assert "{{IMAGE:" not in result
    assert "![visible.png]" in result


def test_xlsx_sized_images_remain_two_columns_and_become_drawings(tmp_path: Path) -> None:
    (tmp_path / "pixel.png").write_bytes(_TINY_PNG)
    source = tmp_path / "sized-images.md"
    source.write_text(
        "| Image | Note |\n"
        "| --- | --- |\n"
        "| ![[pixel.png|wiki alt|20x10]] | wiki sentinel |\n"
        "| ![pixel](pixel.png =30x15) | markdown sentinel |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format="xlsx",
        config_values=_link_config(),
    )

    result = MdToXlsxConverter().convert(context)

    assert result.success is True, result.error
    workbook = load_workbook(Path(result.artifacts[0].staging_path))
    try:
        worksheet = workbook.active
        assert worksheet is not None
        assert worksheet.max_column == 2
        assert worksheet["B2"].value == "wiki sentinel"
        assert worksheet["B3"].value == "markdown sentinel"
        images = vars(worksheet).get("_images")
        assert isinstance(images, list)
        assert len(images) == 2
        drawings = sorted(
            (
                image.anchor._from.row,
                image.anchor._from.col,
                round(image.anchor.ext.cx / 9525),
                round(image.anchor.ext.cy / 9525),
            )
            for image in images
        )
        assert drawings == [(1, 0, 20, 10), (2, 0, 30, 15)]
    finally:
        workbook.close()


def test_csv_sized_images_downgrade_to_filenames_without_column_split(
    tmp_path: Path,
) -> None:
    (tmp_path / "pixel.png").write_bytes(_TINY_PNG)
    source = tmp_path / "sized-images.csv.md"
    source.write_text(
        "| Wiki | Markdown |\n| --- | --- |\n| ![[pixel.png|wiki alt|20x10]] | ![pixel](pixel.png =30x15) |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format="csv",
        config_values=_link_config(),
    )

    result = MdToCsvConverter().convert(context)

    assert result.success is True, result.error
    output = Path(result.artifacts[0].staging_path)
    with output.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.reader(handle))
    assert rows[1] == ["pixel.png", "pixel.png"]


@pytest.mark.parametrize(
    ("child_table", "expected_rows"),
    [
        ("H | N\n--- | ---\nleft | right\n", [["H", "N"], ["left", "right"]]),
        ("| H |\n| --- |\n| value |\n", [["H"], ["value"]]),
        (
            "> | H | N |\n> | --- | --- |\n> | left | right |\n",
            [["H", "N"], ["left", "right"]],
        ),
        (
            "- Item\n    | H | N |\n    | --- | --- |\n    | left | right |\n",
            [["H", "N"], ["left", "right"]],
        ),
    ],
)
@pytest.mark.parametrize("target_format", ["xlsx", "csv"])
def test_spreadsheet_standalone_embed_preserves_supported_child_tables(
    tmp_path: Path,
    child_table: str,
    expected_rows: list[list[str]],
    target_format: str,
) -> None:
    (tmp_path / "child.md").write_text(child_table, encoding="utf-8")
    source = tmp_path / f"standalone-child-{target_format}.md"
    source.write_text("![[child.md]]", encoding="utf-8")
    context, _workspace = make_context(
        str(source),
        target_format=target_format,
        config_values=_link_config(),
    )

    converter = MdToXlsxConverter() if target_format == "xlsx" else MdToCsvConverter()
    result = converter.convert(context)

    assert result.success is True, result.error
    output = Path(result.artifacts[0].staging_path)
    if target_format == "xlsx":
        workbook = load_workbook(output)
        try:
            worksheet = workbook.active
            assert worksheet is not None
            actual_rows = [
                [worksheet.cell(row=row, column=column).value for column in range(1, len(expected_rows[0]) + 1)]
                for row in range(1, len(expected_rows) + 1)
            ]
        finally:
            workbook.close()
    else:
        with output.open(encoding="utf-8-sig", newline="") as handle:
            actual_rows = list(csv.reader(handle))
    assert actual_rows == expected_rows


@pytest.mark.parametrize("target_format", ["xlsx", "csv"])
def test_table_cell_embed_keeps_full_line_context_around_inline_code(
    tmp_path: Path,
    target_format: str,
) -> None:
    (tmp_path / "child.md").write_text("A\nB\n", encoding="utf-8")
    source = tmp_path / f"inline-code-context-{target_format}.md"
    source.write_text(
        "Value | Note\n--- | ---\n`guard` ![[child.md]] | ok\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format=target_format,
        config_values=_link_config(),
    )

    converter = MdToXlsxConverter() if target_format == "xlsx" else MdToCsvConverter()
    result = converter.convert(context)

    assert result.success is True, result.error
    output = Path(result.artifacts[0].staging_path)
    if target_format == "xlsx":
        workbook = load_workbook(output)
        try:
            worksheet = workbook.active
            assert worksheet is not None
            value = worksheet["A2"].value
            note = worksheet["B2"].value
            assert worksheet.max_column == 2
        finally:
            workbook.close()
    else:
        with output.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        value = rows[1][0]
        note = rows[1][1]
        assert len(rows[1]) == 2
    assert isinstance(value, str)
    assert "A\nB" in value
    assert note == "ok"


@pytest.mark.parametrize("target_format", ["xlsx", "csv"])
def test_no_outer_pipe_sized_image_remains_one_cell(
    tmp_path: Path,
    target_format: str,
) -> None:
    (tmp_path / "pixel.png").write_bytes(_TINY_PNG)
    source = tmp_path / f"no-outer-image-{target_format}.md"
    source.write_text(
        "Image | Note\n--- | ---\n![[pixel.png|20x10]] | ok\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format=target_format,
        config_values=_link_config(),
    )

    converter = MdToXlsxConverter() if target_format == "xlsx" else MdToCsvConverter()
    result = converter.convert(context)

    assert result.success is True, result.error
    output = Path(result.artifacts[0].staging_path)
    if target_format == "xlsx":
        workbook = load_workbook(output)
        try:
            worksheet = workbook.active
            assert worksheet is not None
            images = vars(worksheet).get("_images")
            assert worksheet.max_column == 2
            assert worksheet["B2"].value == "ok"
            assert isinstance(images, list)
            assert len(images) == 1
        finally:
            workbook.close()
    else:
        with output.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        assert rows[1] == ["pixel.png", "ok"]


def test_xlsx_literal_and_inline_image_markers_do_not_create_drawings(
    tmp_path: Path,
) -> None:
    image = tmp_path / "literal.png"
    image.write_bytes(_TINY_PNG)
    marker = f"{{{{IMAGE:{image.as_posix()}}}}}"
    source = tmp_path / "literal-markers.xlsx.md"
    source.write_text(
        "| Kind | Value |\n"
        "| --- | --- |\n"
        f"| literal | {marker} |\n"
        f"| inline | `{marker}` |\n"
        "| control | ![control](literal.png) |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format="xlsx",
        config_values=_link_config(),
    )

    result = MdToXlsxConverter().convert(context)

    assert result.success is True, result.error
    workbook = load_workbook(Path(result.artifacts[0].staging_path))
    try:
        worksheet = workbook.active
        assert worksheet is not None
        assert worksheet["B2"].value == marker
        assert marker in str(worksheet["B3"].value)
        images = vars(worksheet).get("_images")
        assert isinstance(images, list)
        assert len(images) == 1
        assert images[0].anchor._from.row == 3
    finally:
        workbook.close()


def test_csv_literal_and_inline_image_markers_stay_literal(
    tmp_path: Path,
) -> None:
    image = tmp_path / "literal.png"
    image.write_bytes(_TINY_PNG)
    marker = f"{{{{IMAGE:{image.as_posix()}}}}}"
    source = tmp_path / "literal-markers.csv.md"
    source.write_text(
        f"| Literal | Inline | Control |\n| --- | --- | --- |\n| {marker} | `{marker}` | ![control](literal.png) |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format="csv",
        config_values=_link_config(),
    )

    result = MdToCsvConverter().convert(context)

    assert result.success is True, result.error
    with Path(result.artifacts[0].staging_path).open(
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        rows = list(csv.reader(handle))
    assert rows[1][0] == marker
    assert marker in rows[1][1]
    assert rows[1][2] == "literal.png"


def test_xlsx_template_trusts_only_markers_present_before_source_fill(
    tmp_path: Path,
) -> None:
    image = tmp_path / "logo.png"
    image.write_bytes(_TINY_PNG)
    marker = "{{IMAGE:logo.png}}"
    template = tmp_path / "image-capability-template.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = marker
    worksheet["B1"] = "{{payload}}"
    worksheet["A3"] = "{{" + chr(0x2193) + "Image}}"
    workbook.save(template)
    workbook.close()

    source = tmp_path / "image-capability-source.md"
    source.write_text(
        f'---\npayload: "{marker}"\n---\n\n| Image | Note |\n| --- | --- |\n| {marker} | source marker |\n',
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format="xlsx",
        options={"template_name": str(template)},
        config_values=_link_config(),
    )

    result = MdToXlsxConverter().convert(context)

    assert result.success is True, result.error
    output = load_workbook(Path(result.artifacts[0].staging_path))
    try:
        sheet = output.active
        assert sheet is not None
        assert sheet["A1"].value is None
        assert sheet["B1"].value == marker
        assert sheet["A3"].value == marker
        images = vars(sheet).get("_images")
        assert isinstance(images, list)
        assert len(images) == 1
        assert result.artifacts[0].metadata["image_placeholders"] == 1
    finally:
        output.close()


def test_xlsx_template_origin_image_marker_accepts_braced_filename(
    tmp_path: Path,
) -> None:
    image = tmp_path / "a{b}.png"
    image.write_bytes(_TINY_PNG)
    template = tmp_path / "braced-image-template.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "{{IMAGE:a{b}.png}}"
    workbook.save(template)
    workbook.close()
    source = tmp_path / "braced-template-source.md"
    source.write_text("| Value |\n| --- |\n| body |\n", encoding="utf-8")
    context, _workspace = make_context(
        str(source),
        target_format="xlsx",
        options={"template_name": str(template)},
        config_values=_link_config(),
    )

    result = MdToXlsxConverter().convert(context)

    assert result.success is True, result.error
    output = load_workbook(Path(result.artifacts[0].staging_path))
    try:
        sheet = output.active
        assert sheet is not None
        images = vars(sheet).get("_images")
        assert isinstance(images, list)
        assert len(images) == 1
        assert sheet["A1"].value is None
    finally:
        output.close()


def test_xlsx_yaml_front_matter_is_not_link_processed_before_template_fill(
    tmp_path: Path,
) -> None:
    image = tmp_path / "logo.png"
    image.write_bytes(_TINY_PNG)
    template = tmp_path / "yaml-link-template.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "{{image_syntax}}"
    worksheet["B1"] = "{{link_syntax}}"
    worksheet["A3"] = "{{" + chr(0x2193) + "Value}}"
    workbook.save(template)
    workbook.close()

    source = tmp_path / "yaml-link-source.md"
    source.write_text(
        "\ufeff---\n"
        'image_syntax: "![logo](logo.png)"\n'
        'link_syntax: "[drop](https://drop.example)"\n'
        "---\n\n"
        "| Value |\n"
        "| --- |\n"
        "| body |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format="xlsx",
        options={"template_name": str(template)},
        config_values=_link_config(
            markdown_mode="remove",
            markdown_image_mode="embed",
        ),
    )

    result = MdToXlsxConverter().convert(context)

    assert result.success is True, result.error
    output = load_workbook(Path(result.artifacts[0].staging_path))
    try:
        sheet = output.active
        assert sheet is not None
        assert sheet["A1"].value == "![logo](logo.png)"
        assert sheet["B1"].value == "[drop](https://drop.example)"
        images = vars(sheet).get("_images")
        assert isinstance(images, list)
        assert not images
    finally:
        output.close()
