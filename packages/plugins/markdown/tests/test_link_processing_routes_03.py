"""Focused tests split from test_link_processing_routes.py."""

from __future__ import annotations

from ._link_processing_routes_support import (
    _TINY_PNG,
    _WP_NS,
    ET,
    Any,
    Document,
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


def test_docx_yaml_front_matter_is_not_link_processed_before_template_fill(
    tmp_path: Path,
) -> None:
    template = tmp_path / "yaml-link-template.docx"
    template_doc = Document()
    template_doc.add_paragraph("{{title}}")
    template_doc.add_paragraph("{{正文}}")
    template_doc.save(str(template))
    source = tmp_path / "yaml-link-source.md"
    source.write_text(
        '\ufeff---\ntitle: "[keep literal](https://yaml.example)"\n---\n\nBody [drop](https://body.example).\n',
        encoding="utf-8",
    )

    observation = _convert_docx(
        source,
        _link_config(markdown_mode="remove"),
        options={"template_name": str(template)},
    )

    assert "[keep literal](https://yaml.example)" in observation.text
    assert "yaml.example" in observation.text
    assert "body.example" not in observation.text


def test_csv_literal_percent_7c_filename_round_trips_exactly_once(
    tmp_path: Path,
) -> None:
    image = tmp_path / "a%7Cb.png"
    image.write_bytes(_TINY_PNG)
    source = tmp_path / "encoded-percent-pipe.md"
    source.write_text(
        "| Image |\n| --- |\n| ![encoded](a%257Cb.png) |\n",
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
    assert rows == [["Image"], ["a%7Cb.png"]]


def test_csv_yaml_front_matter_is_not_link_processed_before_template_fill(
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
        target_format="csv",
        options={"template_name": str(template)},
        config_values=_link_config(
            markdown_mode="remove",
            markdown_image_mode="embed",
        ),
    )

    result = MdToCsvConverter().convert(context)

    assert result.success is True, result.error
    with Path(result.artifacts[0].staging_path).open(
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        rows = list(csv.reader(handle))
    assert rows[0] == ["![logo](logo.png)", "[drop](https://drop.example)"]


@pytest.mark.parametrize("target_format", ["docx", "xlsx", "csv"])
@pytest.mark.parametrize(
    ("filename", "destination"),
    [
        ("my(file).png", r"my\(file\).png"),
        ("my#file.png", "my%23file.png"),
        ("a%20b.png", "a%2520b.png"),
    ],
)
def test_markdown_image_destination_preserves_literal_filename_delimiters(
    tmp_path: Path,
    target_format: str,
    filename: str,
    destination: str,
) -> None:
    image = tmp_path / filename
    image.write_bytes(_TINY_PNG)
    source = tmp_path / f"literal-delimiter-{target_format}.md"
    if target_format == "docx":
        source.write_text(f"![literal delimiter]({destination})", encoding="utf-8")
        observation = _convert_docx(source, _link_config())
        assert len(observation.media_names) == 1
        assert "File not found" not in observation.text
        return

    source.write_text(
        f"| Image | Note |\n| --- | --- |\n| ![literal delimiter]({destination}) | control |\n",
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

    if target_format == "xlsx":
        output = load_workbook(Path(result.artifacts[0].staging_path))
        try:
            sheet = output.active
            assert sheet is not None
            images = vars(sheet).get("_images")
            assert isinstance(images, list)
            assert len(images) == 1
        finally:
            output.close()
    else:
        with Path(result.artifacts[0].staging_path).open(
            encoding="utf-8-sig",
            newline="",
        ) as handle:
            rows = list(csv.reader(handle))
        assert rows[1][0] == filename


@pytest.mark.parametrize("target_format", ["docx", "xlsx", "csv"])
def test_scoped_image_marker_round_trips_braced_filename(
    tmp_path: Path,
    target_format: str,
) -> None:
    image = tmp_path / "a{b}.png"
    image.write_bytes(_TINY_PNG)
    source = tmp_path / f"braced-{target_format}.md"
    if target_format == "docx":
        source.write_text("![braced](a{b}.png)", encoding="utf-8")
        observation = _convert_docx(source, _link_config())
        assert len(observation.media_names) == 1
        assert "IMAGE@" not in observation.text
        return

    source.write_text(
        "| Image | Note |\n| --- | --- |\n| ![braced](a{b}.png) | control |\n",
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

    if target_format == "xlsx":
        output = load_workbook(Path(result.artifacts[0].staging_path))
        try:
            sheet = output.active
            assert sheet is not None
            images = vars(sheet).get("_images")
            assert isinstance(images, list)
            assert len(images) == 1
            assert "IMAGE@" not in str(sheet["A2"].value)
        finally:
            output.close()
    else:
        with Path(result.artifacts[0].staging_path).open(
            encoding="utf-8-sig",
            newline="",
        ) as handle:
            rows = list(csv.reader(handle))
        assert rows[1][0] == "a{b}.png"


@pytest.mark.parametrize(
    ("syntax", "config_override"),
    [
        ("[**bold**](https://example.com)", {"markdown_mode": "keep"}),
        ("![**alt**](pixel.png)", {"markdown_image_mode": "keep"}),
        ("[[target.md|**bold**]]", {"wiki_mode": "keep"}),
        ("![[pixel.png|**alt**]]", {"wiki_image_mode": "keep"}),
        ("![[child.md|**alias**]]", {"md_file_mode": "keep"}),
    ],
)
def test_docx_keep_modes_render_exact_source_syntax(
    tmp_path: Path,
    syntax: str,
    config_override: dict[str, Any],
) -> None:
    (tmp_path / "pixel.png").write_bytes(_TINY_PNG)
    (tmp_path / "target.md").write_text("target\n", encoding="utf-8")
    (tmp_path / "child.md").write_text("child\n", encoding="utf-8")
    source = tmp_path / "keep-source.md"
    source.write_text(syntax, encoding="utf-8")

    observation = _convert_docx(source, _link_config(**config_override))

    assert observation.text == syntax


def test_docx_hyperlink_preserves_existing_markdown_label_escapes(
    tmp_path: Path,
) -> None:
    source = tmp_path / "escaped-label.md"
    source.write_text(r"[\*literal\*](https://example.com)", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert observation.text == "*literal*"
    assert observation.hyperlink_targets == {"https://example.com"}
    assert "<w:i" not in observation.document_xml


@pytest.mark.parametrize(
    "syntax",
    [
        "[a  \nb](https://example.com)",
        "[a\\\nb](https://example.com)",
    ],
)
def test_docx_hyperlink_keeps_hard_break_label_separation(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = tmp_path / "hard-break-label.md"
    source.write_text(syntax, encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert observation.text == "a b"
    assert "<w:br" in observation.document_xml
    assert observation.hyperlink_targets == {"https://example.com"}


@pytest.mark.parametrize(
    "syntax",
    [
        "![remote](//example.com/pixel.png)",
        "![[//example.com/pixel.png]]",
    ],
)
def test_protocol_relative_images_never_embed_a_local_collision(
    tmp_path: Path,
    syntax: str,
) -> None:
    collision = tmp_path / "example.com" / "pixel.png"
    collision.parent.mkdir()
    collision.write_bytes(_TINY_PNG)
    source = tmp_path / "protocol-relative.md"
    source.write_text(syntax, encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert not observation.media_names
    assert "Remote embed fetching is unsupported" in observation.text
    assert "File not found" not in observation.text


@pytest.mark.parametrize("filename", ["a%20b.png", "a#b.png"])
def test_docx_file_uri_image_decodes_path_exactly_once(
    tmp_path: Path,
    filename: str,
) -> None:
    image = tmp_path / filename
    image.write_bytes(_TINY_PNG)
    uri = image.as_uri().replace("file:", "FILE:", 1)
    source = tmp_path / "file-uri.md"
    source.write_text(f"![file uri](<{uri}>)", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert len(observation.media_names) == 1
    assert "File not found" not in observation.text


def test_docx_wiki_escaped_hash_targets_literal_filename(tmp_path: Path) -> None:
    (tmp_path / "my#file.png").write_bytes(_TINY_PNG)
    (tmp_path / "my#file.md").write_text("target\n", encoding="utf-8")
    source = tmp_path / "wiki-escaped-hash.md"
    source.write_text(
        r"![[my\#file.png]] and [[my\#file.md|Shown]]",
        encoding="utf-8",
    )

    observation = _convert_docx(source, _link_config())

    assert len(observation.media_names) == 1
    assert "Shown" in observation.text
    assert any("my%23file.md" in target for target in observation.hyperlink_targets)


def test_docx_nested_parentheses_hyperlink_remains_one_relationship(
    tmp_path: Path,
) -> None:
    source = tmp_path / "nested-parens.md"
    source.write_text(
        "[Target](https://example.com/a_(b))",
        encoding="utf-8",
    )

    observation = _convert_docx(source, _link_config())

    assert observation.text == "Target"
    assert observation.hyperlink_targets == {"https://example.com/a_(b)"}


def test_docx_multiline_image_title_does_not_enter_the_filename(
    tmp_path: Path,
) -> None:
    (tmp_path / "pixel.png").write_bytes(_TINY_PNG)
    source = tmp_path / "multiline-title.md"
    source.write_text(
        '![pixel](pixel.png\n "caption")',
        encoding="utf-8",
    )

    observation = _convert_docx(source, _link_config())

    assert len(observation.media_names) == 1
    assert "File not found" not in observation.text


def test_docx_image_size_uses_source_dpi_and_page_width_clamp(
    tmp_path: Path,
) -> None:
    from PIL import Image

    image = tmp_path / "high-dpi.png"
    Image.new("RGB", (300, 150), "white").save(image, dpi=(300, 300))
    source = tmp_path / "high-dpi.md"
    source.write_text("![[high-dpi.png|300x150]]", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    document_root = ET.fromstring(observation.document_xml)
    extent = next(document_root.iter(f"{{{_WP_NS}}}extent"))
    assert int(extent.attrib["cx"]) == pytest.approx(914400, abs=8)
    assert int(extent.attrib["cy"]) == pytest.approx(457200, abs=8)

    wide = tmp_path / "wide.png"
    Image.new("RGB", (3000, 100), "white").save(wide, dpi=(96, 96))
    source.write_text("![[wide.png|3000x100]]", encoding="utf-8")
    wide_observation = _convert_docx(source, _link_config())
    wide_root = ET.fromstring(wide_observation.document_xml)
    wide_extent = next(wide_root.iter(f"{{{_WP_NS}}}extent"))
    assert int(wide_extent.attrib["cx"]) <= int(6.5 * 914400)


def test_xlsx_nested_table_codec_preserves_escaped_pipes_and_literal_tokens(
    tmp_path: Path,
) -> None:
    fake_scoped_token = "{{DOCWEN_BR@" + ("a" * 32) + ".0000000000000000}}"
    child = tmp_path / "child.md"
    child.write_text(
        r"A \| B {{DOCWEN_BR}} " + fake_scoped_token + "\nC",
        encoding="utf-8",
    )
    source = tmp_path / "table-codec.md"
    source.write_text(
        "| Value | Note |\n| --- | --- |\n| ![[child.md]] | ok |\n",
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
        sheet = workbook.active
        assert sheet is not None
        assert sheet.max_column == 2
        assert sheet["B2"].value == "ok"
        assert sheet["A2"].value == ("A | B {{DOCWEN_BR}} " + fake_scoped_token + "\nC")
    finally:
        workbook.close()


@pytest.mark.parametrize("target_format", ["xlsx", "csv"])
@pytest.mark.parametrize(
    ("syntax", "config_override", "expected_cell"),
    [
        (
            "[[target.md|Shown]]",
            {"wiki_mode": "keep"},
            "[[target.md|Shown]]",
        ),
        (
            "[[target.md|Shown]]",
            {"wiki_mode": "hyperlink"},
            "[[target.md|Shown]]",
        ),
        (
            "[[target.md|Shown|More]]",
            {"wiki_mode": "extract_text"},
            "Shown|More",
        ),
        (
            "![[missing.md|Alias|More]]",
            {"md_file_mode": "keep", "file_not_found_mode": "keep"},
            "![[missing.md|Alias|More]]",
        ),
        (
            "![Alt|More](missing.png)",
            {"markdown_image_mode": "keep"},
            "![Alt|More](missing.png)",
        ),
        (
            "![[pixel.png|Alias|More]]",
            {"wiki_image_mode": "extract_text"},
            "Alias|More",
        ),
        (
            "[Shown|More](https://example.com)",
            {"markdown_mode": "keep"},
            "[Shown|More](https://example.com)",
        ),
        (
            "[Shown|More](https://example.com)",
            {"markdown_mode": "hyperlink"},
            "[Shown|More](https://example.com)",
        ),
    ],
)
def test_spreadsheet_link_replacements_keep_wiki_pipes_in_one_cell(
    tmp_path: Path,
    target_format: str,
    syntax: str,
    config_override: dict[str, Any],
    expected_cell: str,
) -> None:
    (tmp_path / "target.md").write_text("target\n", encoding="utf-8")
    (tmp_path / "pixel.png").write_bytes(_TINY_PNG)
    source = tmp_path / f"table-pipes-{target_format}.md"
    source.write_text(
        f"| Value | Note |\n| --- | --- |\n| {syntax} | ok |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format=target_format,
        config_values=_link_config(**config_override),
    )

    converter = MdToXlsxConverter() if target_format == "xlsx" else MdToCsvConverter()
    result = converter.convert(context)

    assert result.success is True, result.error
    output = Path(result.artifacts[0].staging_path)
    if target_format == "xlsx":
        workbook = load_workbook(output)
        try:
            sheet = workbook.active
            assert sheet is not None
            assert sheet.max_column == 2
            row = [sheet["A2"].value, sheet["B2"].value]
        finally:
            workbook.close()
    else:
        with output.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        row = rows[1]
        assert len(row) == 2
    assert row == [expected_cell, "ok"]
