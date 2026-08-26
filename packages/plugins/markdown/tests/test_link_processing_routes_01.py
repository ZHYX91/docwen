"""Focused tests split from test_link_processing_routes.py."""

from __future__ import annotations

from ._link_processing_routes_support import (
    _TINY_PNG,
    MdToDocxConverter,
    MdToXlsxConverter,
    Path,
    _assert_target,
    _convert_declared_docx,
    _convert_docx,
    _link_config,
    load_workbook,
    make_context,
    pytest,
)

pytestmark = pytest.mark.contract


def test_docx_declared_resource_ignores_physical_sibling_decoy(tmp_path: Path) -> None:
    source_dir = tmp_path / "source-physical"
    declared_dir = tmp_path / "declared-physical"
    source_dir.mkdir()
    declared_dir.mkdir()
    source = source_dir / "report.md"
    source.write_text("![declared](assets/pixel.png)", encoding="utf-8")
    (source_dir / "assets").mkdir()
    (source_dir / "assets" / "pixel.png").write_bytes(b"not-a-png-decoy")
    declared = declared_dir / "opaque.png"
    declared.write_bytes(_TINY_PNG)

    observation = _convert_declared_docx(
        source,
        source_logical_path="documents/report.md",
        resource=declared,
        resource_logical_path="documents/assets/pixel.png",
    )

    assert len(observation.media_names) == 1
    assert "<w:drawing" in observation.document_xml


def test_docx_typed_route_rejects_undeclared_physical_sibling(tmp_path: Path) -> None:
    source = tmp_path / "report.md"
    source.write_text("![undeclared](assets/pixel.png)", encoding="utf-8")
    (tmp_path / "assets").mkdir()
    (tmp_path / "assets" / "pixel.png").write_bytes(_TINY_PNG)

    context, workspace = make_context(
        str(source),
        target_format="docx",
        config_values=_link_config(markdown_image_mode="embed"),
    )
    from docwen_core.models.file_ref import FileRef

    source_ref = FileRef(
        path=str(source),
        format="markdown",
        category="document",
        input_kind="document",
        input_role="source",
        logical_path="documents/report.md",
    )
    context.request.input_refs = [source_ref]
    workspace._input_refs = (source_ref,)
    result = MdToDocxConverter().convert(context)

    assert result.success is False
    assert result.error is not None
    assert "undeclared linked resource" in result.error.message


@pytest.mark.parametrize(
    "syntax",
    [
        "![[assets/pixel.png]]",
        "![[outside.md]]",
        "[[outside.md|outside]]",
        "[outside](outside.md)",
    ],
)
def test_docx_typed_route_rejects_wiki_transclusion_before_filesystem_lookup(
    tmp_path: Path,
    syntax: str,
) -> None:
    source = tmp_path / "report.md"
    source.write_text(syntax, encoding="utf-8")
    (tmp_path / "assets").mkdir()
    (tmp_path / "assets" / "pixel.png").write_bytes(_TINY_PNG)
    (tmp_path / "outside.md").write_text("external sentinel", encoding="utf-8")
    context, workspace = make_context(
        str(source),
        target_format="docx",
        config_values=_link_config(wiki_image_mode="embed", md_file_mode="embed"),
    )
    from docwen_core.models.file_ref import FileRef

    source_ref = FileRef(
        path=str(source),
        format="markdown",
        category="document",
        input_kind="document",
        input_role="source",
        logical_path="documents/report.md",
    )
    context.request.input_refs = [source_ref]
    workspace._input_refs = (source_ref,)

    result = MdToDocxConverter().convert(context)

    assert result.success is False
    assert result.error is not None
    assert "unavailable for declared-input requests" in result.error.message


def test_docx_typed_route_keeps_remote_and_fragment_links(tmp_path: Path) -> None:
    source = tmp_path / "report.md"
    source.write_text("[web](https://example.com) [mail](mailto:a@example.com) [part](#part)", encoding="utf-8")
    declared = tmp_path / "declared.png"
    declared.write_bytes(_TINY_PNG)

    observation = _convert_declared_docx(
        source,
        source_logical_path="documents/report.md",
        resource=declared,
        resource_logical_path="documents/assets/unused.png",
    )

    assert "web" in observation.text
    assert "mail" in observation.text
    assert "part" in observation.text
    assert "https://example.com" in observation.hyperlink_targets
    assert "mailto:a@example.com" in observation.hyperlink_targets


def test_docx_typed_route_does_not_bind_image_syntax_inside_code(tmp_path: Path) -> None:
    source = tmp_path / "report.md"
    source.write_text("`![inline](missing.png)`\n\n```md\n![fenced](missing.png)\n```\n", encoding="utf-8")
    declared = tmp_path / "declared.png"
    declared.write_bytes(_TINY_PNG)

    observation = _convert_declared_docx(
        source,
        source_logical_path="documents/report.md",
        resource=declared,
        resource_logical_path="documents/assets/unused.png",
    )

    assert "![inline](missing.png)" in observation.text
    assert "![fenced](missing.png)" in observation.text
    assert not observation.media_names


@pytest.mark.parametrize("kind", ["wiki", "markdown"])
@pytest.mark.parametrize("mode", ["keep", "extract_text", "remove", "hyperlink"])
def test_docx_consumes_request_scoped_non_embed_link_modes(
    tmp_path: Path,
    kind: str,
    mode: str,
) -> None:
    """The DOCX route distinguishes literal, text-only, removed, and clickable links."""
    (tmp_path / "guide.md").write_text("Guide target", encoding="utf-8")
    if kind == "wiki":
        syntax = "[[guide|Wiki Label]]"
        label = "Wiki Label"
        target = "guide.md"
    else:
        syntax = "[Markdown Label](https://example.com/markdown)"
        label = "Markdown Label"
        target = "https://example.com/markdown"

    source = tmp_path / f"non-embed-{kind}-{mode}.md"
    source.write_text(f"Before {syntax} After", encoding="utf-8")
    observation = _convert_docx(
        source,
        _link_config(
            wiki_mode=mode if kind == "wiki" else "keep",
            markdown_mode=mode if kind == "markdown" else "keep",
        ),
    )

    if mode == "keep":
        assert syntax in observation.text
        assert not observation.hyperlink_targets
    elif mode == "extract_text":
        assert label in observation.text
        assert syntax not in observation.text
        assert not observation.hyperlink_targets
    elif mode == "remove":
        assert label not in observation.text
        assert syntax not in observation.text
        assert not observation.hyperlink_targets
    else:
        assert label in observation.text
        assert syntax not in observation.text
        assert "<w:hyperlink" in observation.document_xml
        _assert_target(observation, target)


def test_xlsx_hyperlink_mode_preserves_source_syntax_in_cells(tmp_path: Path) -> None:
    """Spreadsheet targets have no DOCX relationship sink, so hyperlink mode keeps syntax."""
    source = tmp_path / "spreadsheet-links.md"
    source.write_text(
        "| Wiki | Markdown |\n| --- | --- |\n| [[guide]] | [Markdown Label](https://example.com/markdown) |\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source),
        target_format="xlsx",
        config_values=_link_config(wiki_mode="hyperlink", markdown_mode="hyperlink"),
    )

    result = MdToXlsxConverter().convert(context)

    assert result.success is True, result.error
    workbook = load_workbook(Path(result.artifacts[0].staging_path))
    try:
        worksheet = workbook.active
        assert worksheet is not None
        assert worksheet["A2"].value == "[[guide]]"
        assert worksheet["B2"].value == "[Markdown Label](https://example.com/markdown)"
    finally:
        workbook.close()


@pytest.mark.parametrize("kind", ["wiki", "markdown"])
@pytest.mark.parametrize("mode", ["keep", "extract_text", "remove", "embed"])
def test_docx_consumes_request_scoped_image_modes(
    tmp_path: Path,
    kind: str,
    mode: str,
) -> None:
    """Wiki and standard Markdown images obey independent request policies."""
    image = tmp_path / "pixel.png"
    image.write_bytes(_TINY_PNG)
    image_target = image.as_posix()
    syntax = f"![[{image_target}|pixel.png]]" if kind == "wiki" else f"![pixel.png]({image_target})"
    source = tmp_path / f"image-{kind}-{mode}.md"
    source.write_text(f"Before\n\n{syntax}\n\nAfter", encoding="utf-8")

    observation = _convert_docx(
        source,
        _link_config(
            wiki_image_mode=mode if kind == "wiki" else "embed",
            markdown_image_mode=mode if kind == "markdown" else "embed",
        ),
    )

    if mode == "keep":
        assert syntax in observation.text
        assert not observation.media_names
    elif mode == "extract_text":
        assert "pixel.png" in observation.text
        assert syntax not in observation.text
        assert "[Image:" not in observation.text
        assert not observation.media_names
    elif mode == "remove":
        assert "pixel.png" not in observation.text
        assert syntax not in observation.text
        assert not observation.media_names
    else:
        assert len(observation.media_names) == 1
        assert "<w:drawing" in observation.document_xml


@pytest.mark.parametrize("mode", ["keep", "extract_text", "remove", "embed"])
def test_docx_consumes_request_scoped_embedded_markdown_modes(
    tmp_path: Path,
    mode: str,
) -> None:
    """The real DOCX route applies all four embedded-Markdown policies."""
    child = tmp_path / "child.md"
    child.write_text("Embedded child sentinel.", encoding="utf-8")
    syntax = "![[child.md]]"
    source = tmp_path / f"embedded-markdown-{mode}.md"
    source.write_text(f"Before\n\n{syntax}\n\nAfter", encoding="utf-8")

    observation = _convert_docx(source, _link_config(md_file_mode=mode))

    if mode == "keep":
        assert syntax in observation.text
        assert "Embedded child sentinel." not in observation.text
    elif mode == "extract_text":
        assert "child" in observation.text
        assert syntax not in observation.text
        assert "Embedded child sentinel." not in observation.text
        assert "[Image:" not in observation.text
    elif mode == "remove":
        assert "child" not in observation.text
        assert syntax not in observation.text
        assert "Embedded child sentinel." not in observation.text
    else:
        assert "Embedded child sentinel." in observation.text
        assert syntax not in observation.text


def test_docx_max_depth_is_request_scoped_across_sequential_conversions(tmp_path: Path) -> None:
    """A shallow request truncates recursion without contaminating the next request."""
    grandchild = tmp_path / "grandchild.md"
    grandchild.write_text("Grandchild depth sentinel.", encoding="utf-8")
    child = tmp_path / "child.md"
    child.write_text("Child level.\n\n![[grandchild.md]]", encoding="utf-8")
    source = tmp_path / "depth-source.md"
    source.write_text("![[child.md]]", encoding="utf-8")

    shallow = _convert_docx(source, _link_config(md_file_mode="embed", max_depth=1))
    deep = _convert_docx(source, _link_config(md_file_mode="embed", max_depth=3))

    assert "Grandchild depth sentinel." not in shallow.text
    assert "Grandchild depth sentinel." in deep.text


@pytest.mark.parametrize(
    ("auto_link_bare_url", "expected_targets"),
    [
        (False, {"https://angle.example/path"}),
        (True, {"https://bare.example/path", "https://angle.example/path"}),
    ],
)
def test_docx_bare_url_policy_does_not_change_angle_bracket_autolinks(
    tmp_path: Path,
    auto_link_bare_url: bool,
    expected_targets: set[str],
) -> None:
    """Bare-URL opt-in is request-scoped; explicit ``<url>`` autolinks stay clickable."""
    source = tmp_path / f"bare-url-{auto_link_bare_url}.md"
    source.write_text(
        "Bare https://bare.example/path and angle <https://angle.example/path>.",
        encoding="utf-8",
    )

    observation = _convert_docx(
        source,
        _link_config(
            markdown_mode="hyperlink",
            auto_link_bare_url=auto_link_bare_url,
        ),
    )

    assert observation.hyperlink_targets == expected_targets


def test_docx_keeps_angle_bracket_email_autolinks_without_bare_url_plugin(
    tmp_path: Path,
) -> None:
    """CommonMark email autolinks do not depend on Mistune's bare-URL plugin."""
    source = tmp_path / "angle-email.md"
    source.write_text("Contact <support@example.com>.", encoding="utf-8")

    observation = _convert_docx(
        source,
        _link_config(auto_link_bare_url=False),
    )

    assert observation.hyperlink_targets == {"mailto:support@example.com"}


def test_docx_embeds_relative_image_inside_table_without_pipe_corruption(
    tmp_path: Path,
) -> None:
    """Image placeholders remain one table cell and resolve from the source directory."""
    (tmp_path / "pixel.png").write_bytes(_TINY_PNG)
    source = tmp_path / "relative-table-image.md"
    source.write_text(
        "| Image | Note |\n| --- | --- |\n| ![pixel](pixel.png) | table sentinel |\n",
        encoding="utf-8",
    )

    observation = _convert_docx(
        source,
        _link_config(markdown_image_mode="embed"),
    )

    assert observation.media_names
    assert "table sentinel" in observation.text
    assert "<w:drawing" in observation.document_xml


@pytest.mark.parametrize(
    ("mode", "expected_text"),
    [
        ("keep", "[[missing.md|Shown]]"),
        ("ignore", "Before  After"),
        ("placeholder", "[File not found: missing.md]"),
    ],
)
def test_docx_missing_wiki_hyperlink_consumes_request_error_policy(
    tmp_path: Path,
    mode: str,
    expected_text: str,
) -> None:
    source = tmp_path / f"missing-wiki-{mode}.md"
    source.write_text("Before [[missing.md|Shown]] After", encoding="utf-8")

    observation = _convert_docx(
        source,
        _link_config(file_not_found_mode=mode),
    )

    assert expected_text in observation.text
    assert not observation.hyperlink_targets


def test_docx_missing_standard_image_keep_renders_exact_source(
    tmp_path: Path,
) -> None:
    syntax = "![Missing](missing.png)"
    source = tmp_path / "missing-standard-image.md"
    source.write_text(syntax, encoding="utf-8")

    observation = _convert_docx(
        source,
        _link_config(file_not_found_mode="keep"),
    )

    assert observation.text == syntax
    assert not observation.media_names


def test_docx_wiki_hyperlink_quotes_space_path_and_fragment(tmp_path: Path) -> None:
    target = tmp_path / "guide space.md"
    target.write_text("# Part 1\n", encoding="utf-8")
    source = tmp_path / "space-wiki.md"
    source.write_text("[[guide space.md#Part 1|Spaced guide]]", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert "Spaced guide" in observation.text
    assert any("guide%20space.md#Part%201" in relationship for relationship in observation.hyperlink_targets)


def test_docx_fragment_only_wiki_link_has_no_false_hyperlink_affordance(
    tmp_path: Path,
) -> None:
    source = tmp_path / "fragment-only.md"
    source.write_text("[[#section-name|Section]]", encoding="utf-8")

    observation = _convert_docx(source, _link_config(wiki_mode="hyperlink"))

    assert observation.text == "Section"
    assert not observation.hyperlink_targets
    assert "<w:hyperlink" not in observation.document_xml
    assert "<w:bookmarkStart" not in observation.document_xml
    assert "<w:u" not in observation.document_xml


def test_docx_fragment_only_standard_link_has_no_false_hyperlink_affordance(
    tmp_path: Path,
) -> None:
    source = tmp_path / "standard-fragment-only.md"
    source.write_text("[Section](#section-name)", encoding="utf-8")

    observation = _convert_docx(
        source,
        _link_config(markdown_mode="hyperlink"),
    )

    assert observation.text == "Section"
    assert not observation.hyperlink_targets
    assert "<w:hyperlink" not in observation.document_xml
    assert "<w:u" not in observation.document_xml


def test_docx_windows_absolute_hyperlink_keeps_cross_drive_target(tmp_path: Path) -> None:
    source = tmp_path / "windows-absolute.md"
    source.write_text("[Cross drive](<Z:/docs/a b.md#frag space>)", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert observation.hyperlink_targets == {"file:///Z:/docs/a%20b.md#frag%20space"}


def test_docx_windows_absolute_hyperlink_preserves_query_and_fragment(
    tmp_path: Path,
) -> None:
    source = tmp_path / "windows-absolute-query.md"
    source.write_text(
        "[Cross drive](<Z:/docs/a b.md?download=1#frag space>)",
        encoding="utf-8",
    )

    observation = _convert_docx(source, _link_config())

    assert observation.hyperlink_targets == {"file:///Z:/docs/a%20b.md?download=1#frag%20space"}


def test_docx_relative_hyperlink_preserves_query_and_fragment(tmp_path: Path) -> None:
    target = tmp_path / "target.md"
    target.write_text("# Part\n", encoding="utf-8")
    source = tmp_path / "relative-query.md"
    source.write_text("[Target](target.md?download=1#Part)", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert len(observation.hyperlink_targets) == 1
    relationship = next(iter(observation.hyperlink_targets))
    assert relationship.endswith("/target.md?download=1#Part")


def test_docx_uppercase_http_scheme_remains_clickable(tmp_path: Path) -> None:
    source = tmp_path / "uppercase-scheme.md"
    source.write_text("[Target](HTTP://example.com/path)", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert observation.hyperlink_targets == {"http://example.com/path"}


def test_docx_protocol_relative_link_does_not_become_a_local_file(
    tmp_path: Path,
) -> None:
    source = tmp_path / "protocol-relative.md"
    source.write_text("[Target](//example.com/path)", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert observation.text == "Target"
    assert not observation.hyperlink_targets


def test_docx_wiki_hyperlink_escapes_brackets_in_display_label(tmp_path: Path) -> None:
    target = tmp_path / "target.md"
    target.write_text("body\n", encoding="utf-8")
    source = tmp_path / "wiki-bracket-label.md"
    source.write_text("[[target.md|See [draft]]]", encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert observation.text == "See [draft]"
    assert len(observation.hyperlink_targets) == 1


def test_docx_markdown_image_with_title_embeds_relative_file(tmp_path: Path) -> None:
    (tmp_path / "pixel.png").write_bytes(_TINY_PNG)
    source = tmp_path / "image-title.md"
    source.write_text('![pixel](pixel.png "caption")', encoding="utf-8")

    observation = _convert_docx(source, _link_config())

    assert len(observation.media_names) == 1
    assert "<w:drawing" in observation.document_xml
    assert "File not found" not in observation.text
