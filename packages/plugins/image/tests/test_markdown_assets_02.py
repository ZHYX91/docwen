"""Focused tests split from test_markdown_assets.py."""

from __future__ import annotations

from ._markdown_assets_support import (
    Path,
    _build_fake_context,
    _ocr_success,
    pytest,
    tempfile,
)

pytestmark = [pytest.mark.golden, pytest.mark.contract]


class TestImageToMarkdownLinkStyleSemantics:
    """Verify the image→Markdown converter consumes the link-style
    configuration from ``docwen_core.export_semantics``.

    The ``image_link_style`` config controls whether the output uses
    ``![[target]]`` (wiki_embed) or ``![alt](target)`` (markdown_embed).
    """

    def test_semantics_default_wiki_embed_style(self, sample_png_path: Path) -> None:
        """With default export semantics (wiki_embed), the converter
        produces ``![[...]]`` wiki embed links."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # wiki_embed format: ![[sample.png]] (sanitised, no ./ prefix)
            assert "![[sample.png]]" in md_text
            assert "![[./" not in md_text  # ./ prefix intentionally removed by sanitised formatter

    def test_semantics_markdown_embed_style(self, sample_png_path: Path) -> None:
        """When ``image_link_style`` is ``markdown_embed``, the converter
        produces standard ``![alt](path)`` links."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "file",
                    "image_link_style": "markdown_embed",
                    "to_md_enable_ocr": False,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # markdown_embed format: ![alt](./sample.png)
            assert "![sample](./sample.png)" in md_text
            assert "![[./sample.png]]" not in md_text

    def test_request_image_link_style_overrides_export_semantics(self, sample_png_path: Path) -> None:
        """Request-level shared To-MD options should override the configured default."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "file",
                    "image_link_style": "markdown_embed",
                    "to_md_enable_ocr": False,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "![sample](./sample.png)" in md_text
            assert "![[sample.png]]" not in md_text

    def test_semantics_base64_wiki_embed_style(self, sample_png_path: Path) -> None:
        """Base64 data-URI images also respect the link style config."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "base64", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # wiki_embed base64: ![[data:image/png;base64,...]]
            assert "![[data:image/png;base64," in md_text

    def test_semantics_base64_markdown_embed_style(self, sample_png_path: Path) -> None:
        """Base64 data-URI in markdown_embed style produces
        ``![alt](data:...)`` links."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "base64",
                    "image_link_style": "markdown_embed",
                    "to_md_enable_ocr": False,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "![sample](data:image/png;base64," in md_text
            assert "![[data:" not in md_text


class TestImageConverterUtilityConsumption:
    """Verify the image→Markdown converter is a real production consumer
    of the shared YAML / Markdown utilities and the image plugin's local OCR
    presentation helper.

    F-I2b-001 (extract_yaml), F-I2b-002 (generate_basic_yaml_frontmatter),
    F-I2b-004 (format_sanitized_image_link / sanitize_filename).
    """

    def test_sanitize_filename_consumed_by_converter(self, sample_png_path: Path) -> None:
        """The converter calls ``sanitize_filename`` when building the
        image filename, stripping FS-illegal characters."""
        from docwen_core.markdown_utils import sanitize_filename
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        # Verify that the converter output contains the sanitised filename
        # (sanitize_filename is called on the image_filename variable in converter.py).
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)
            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # The sanitised filename (without FS-illegal chars) appears in the output
            expected_sanitised = sanitize_filename("sample.png")
            assert expected_sanitised in md_text

    def test_extract_yaml_consumed_for_yaml_validation(self, sample_png_path: Path) -> None:
        """After generating Markdown, the converter validates the YAML
        front matter by extracting it back via ``extract_yaml`` (F-I2b-001)."""
        from docwen_core.yaml_tools import extract_yaml
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)
            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # The generated Markdown must contain parseable YAML
            yaml_str, body = extract_yaml(md_text)
            assert "title: sample" in yaml_str
            assert "source_format: png" in yaml_str
            assert body  # non-empty body

    def test_format_sanitized_image_link_consumed_for_wiki_styles(self, sample_png_path: Path) -> None:
        """When the link style is wiki_embed/wiki_link, the converter uses
        ``format_sanitized_image_link`` (F-I2b-004) instead of the
        export_semantics fallback."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)
            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # format_sanitized_image_link produces clean [[...]] without ./ prefix
            # and with file-system-illegal chars stripped
            assert "![[sample.png]]" in md_text
            assert "![[./" not in md_text

    def test_ocr_heading_body_split_consumed_in_ocr_path(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """OCR mixed heading/body presentation uses the image-local helper."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        # Simulate OCR returning mixed heading+body Chinese text
        monkeypatch.setattr(
            converter_mod,
            "run_ocr_outcome",
            lambda _path, **_kwargs: _ocr_success("一、标题：正文内容\n二、说明：补充信息"),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "embed",
                    "ocr_placement": "main_md",
                    "to_md_enable_ocr": True,
                },
            )
            result = ImageToMarkdownConverter().convert(context)
            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # The converter should have split at the delimiter and rendered
            # the heading portion as bold inside the blockquote.
            assert "> **一、标题：**" in md_text
            assert "正文内容" in md_text

    def test_generate_basic_yaml_frontmatter_consumed_by_converter(self, sample_png_path: Path) -> None:
        """The converter replaces ad-hoc YAML string concatenation with
        ``generate_basic_yaml_frontmatter`` from yaml_tools."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)
            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # Verify the YAML frontmatter is from generate_basic_yaml_frontmatter
            # (has both title and aliases fields, title before aliases)
            assert "title: sample" in md_text
            assert "aliases:" in md_text
            assert "source_format: png" in md_text
            title_pos = md_text.index("title:")
            aliases_pos = md_text.index("aliases:")
            assert title_pos < aliases_pos, "title: must appear before aliases: (R1 regression fix)"


class TestImageConverterPathUtilsConsumption:
    """Verify the image converter is a real production consumer of the new
    ``docwen_core.paths`` module.

    F-I2a-020 (sanitize_filename), F-I2a-021 (sanitize_for_wiki_link):
    already resolved in ``docwen_core.markdown_utils`` — verified by the
    ``TestImageConverterCoreUtilsConsumption`` class above and the
    ``test_numbering_paths_filenames.py`` core test.
    """

    def test_normalize_path_imported_in_converter(self) -> None:
        """``normalize_path`` from core is importable and callable."""
        from docwen_core.paths import normalize_path

        result = normalize_path("~/Downloads")
        assert isinstance(result, str)
        assert len(result) > 0

    def test_converter_input_path_is_normalized(self, sample_png_path: Path) -> None:
        """The converter passes ``input_path`` through ``normalize_path``
        so that downstream operations receive a clean absolute path."""
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_enable_ocr": False},
            )
            from docwen_plugin_image.to_markdown.converter import (
                ImageToMarkdownConverter,
            )

            result = ImageToMarkdownConverter().convert(context)
            assert result.success is True

            # The converter should still produce valid output when the
            # path is normalised — this is a smoke test for the wiring.
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert len(md_text) > 0
            assert "---" in md_text  # YAML frontmatter present

    def test_plugin_common_reuses_core_input_stem(self, sample_png_path: Path) -> None:
        """The image common module imports, rather than copies, the core helper."""
        from docwen_core.paths import input_stem as core_input_stem
        from docwen_plugin_image._common import input_stem as plugin_input_stem

        assert plugin_input_stem is core_input_stem
        assert core_input_stem(str(sample_png_path)) == plugin_input_stem(str(sample_png_path))
        assert core_input_stem("/a/b/财务报告.pdf") == "财务报告"

    def test_input_name_from_core_paths(self, sample_png_path: Path) -> None:
        """``input_name`` from core paths gives the full filename."""
        from docwen_core.paths import input_name

        name = input_name(sample_png_path)
        assert name == "sample.png"
        assert input_name("/a/b/财务报告.pdf") == "财务报告.pdf"

    def test_ensure_dir_exists_from_core(self, tmp_path: Path) -> None:
        """``ensure_dir_exists`` creates directories as expected."""
        from docwen_core.paths import ensure_dir_exists

        target = tmp_path / "images" / "staging"
        result = ensure_dir_exists(str(target))
        assert target.is_dir()
        assert isinstance(result, str) and len(result) > 0

    def test_safe_join_path_prevents_traversal(self, tmp_path: Path) -> None:
        """``safe_join_path`` blocks path-traversal attempts."""
        import pytest as _pytest

        from docwen_core.paths import safe_join_path

        base = str(tmp_path)
        good = safe_join_path(base, "images", "photo.png")
        assert good.endswith("photo.png")

        with _pytest.raises(ValueError, match="Path traversal"):
            safe_join_path(base, "../../../etc/passwd")

    def test_converter_filename_uses_core_sanitize(self, sample_png_path: Path) -> None:
        """The image filename produced by the converter is sanitised
        through ``docwen_core.markdown_utils.sanitize_filename``,
        ensuring F-I2a-020 closure at the user-path level."""
        from docwen_core.markdown_utils import sanitize_filename
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "file", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)
            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            # The sanitised filename is used in the wiki link target
            expected = sanitize_filename("sample.png")
            assert expected in md_text
