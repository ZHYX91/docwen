"""Focused tests split from test_markdown_assets.py."""

from __future__ import annotations

from ._markdown_assets_support import (
    Path,
    _build_fake_context,
    _ocr_success,
    base64,
    format_image_placeholder,
    is_data_uri_image,
    pytest,
    re,
    tempfile,
)

pytestmark = [pytest.mark.golden, pytest.mark.contract]


class TestImageToMarkdownEmbedMode:
    """Verify the image converter produces ``{{IMAGE:...}}`` placeholders
    when ``image_mode == "embed"``."""

    def test_embed_mode_produces_placeholder(self, sample_png_path: Path) -> None:
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "embed", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1  # only MD, no image copy
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "{{IMAGE:" in md_text
            assert "sample.png" in md_text
            assert result.metrics.output_bytes > 0

    def test_embed_mode_placeholder_uses_shared_format_image_placeholder(self, sample_png_path: Path) -> None:
        """The placeholder text produced by the converter must match what
        ``format_image_placeholder`` returns for the same input."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "embed", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")

            expected_placeholder = format_image_placeholder("./sample.png")
            assert expected_placeholder in md_text

    def test_embed_mode_gif_image(self, sample_gif_path: Path) -> None:
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_gif_path),
                staging,
                "md",
                {"image_mode": "embed", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            expected = format_image_placeholder("./sample.gif")
            assert expected in md_text

    def test_embed_mode_jpg_image(self, sample_jpg_path: Path) -> None:
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_jpg_path),
                staging,
                "md",
                {"image_mode": "embed", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            expected = format_image_placeholder("./sample.jpg")
            assert expected in md_text

    def test_embed_mode_yaml_frontmatter_present(self, sample_png_path: Path) -> None:
        """Embed mode still produces YAML frontmatter with image metadata."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "embed", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "---" in md_text
            assert "title: sample" in md_text
            assert "source_format: png" in md_text

    def test_embed_mode_consumes_locale_yaml_title_label(self, sample_png_path: Path) -> None:
        """Image→Markdown consumes pre-resolved YAML labels from the app edge."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "embed",
                    "to_md_enable_ocr": False,
                    "yaml_key_labels": {"title": "Titel"},
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "Titel: sample" in md_text
            assert "title: sample" not in md_text
            assert "source_format: png" in md_text

    def test_embed_mode_no_image_artifact_registered(self, sample_png_path: Path) -> None:
        """Embed mode produces only the primary MD artifact — no image copy."""
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "embed", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            assert result.artifacts[0].kind == "primary"
            assert result.artifacts[0].is_primary is True

    def test_embed_mode_with_ocr(self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """Embed mode combined with OCR appends blockquote after placeholder."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", lambda _path, **_kwargs: _ocr_success("识别文本"))

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
            assert "{{IMAGE:" in md_text
            assert "> 识别文本" in md_text
            assert any(d.code == "IMG2MD-OCR-OK" for d in result.diagnostics)

    def test_embed_mode_metrics_do_not_count_image_artifact(self, sample_png_path: Path) -> None:
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {"image_mode": "embed", "to_md_enable_ocr": False},
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert result.metrics.extra["artifact_count"] == 1
            assert result.metrics.output_bytes > 0


class TestImageToMarkdownBase64IsDataUri:
    """Verify the base64 output of the image converter is recognised by the
    shared ``is_data_uri_image`` utility (F-H2-006 integration)."""

    def test_base64_output_is_detected_as_data_uri_image(self, sample_png_path: Path) -> None:
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

            # Extract the data URI from the wiki embed link (![[data:...]])
            import re

            m = re.search(r"!\[\[(data:image/[^\]]+)\]\]", md_text)
            assert m is not None
            data_uri = m.group(1)

            assert is_data_uri_image(data_uri) is True

    def test_base64_output_consumes_compression_semantics(self, tmp_path: Path) -> None:
        from PIL import Image

        from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

        source = tmp_path / "noise.png"
        Image.effect_noise((512, 512), 100).convert("RGB").save(source, format="PNG")
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(source),
                staging,
                "md",
                {"image_mode": "base64", "to_md_enable_ocr": False},
                config_snapshot={
                    "conversion": {
                        "export": {
                            "base64_compress_enabled": True,
                            "base64_compress_threshold_kb": 100,
                        }
                    }
                },
            )
            result = ImageToMarkdownConverter().convert(context)
            assert result.success is True
            markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")

        match = re.search(r"data:([^;]+);base64,([A-Za-z0-9+/=]+)", markdown)
        assert match is not None
        payload = base64.b64decode(match.group(2))
        assert match.group(1) == "image/jpeg"
        assert payload.startswith(b"\xff\xd8\xff")
        assert len(payload) <= 100 * 1024
        assert len(payload) < source.stat().st_size


class TestImageToMarkdownEmbedGuard:
    """The logic ``(not keep_images and ... != "embed")`` must not treat
    embed mode as omit."""

    def test_embed_mode_with_keep_images_false_still_embeds(self, sample_png_path: Path) -> None:
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "embed",
                    "to_md_keep_images": False,
                    "to_md_enable_ocr": False,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            md_text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "{{IMAGE:" in md_text
            # Must NOT be an omit comment
            assert "<!-- image omitted:" not in md_text


def test_nonempty_snapshot_owns_image_links_and_ocr_title(
    sample_png_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """One admitted image request owns its complete export policy."""
    import docwen_plugin_image.to_markdown.converter as converter_mod
    from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter

    snapshot = {
        "gui": {"language": {"locale": "zh_CN"}},
        "link": {
            "format": {
                "image_link_style": "markdown_embed",
                "md_file_link_style": "markdown_link",
            }
        },
        "conversion": {
            "ocr_output": {
                "show_blockquote_title": True,
                "blockquote_title_override_by_locale": {"zh_CN": "SNAPSHOT TITLE"},
            }
        },
    }
    monkeypatch.setattr(converter_mod, "run_ocr_outcome", lambda _path, **_kwargs: _ocr_success("owned text"))
    with tempfile.TemporaryDirectory() as staging:
        context = _build_fake_context(
            str(sample_png_path),
            staging,
            "md",
            {
                "image_mode": "file",
                "to_md_enable_ocr": True,
                "to_md_keep_images": True,
            },
            config_snapshot=snapshot,
            ocr_blockquote_title="SNAPSHOT TITLE",
        )
        result = ImageToMarkdownConverter().convert(context)
        assert result.success is True
        primary_md = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")

    assert "![sample](./sample.png)" in primary_md
    assert "> **SNAPSHOT TITLE**" in primary_md


class TestImageConverterErrorSemanticsIntegration:
    """Verify the image→Markdown converter correctly consumes the shared
    ``docwen_core.links`` error-semantics infrastructure.

    F-H2-010 ─ F-H2-014: error placeholder / keep-link output must be
    stable and diagnostic so users can see what went wrong.
    """

    def test_embed_mode_error_path_placeholder_importable(self) -> None:
        """Core error-semantics symbols are importable from the image
        plugin test context — proving the shared module is consumable."""
        from docwen_core.links import (
            LinkErrorKind,
            make_error_placeholder,
        )

        # Smoke: every error kind produces non-empty placeholder text
        for kind in LinkErrorKind:
            text = make_error_placeholder(kind, "test.png")
            assert len(text) > 0
            assert "test.png" in text

    def test_dispatch_error_output_all_modes_produce_stable_output(self) -> None:
        """Every mode (ignore / keep / placeholder) returns a deterministic
        string for each error kind."""
        from docwen_core.links import LinkErrorKind, dispatch_error_output

        for kind in LinkErrorKind:
            ignore = dispatch_error_output(kind, "ignore", "f.png")
            keep = dispatch_error_output(kind, "keep", "f.png")
            placeholder = dispatch_error_output(kind, "placeholder", "f.png")

            assert ignore == ""
            assert "![[f.png]]" in keep
            assert placeholder.startswith("[")
            assert "f.png" in placeholder

    def test_error_semantics_max_depth_not_silently_swallowed(self, tmp_path: Path) -> None:
        """Max-depth errors are no longer silently ignored — the
        placeholder default ensures diagnostic output."""

        from docwen_core.links import resolve_embedded_links

        # Write a chain that exceeds max_depth
        vault = tmp_path / "vault"
        vault.mkdir(parents=True, exist_ok=True)
        src = vault / "src.md"
        src.write_text("![[a.md]]\n", encoding="utf-8")
        (vault / "a.md").write_text("![[b.md]]\n", encoding="utf-8")
        (vault / "b.md").write_text("![[c.md]]\n", encoding="utf-8")
        (vault / "c.md").write_text("Deep.\n", encoding="utf-8")

        # Default behavior: placeholder (diagnostic, not silent)
        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=1,
        )
        assert "[Max depth reached:" in result

        # User can opt into ignore (silent truncation)
        result_ignore = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            max_depth=1,
            on_max_depth="ignore",
        )
        assert "[Max depth reached:" not in result_ignore

    def test_error_semantics_file_not_found_not_silently_swallowed(
        self,
        tmp_path: Path,
    ) -> None:
        """File-not-found errors produce diagnostic placeholders by default."""
        from docwen_core.links import resolve_embedded_links

        src = tmp_path / "vault" / "main.md"
        src.parent.mkdir(parents=True, exist_ok=True)
        src.write_text(".\n", encoding="utf-8")

        result = resolve_embedded_links(
            "![[ghost.png]]",
            str(src),
            on_not_found="placeholder",
        )
        assert "[File not found: ghost.png]" in result
        assert "![[ghost.png]]" not in result

    def test_error_semantics_circular_not_silently_swallowed(
        self,
        tmp_path: Path,
    ) -> None:
        """Circular references produce diagnostic placeholders by default."""
        from docwen_core.links import resolve_embedded_links

        src = tmp_path / "vault" / "a.md"
        src.parent.mkdir(parents=True, exist_ok=True)
        src.write_text("![[b.md]]\n", encoding="utf-8")
        (tmp_path / "vault" / "b.md").write_text("![[a.md]]\n", encoding="utf-8")

        result = resolve_embedded_links(
            "![[a.md]]",
            str(src),
            on_circular="placeholder",
        )
        assert "[Circular reference:" in result
