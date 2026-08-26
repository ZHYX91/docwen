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


class TestImageToMarkdownHrefNormalization:
    """Verify the image→Markdown converter's href normalisation reuses the
    shared ``docwen_core.links.normalize_link_target`` (F-G2-004).

    The old ``_normalize_href`` performed four operations:
    1. URL-decode (``%20`` → `` ``)
    2. Replace backslashes (``\\`` → ``/``)
    3. Strip leading slashes
    4. (optionally) extract path from full URL via ``urlparse``

    The shared ``normalize_link_target`` now covers steps 1–3.  These
    tests prove the normalisation chain works end-to-end through the
    image converter.
    """

    # ── Core unit: normalize_link_target ──────────────────────────────

    def test_normalize_backslashes_to_forward_slashes(self) -> None:
        """Windows-style backslash paths become forward-slash paths."""
        from docwen_core.links import normalize_link_target

        assert normalize_link_target("images\\photo.png") == "images/photo.png"
        assert normalize_link_target("..\\..\\assets\\img.jpg") == "../../assets/img.jpg"

    def test_normalize_strips_leading_slashes(self) -> None:
        """Leading slashes are stripped so the output is relative."""
        from docwen_core.links import normalize_link_target

        assert normalize_link_target("/images/photo.png") == "images/photo.png"
        assert normalize_link_target("///absolute/path/img.jpg") == "absolute/path/img.jpg"

    def test_normalize_percent_encoded_with_backslashes(self) -> None:
        """Percent-encoding and backslashes are both resolved."""
        from docwen_core.links import normalize_link_target

        result = normalize_link_target("my%20docs\\photo%201.png")
        assert result == "my docs/photo 1.png"

    def test_normalize_combined_leading_slash_and_backslash(self) -> None:
        """Leading slash stripping + backslash normalisation combined."""
        from docwen_core.links import normalize_link_target

        result = normalize_link_target("/network\\share\\data.png")
        assert result == "network/share/data.png"

    def test_normalize_anchor_with_backslash(self) -> None:
        """Anchor stripping still works with backslash paths."""
        from docwen_core.links import normalize_link_target

        result = normalize_link_target("docs\\readme.md#section-1")
        assert result == "docs/readme.md"
        assert "#" not in result

    def test_normalize_query_with_leading_slash(self) -> None:
        """Query stripping still works with leading-slash paths."""
        from docwen_core.links import normalize_link_target

        result = normalize_link_target("/images/photo.png?v=2")
        assert result == "images/photo.png"
        assert "?" not in result

    def test_normalize_preserves_dot_slash_prefix(self) -> None:
        """``./`` prefixed paths keep their relative prefix (no leading ``/`` to strip)."""
        from docwen_core.links import normalize_link_target

        assert normalize_link_target("./sample.png") == "./sample.png"
        assert normalize_link_target("./subdir/file.md") == "./subdir/file.md"

    # ── Integration: image converter uses shared normalisation ────────

    def test_converter_uses_shared_normalize_link_target(self, sample_png_path: Path) -> None:
        """The image converter imports ``normalize_link_target`` from
        ``docwen_core.links``, not from a plugin-private copy."""
        import ast
        from pathlib import Path as P

        converter_file = P(__file__).parent.parent / "src" / "docwen_plugin_image" / "to_markdown" / "converter.py"
        tree = ast.parse(converter_file.read_text(encoding="utf-8"))

        # Check that normalize_link_target is imported from docwen_core.links
        imports_from_core_links = []
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module == "docwen_core.links":
                for alias in node.names:
                    imports_from_core_links.append(alias.name)

        assert "normalize_link_target" in imports_from_core_links, (
            f"Converter must import normalize_link_target from docwen_core.links; "
            f"found core links imports: {imports_from_core_links}"
        )

    def test_converter_file_mode_produces_normalized_link(self, sample_png_path: Path) -> None:
        """In ``file`` mode with markdown_embed style, the output link
        target is normalised through ``normalize_link_target`` (F-G2-004)."""
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
            # markdown_embed path goes through normalize_link_target('./sample.png')
            assert "![sample](./sample.png)" in md_text
            # No backslashes, no percent-encoding, no leading slashes in target
            assert "\\" not in md_text[md_text.index("![sample](") :]
            assert "%" not in md_text[md_text.index("![sample](") :]
            assert "//sample" not in md_text

    def test_converter_base64_mode_preserves_data_uri(self, sample_png_path: Path) -> None:
        """Base64 data-URI targets are NOT mangled by normalisation —
        they are passed directly to ``format_image_link``."""
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
            # The data URI base64 payload must survive intact
            assert "data:image/png;base64," in md_text
            # No double-encoding or decoding
            assert "data:image/png%3Bbase64" not in md_text

    def test_normalize_imported_from_core_not_redefined(self) -> None:
        """The image converter MUST NOT define its own ``_normalize_href``
        or similar private copy — it imports from ``docwen_core.links``."""
        import ast
        from pathlib import Path as P

        converter_file = P(__file__).parent.parent / "src" / "docwen_plugin_image" / "to_markdown" / "converter.py"
        tree = ast.parse(converter_file.read_text(encoding="utf-8"))

        # No private normalisation redefinition
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                name = node.name
                assert "normalize" not in name.lower() or name == "normalize_format", (
                    f"Converter must not define its own normalisation: {name}"
                )
                assert "href" not in name.lower(), f"Converter must not define _normalize_href: {name}"

    def test_normalize_link_target_exported_from_core_init(self) -> None:
        """``normalize_link_target`` is part of the public API of
        ``docwen_core.links``."""
        from docwen_core.links import normalize_link_target as nlt

        # The function must be the same object regardless of import path
        from docwen_core.links._resolver import (
            normalize_link_target as nlt_direct,
        )

        assert nlt is nlt_direct


class TestImageToMarkdownOcrPlacementImageMd:
    """F-G1-005, F-G2-003: Verify the ``image_md`` OCR placement mode
    creates a separate per-image .md file and links to it from the
    primary output instead of appending OCR inline."""

    def test_resource_writeback_image_md_creates_auxiliary_and_primary_md(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """With ocr_placement_mode=image_md, the converter produces:
        - A primary .md containing a .md file link
        - An auxiliary .md containing the image link and OCR text."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        monkeypatch.setattr(
            converter_mod,
            "run_ocr_outcome",
            lambda _path, **_kwargs: _ocr_success("第1页：正文内容"),
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "file",
                    "ocr_placement": "image_md",
                    "to_md_enable_ocr": True,
                    "to_md_keep_images": True,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert len(result.artifacts) == 3  # primary.md + image + auxiliary_ocr.md
            artifacts_by_kind = {a.kind: a for a in result.artifacts}

            # Primary output: .md file link (not inline OCR)
            assert "primary" in artifacts_by_kind
            primary_md = Path(artifacts_by_kind["primary"].staging_path).read_text(encoding="utf-8")
            assert "![[sample_ocr.md]]" in primary_md
            assert "第1页" not in primary_md  # OCR NOT inline

            # Auxiliary output: image link + OCR text
            assert "auxiliary" in artifacts_by_kind
            aux_md = Path(artifacts_by_kind["auxiliary"].staging_path).read_text(encoding="utf-8")
            assert "![[sample.png]]" in aux_md  # image link preserved
            assert "第1页" in aux_md  # OCR text present

            # Image artifact still registered
            assert "image" in artifacts_by_kind

    def test_image_md_placement_empty_ocr_still_creates_auxiliary(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Legacy image_md keeps the companion resource even when OCR finds no text."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", lambda _path, **_kwargs: _ocr_success(""))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "file",
                    "ocr_placement": "image_md",
                    "to_md_enable_ocr": True,
                    "to_md_keep_images": True,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            assert [artifact.kind for artifact in result.artifacts] == ["primary", "image", "auxiliary"]
            primary = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            sidecar = Path(result.artifacts[2].staging_path).read_text(encoding="utf-8")
            assert "![[sample_ocr.md]]" in primary
            assert "![[sample.png]]" in sidecar
            assert ">" not in sidecar
            assert result.metrics.extra["ocr_chars"] == 0

    def test_image_md_placement_keep_images_false_still_creates_auxiliary(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """When keep_images=False and ocr_placement_mode=image_md, the
        auxiliary .md contains OCR text (no image link) and the primary
        links to it."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", lambda _path, **_kwargs: _ocr_success("识别结果"))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "file",
                    "ocr_placement": "image_md",
                    "to_md_enable_ocr": True,
                    "to_md_keep_images": False,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            # primary.md + auxiliary_ocr.md  (no image artifact — keep_images=False)
            assert len(result.artifacts) == 2
            artifacts_by_kind = {a.kind: a for a in result.artifacts}

            primary_md = Path(artifacts_by_kind["primary"].staging_path).read_text(encoding="utf-8")
            assert "![[sample_ocr.md]]" in primary_md

            aux_md = Path(artifacts_by_kind["auxiliary"].staging_path).read_text(encoding="utf-8")
            assert "<!-- image omitted:" in aux_md
            assert "识别结果" in aux_md

    def test_image_md_placement_embed_mode_with_ocr(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """image_md placement + embed image_mode: the auxiliary .md carries
        the embed placeholder; the primary links to it."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", lambda _path, **_kwargs: _ocr_success("OCR嵌入模式"))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "embed",
                    "ocr_placement": "image_md",
                    "to_md_enable_ocr": True,
                    "to_md_keep_images": True,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            # primary.md + auxiliary_ocr.md (embed mode does NOT register image artifact)
            assert len(result.artifacts) == 2
            artifacts_by_kind = {a.kind: a for a in result.artifacts}

            primary_md = Path(artifacts_by_kind["primary"].staging_path).read_text(encoding="utf-8")
            assert "![[sample_ocr.md]]" in primary_md
            assert "{{IMAGE:" not in primary_md  # placeholder NOT inline

            aux_md = Path(artifacts_by_kind["auxiliary"].staging_path).read_text(encoding="utf-8")
            assert "{{IMAGE:" in aux_md
            assert "OCR嵌入模式" in aux_md

    def test_main_md_placement_still_appends_ocr_inline(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """With ocr_placement_mode=main_md (the default), OCR text is
        appended as a blockquote — no auxiliary .md is created."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", lambda _path, **_kwargs: _ocr_success("主文档OCR"))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "file",
                    "ocr_placement": "main_md",
                    "to_md_enable_ocr": True,
                    "to_md_keep_images": True,
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            # primary.md + image artifact (no auxiliary)
            assert len(result.artifacts) == 2
            kinds = {a.kind for a in result.artifacts}
            assert kinds == {"primary", "image"}

            primary_md = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert "> 主文档OCR" in primary_md  # OCR inline
            assert "![[sample.png]]" in primary_md

    def test_image_md_placement_auxiliary_has_yaml_frontmatter(
        self, sample_png_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """The auxiliary .md generated by image_md mode has its own YAML
        frontmatter with OCR metadata."""
        import docwen_plugin_image.to_markdown.converter as converter_mod
        from docwen_plugin_image.to_markdown.converter import (
            ImageToMarkdownConverter,
        )

        monkeypatch.setattr(converter_mod, "run_ocr_outcome", lambda _path, **_kwargs: _ocr_success("OCR文本"))

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_png_path),
                staging,
                "md",
                {
                    "image_mode": "file",
                    "ocr_placement": "image_md",
                    "to_md_enable_ocr": True,
                    "to_md_keep_images": True,
                    "yaml_key_labels": {"title": "Titel"},
                },
            )
            result = ImageToMarkdownConverter().convert(context)

            assert result.success is True
            aux = next(a for a in result.artifacts if a.kind == "auxiliary")
            aux_md = Path(aux.staging_path).read_text(encoding="utf-8")
            assert "---" in aux_md
            assert "Titel: sample_ocr" in aux_md
            assert "title: sample_ocr" not in aux_md
            assert "ocr: True" in aux_md  # YAML serialises Python True as "True"
