"""Focused tests split from test_runtime_config_wiring.py."""

from __future__ import annotations

from ._runtime_config_wiring_support import (
    PROJECT_CONFIGS,
    Path,
    pytest,
    tempfile,
    tomllib,
)

pytestmark = pytest.mark.unit


class TestConfigSnapshotProjectsExportSemantics:
    """Verify loaded snapshots project deterministic export semantics."""

    def test_reload_exposes_request_snapshot_for_projection(self) -> None:
        from docwen_core.export_semantics import MarkdownExportSemantics
        from docwen_runtime.config.loader import ConfigLoader

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            semantics = MarkdownExportSemantics.from_config_snapshot(loader.config.as_dict())
            assert semantics.image_link_style == "wiki_embed"
            assert semantics.yaml_list_separator == "、"

    def test_config_value_change_updates_request_projection(self) -> None:
        from docwen_core.export_semantics import MarkdownExportSemantics
        from docwen_runtime.config.loader import ConfigLoader

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))

            loader.set_value("link.format.image_link_style", "markdown_embed")
            semantics = MarkdownExportSemantics.from_config_snapshot(loader.config.as_dict())
            assert semantics.image_link_style == "markdown_embed"

    @pytest.mark.parametrize("separator", [", ", ""])
    def test_snapshot_projection_consumes_nested_yaml_list_separator_override(
        self,
        tmp_path: Path,
        separator: str,
    ) -> None:
        from docwen_core.export_semantics import MarkdownExportSemantics
        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "user-config")
        assert loader.set_value("conversion.md_to_docx.list_separator", separator)

        semantics = MarkdownExportSemantics.from_config_snapshot(loader.config.as_dict())
        assert semantics.yaml_list_separator == separator

    def test_export_semantics_config_values_match_loader_defaults(self) -> None:
        """Projected semantics must match the base file values."""
        from docwen_core.export_semantics import MarkdownExportSemantics
        from docwen_runtime.config.loader import ConfigLoader

        base_export = tomllib.loads((PROJECT_CONFIGS / "conversion.toml").read_text(encoding="utf-8"))["export"]

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            semantics = MarkdownExportSemantics.from_config_snapshot(loader.config.as_dict())

            assert semantics.export_base64_compress_enabled == base_export["base64_compress_enabled"]
            assert semantics.export_base64_compress_threshold_kb == base_export["base64_compress_threshold_kb"]

    def test_export_semantics_consumes_conversion_export_user_override(self, tmp_path: Path) -> None:
        """Settings-persisted overrides reach request-owned semantics."""
        from docwen_core.export_semantics import MarkdownExportSemantics
        from docwen_runtime.config.loader import ConfigLoader

        user_dir = tmp_path / "user-config"
        loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
        assert loader.set_values(
            {
                "conversion.export.base64_compress_enabled": False,
                "conversion.export.base64_compress_threshold_kb": 512,
            }
        )

        semantics = MarkdownExportSemantics.from_config_snapshot(loader.config.as_dict())

        assert semantics.export_base64_compress_enabled is False
        assert semantics.export_base64_compress_threshold_kb == 512

    def test_export_markdown_modes_match_export_toml_defaults(self) -> None:
        """Projected markdown mode defaults must come from export.toml."""
        from docwen_core.export_semantics import MarkdownExportSemantics, get_markdown_export_modes
        from docwen_runtime.config.loader import ConfigLoader

        export_defaults = tomllib.loads((PROJECT_CONFIGS / "export.toml").read_text(encoding="utf-8"))

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            semantics = MarkdownExportSemantics.from_config_snapshot(loader.config.as_dict())
            modes = get_markdown_export_modes("image", semantics=semantics)

            assert semantics.image_extraction_mode == export_defaults["to_md_image_extraction_mode"]
            assert semantics.ocr_placement_mode == export_defaults["to_md_ocr_placement_mode"]
            assert modes["image_extraction_mode"] == export_defaults["to_md_image_extraction_mode"]
            assert modes["ocr_placement_mode"] == export_defaults["to_md_ocr_placement_mode"]

    def test_link_runtime_config_values_match_loader_defaults(self) -> None:
        """Request-scoped link config projection must match loader defaults."""
        from docwen_core.export_semantics import LinkRuntimeConfig
        from docwen_runtime.config.loader import ConfigLoader

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            link_cfg = LinkRuntimeConfig.from_config(loader.config.as_dict()["link"])

            assert link_cfg.max_depth == 3
            assert isinstance(link_cfg.search_dirs, tuple)
            assert "." in link_cfg.search_dirs
            assert link_cfg.detect_circular is True
            assert link_cfg.file_not_found_mode == "placeholder"

    def test_docx_style_detector_config_matches_document_toml_defaults(self) -> None:
        """Document style aliases project from the loader snapshot."""
        from docwen_core.docx_parsing.format_features import (
            detect_paragraph_style_type,
            style_detector_config_from_document_config,
        )
        from docwen_runtime.config.loader import ConfigLoader

        class _Style:
            name = "HTML Preformatted"

        class _Paragraph:
            style = _Style()

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            snapshot = loader.config.as_dict()
            config = style_detector_config_from_document_config(snapshot["document"])

            assert config is not None
            assert detect_paragraph_style_type(_Paragraph(), config=config) == ("code_block", True)

    def test_docx_markdown_formatting_config_matches_conversion_toml_defaults(self) -> None:
        """DOCX->MD formatting switches project from the loader snapshot."""
        from docwen_core.docx_parsing.format_features import (
            docx_markdown_formatting_config_from_conversion_config,
        )
        from docwen_runtime.config.loader import ConfigLoader

        conversion_defaults = tomllib.loads((PROJECT_CONFIGS / "conversion.toml").read_text(encoding="utf-8"))
        docx_to_md = conversion_defaults["docx_to_md"]

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            snapshot = loader.config.as_dict()
            config = docx_markdown_formatting_config_from_conversion_config(snapshot["conversion"])

            assert config.preserve_formatting == docx_to_md["preserve_formatting"]
            assert config.preserve_heading_formatting == docx_to_md["preserve_heading_formatting"]
            assert config.preserve_table_header_formatting == docx_to_md["preserve_table_header_formatting"]

    def test_docx_markdown_syntax_config_matches_conversion_toml_defaults(self) -> None:
        """DOCX->MD inline syntax choices project from the loader snapshot."""
        from docwen_core.docx_parsing.format_features import (
            docx_markdown_syntax_config_from_conversion_config,
        )
        from docwen_runtime.config.loader import ConfigLoader

        conversion_defaults = tomllib.loads((PROJECT_CONFIGS / "conversion.toml").read_text(encoding="utf-8"))
        syntax = conversion_defaults["syntax"]

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            snapshot = loader.config.as_dict()
            config = docx_markdown_syntax_config_from_conversion_config(snapshot["conversion"])

            assert config.bold == syntax["bold"]
            assert config.italic == syntax["italic"]
            assert config.strikethrough == syntax["strikethrough"]
            assert config.highlight == syntax["highlight"]
            assert config.superscript == syntax["superscript"]
            assert config.subscript == syntax["subscript"]
            assert config.unordered_list == syntax["unordered_list"]
            assert config.indent_spaces == syntax["indent_spaces"]
