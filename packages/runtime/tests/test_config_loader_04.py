"""Focused tests split from test_config_loader.py."""

from __future__ import annotations

from ._config_loader_support import (
    PROJECT_CONFIGS,
    Path,
    pytest,
    write_minimal_base_config_tree,
)

pytestmark = pytest.mark.unit


def test_registry_reset_plans_model_cross_file_logical_ownership() -> None:
    from docwen_runtime.config.loader import RESET_EXCLUDED
    from docwen_runtime.config.registry import reset_plan_for_group, spec_for_dotted_key

    general_plan = reset_plan_for_group("general")
    assert general_plan.files == ()
    assert set(general_plan.dotted_keys) == {
        "gui.theme",
        "gui.window",
        "gui.transparency",
        "gui.language",
    }

    text_plan = reset_plan_for_group("text")
    assert text_plan.files == ()
    assert set(text_plan.dotted_keys) == {
        "text.remove_numbering",
        "text.add_numbering",
        "text.numbering_scheme",
        "text.heading_numbering_render_mode",
        "gui.template.md_default_template",
        "numbering.add.settings.default_scheme",
    }

    export_plan = reset_plan_for_group("export")
    assert export_plan.files == ("export.toml",)
    assert set(export_plan.dotted_keys) == {
        "conversion.ocr_output.show_blockquote_title",
        "conversion.ocr_output.blockquote_title_override_by_locale",
        "conversion.export.base64_compress_enabled",
        "conversion.export.base64_compress_threshold_kb",
    }

    formatting_plan = reset_plan_for_group("formatting")
    assert formatting_plan.files == ()
    assert len(formatting_plan.dotted_keys) == 26
    assert {
        "conversion.md_to_docx.heading_merge_mode",
        "conversion.md_to_docx.heading_merge_punctuation",
        "conversion.md_to_docx.list_separator",
        "document.style.table.md_to_docx.table_style_mode",
        "document.style.table.md_to_docx.builtin_style_key",
        "document.style.table.md_to_docx.custom_style_name",
    }.issubset(formatting_plan.dotted_keys)
    assert {
        "conversion.horizontal_rule.enabled",
        "conversion.code_detection.code_font",
        "conversion.export.base64_compress_enabled",
        "conversion.ocr_output.show_blockquote_title",
    }.isdisjoint(formatting_plan.dotted_keys)

    document_plan = reset_plan_for_group("document")
    assert document_plan.files == ()
    assert set(document_plan.dotted_keys) == {
        "document.to_md_keep_images",
        "document.to_md_enable_ocr",
        "document.to_md_table_merge_export_strategy",
        "document.to_md_remove_numbering",
        "document.to_md_add_numbering",
        "document.to_md_default_scheme",
        "document.to_md_enable_optimization",
        "document.to_md_optimization_type",
        "software.default_priority.word_processors",
        "software.special_conversions.odt",
        "software.special_conversions.document_to_pdf",
    }

    spreadsheet_plan = reset_plan_for_group("spreadsheet")
    assert spreadsheet_plan.files == ()
    assert set(spreadsheet_plan.dotted_keys) == {
        "spreadsheet.to_md_keep_images",
        "spreadsheet.to_md_enable_ocr",
        "spreadsheet.to_md_table_merge_export_strategy",
        "spreadsheet.merge_mode",
        "software.default_priority.spreadsheet_processors",
        "software.special_conversions.ods",
        "software.special_conversions.spreadsheet_to_pdf",
    }

    layout_plan = reset_plan_for_group("layout")
    assert layout_plan.files == ()
    assert set(layout_plan.dotted_keys) == {
        "layout.to_md_keep_images",
        "layout.to_md_enable_ocr",
        "layout.to_md_enable_optimization",
        "layout.to_md_optimization_type",
        "layout.render_dpi",
        "software.special_conversions.pdf_to_office",
    }

    link_plan = reset_plan_for_group("link")
    assert link_plan.files == ()
    assert set(link_plan.dotted_keys) == {
        "link.format.image_link_style",
        "link.format.md_file_link_style",
        "link.non_embed_links.wiki_mode",
        "link.non_embed_links.markdown_mode",
        "link.embed_links.wiki_image_mode",
        "link.embed_links.markdown_image_mode",
        "link.embed_links.md_file_mode",
        "link.embedding.max_depth",
    }

    other_plan = reset_plan_for_group("other")
    assert other_plan.files == ()
    assert set(other_plan.dotted_keys) == {
        "other.to_md_keep_images",
        "other.to_md_enable_ocr",
    }

    output_plan = reset_plan_for_group("output")
    assert output_plan.files == ()
    assert set(output_plan.dotted_keys) == {
        "output.intermediate_files.save_to_output",
        "output.directory.mode",
        "output.directory.custom_path",
        "output.directory.create_date_subfolder",
        "output.directory.date_folder_format",
        "output.behavior.auto_open_folder",
    }

    logging_plan = reset_plan_for_group("logging")
    assert logging_plan.files == ()
    assert set(logging_plan.dotted_keys) == {
        "logger.enable",
        "logger.level",
        "logger.file_prefix",
        "logger.retention_days",
        "logger.console_enable",
        "logger.console_level",
        "logger.console_format",
        "logger.console_colorize",
        "logger.directory_mode",
        "logger.directory",
    }

    proofread_plan = reset_plan_for_group("proofread")
    assert proofread_plan.files == ()
    assert set(proofread_plan.dotted_keys) == {
        "proofread.engine.enable_symbol_pairing",
        "proofread.engine.enable_symbol_correction",
        "proofread.engine.enable_typos_rule",
        "proofread.engine.enable_sensitive_word",
        "proofread.skip.code_blocks",
        "proofread.skip.quote_blocks",
    }

    for plan in (
        general_plan,
        text_plan,
        export_plan,
        formatting_plan,
        document_plan,
        spreadsheet_plan,
        layout_plan,
        link_plan,
        other_plan,
        output_plan,
        logging_plan,
        proofread_plan,
    ):
        assert len(plan.dotted_keys) == len(set(plan.dotted_keys))
        for key in plan.dotted_keys:
            spec = spec_for_dotted_key(key)
            assert spec is not None
            assert spec.rel_path not in RESET_EXCLUDED


@pytest.mark.parametrize(
    ("group", "owned_updates", "owned_defaults"),
    [
        (
            "spreadsheet",
            {
                "software.default_priority.spreadsheet_processors": ["libreoffice"],
                "software.special_conversions.ods": ["libreoffice"],
                "software.special_conversions.spreadsheet_to_pdf": ["libreoffice"],
            },
            {
                "software.default_priority.spreadsheet_processors": [
                    "wps_spreadsheets",
                    "msoffice_excel",
                    "libreoffice",
                ],
                "software.special_conversions.ods": ["msoffice_excel", "libreoffice"],
                "software.special_conversions.spreadsheet_to_pdf": [
                    "wps_spreadsheets",
                    "msoffice_excel",
                    "libreoffice",
                ],
            },
        ),
        (
            "layout",
            {"software.special_conversions.pdf_to_office": ["libreoffice"]},
            {"software.special_conversions.pdf_to_office": ["msoffice_word", "libreoffice"]},
        ),
    ],
)
def test_reset_group_restores_owned_software_values_without_crossing_siblings(
    tmp_path: Path,
    group: str,
    owned_updates: dict[str, list[str]],
    owned_defaults: dict[str, list[str]],
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "user")
    for key, value in owned_updates.items():
        assert loader.set_value(key, value) is True
    sibling = ["libreoffice"]
    assert loader.set_value("software.default_priority.word_processors", sibling) is True

    assert loader.reset_group(group) is True

    data = loader.config.as_dict()
    for key, expected in owned_defaults.items():
        current = data
        for part in key.split("."):
            current = current[part]
        assert current == expected
    assert data["software"]["default_priority"]["word_processors"] == sibling


@pytest.mark.parametrize(
    ("group", "owned_key", "owned_override", "owned_default", "sibling_key", "sibling_value"),
    [
        (
            "text",
            "text.remove_numbering",
            False,
            True,
            "text.to_xlsx_remove_numbering",
            False,
        ),
        (
            "document",
            "document.to_md_keep_images",
            False,
            True,
            "document.extension_probe",
            "base64",
        ),
        (
            "spreadsheet",
            "spreadsheet.to_md_keep_images",
            False,
            True,
            "spreadsheet.extension_probe",
            "base64",
        ),
        (
            "layout",
            "layout.to_md_keep_images",
            False,
            True,
            "layout.extension_probe",
            "base64",
        ),
        (
            "other",
            "other.to_md_keep_images",
            False,
            True,
            "other.extension_probe",
            "base64",
        ),
        (
            "link",
            "link.embedding.max_depth",
            9,
            3,
            "link.path_resolution.search_dirs",
            ["custom-assets"],
        ),
        (
            "output",
            "output.directory.mode",
            "custom",
            "source",
            "output.manifest.save_to_output",
            True,
        ),
        (
            "logging",
            "logger.level",
            "warning",
            "debug",
            "logger.format",
            "CUSTOM {message}",
        ),
    ],
)
def test_reset_group_preserves_non_owned_file_siblings(
    tmp_path: Path,
    group: str,
    owned_key: str,
    owned_override: object,
    owned_default: object,
    sibling_key: str,
    sibling_value: object,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "user")
    assert loader.set_value(owned_key, owned_override) is True
    assert loader.set_value(sibling_key, sibling_value) is True

    assert loader.reset_group(group) is True

    data = loader.config.as_dict()

    def read(key: str) -> object:
        current: object = data
        for part in key.split("."):
            assert isinstance(current, dict)
            current = current[part]
        return current

    assert read(owned_key) == owned_default
    assert read(sibling_key) == sibling_value


def test_root_config_files_wrap_to_declared_namespace(tmp_path) -> None:
    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    write_minimal_base_config_tree(base_dir)
    (base_dir / "document.toml").write_text("keep_images = true\n", encoding="utf-8")

    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

    assert loader.config.document.keep_images is True
    with pytest.raises(AttributeError):
        _ = loader.config.defaults.document


def test_subdirectory_files_wrap_to_declared_namespace(tmp_path) -> None:
    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    write_minimal_base_config_tree(base_dir)
    (base_dir / "numbering" / "add.toml").write_text(
        '[settings]\ndefault_scheme = "gongwen_standard"\n',
        encoding="utf-8",
    )
    (base_dir / "proofread" / "engine.toml").write_text(
        "enable_typos_rule = true\n",
        encoding="utf-8",
    )

    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

    assert loader.config.numbering.add.settings.default_scheme == "gongwen_standard"
    assert loader.config.proofread.engine.enable_typos_rule is True
    with pytest.raises(AttributeError):
        _ = loader.config.engine


def test_user_overrides_base_for_same_file(tmp_path) -> None:
    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    write_minimal_base_config_tree(base_dir)
    (base_dir / "link.toml").write_text('[format]\nimage_link_style = "wiki_embed"\n', encoding="utf-8")
    (user_dir / "link.toml").parent.mkdir(parents=True, exist_ok=True)
    (user_dir / "link.toml").write_text('[format]\nimage_link_style = "markdown_link"\n', encoding="utf-8")

    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

    assert loader.config.link.format.image_link_style == "markdown_link"


def test_invalid_user_override_is_ignored_and_base_remains(tmp_path) -> None:
    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    write_minimal_base_config_tree(base_dir)
    (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")
    invalid_bytes = b"not = valid = toml = !!!\n"
    (user_dir / "gui.toml").write_bytes(invalid_bytes)

    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

    assert loader.config.gui.window.default_mode == "single"
    assert not (user_dir / "gui.toml").exists()
    backups = list(user_dir.glob("gui.toml.bak_parse_failed_*"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == invalid_bytes


def test_semantic_invalid_user_override_is_quarantined_before_trust(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    invalid_bytes = b'level = "verbose"\nretention_days = "abc"\n'
    (user_dir / "logger.toml").write_bytes(invalid_bytes)

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    logger_config = loader.config.as_dict()["logger"]
    assert logger_config["level"] == "debug"
    assert logger_config["retention_days"] == 30
    assert loader.config_state_trusted is True
    assert not (user_dir / "logger.toml").exists()
    backups = list(user_dir.glob("logger.toml.bak_schema_failed_*"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == invalid_bytes


@pytest.mark.parametrize(
    ("rel_path", "invalid_text"),
    (
        ("logger.toml", 'file_prefix = "do:c*wen"\n'),
        ("logger.toml", 'directory_mode = "custom"\ndirectory = ""\n'),
        ("logger.toml", 'console_colorize = "sometimes"\n'),
        ("output.toml", '[directory]\nmode = "invalid"\n'),
        ("gui.toml", '[theme]\ndefault_theme = "blue"\n'),
    ),
)
def test_historical_semantic_constraints_quarantine_invalid_overrides(
    tmp_path: Path,
    rel_path: str,
    invalid_text: str,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_path = user_dir / rel_path
    user_path.parent.mkdir(parents=True)
    invalid_bytes = invalid_text.encode("utf-8")
    user_path.write_bytes(invalid_bytes)

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    assert loader.config_state_trusted is True
    assert not user_path.exists()
    backups = list(user_path.parent.glob(f"{user_path.name}.bak_schema_failed_*"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == invalid_bytes


def test_noncanonical_logger_values_are_quarantined(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    (user_dir / "logger.toml").write_text(
        'level = " INFO "\nconsole_level = "WARNING"\nfile_prefix = " audit "\n',
        encoding="utf-8",
    )

    ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    assert not (user_dir / "logger.toml").exists()
    backups = list(user_dir.glob("logger.toml.bak_schema_failed_*"))
    assert len(backups) == 1


def test_known_shape_validation_covers_non_explicit_registry_files(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    (user_dir / "link.toml").write_text('[embedding]\nmax_depth = "3"\n', encoding="utf-8")

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    assert loader.config.link.embedding.max_depth == 3
    assert len(list(user_dir.glob("link.toml.bak_schema_failed_*"))) == 1


def test_unknown_extension_keys_remain_forward_compatible(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    (user_dir / "link.toml").write_text(
        '[extension]\nowner = "future-plugin"\nrevision = 1\n',
        encoding="utf-8",
    )

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    assert loader.config.link.extension.owner == "future-plugin"
    assert loader.config.link.extension.revision == 1


def test_semantic_invalid_conversion_cannot_poison_request_projection(tmp_path: Path) -> None:
    from docwen_core.export_semantics import MarkdownExportSemantics
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    (user_dir / "conversion.toml").write_text(
        '[export]\nbase64_compress_threshold_kb = "abc"\n',
        encoding="utf-8",
    )

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    semantics = MarkdownExportSemantics.from_config_snapshot(loader.config.as_dict())

    assert semantics.export_base64_compress_threshold_kb == 100
    assert len(list(user_dir.glob("conversion.toml.bak_schema_failed_*"))) == 1
