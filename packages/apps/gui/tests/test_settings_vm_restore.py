"""GUI settings ViewModel logical-owner reset tests.

Proves delegation to the runtime reset plan, including groups whose values
span files or share physical files with another Settings tab.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_runtime.config import ConfigLoader

pytestmark = pytest.mark.unit

PROJECT_CONFIGS = Path(__file__).resolve().parent.parent.parent.parent.parent / "configs"


# ── Fixtures ────────────────────────────────────────────────────────────


@pytest.fixture
def config_dir(tmp_path: Path) -> Path:
    """Provides a fresh user config directory for each test."""
    return tmp_path / "configs"


@pytest.fixture
def config_loader(config_dir: Path) -> ConfigLoader:
    """Create one explicit loader owned by this test composition."""
    return ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=config_dir)


@pytest.fixture
def config_port(config_loader: ConfigLoader) -> ConfigPortAdapter:
    return ConfigPortAdapter(config_loader)


@pytest.fixture
def vm(config_port: ConfigPortAdapter) -> SettingsViewModel:
    """Create a SettingsViewModel backed by the injected config composition."""
    controller = ApplicationController(config_port=config_port)
    return SettingsViewModel(controller=controller)


# Import at module level — SettingsViewModel is needed by fixtures
from docwen_gui.view_models.settings_vm import (
    SECTION_CONVERSION_DEFAULTS,
    SECTION_EXPORT,
    SECTION_FORMATTING,
    SECTION_GUI,
    SECTION_LINK,
    SECTION_LOGGING,
    SECTION_OUTPUT,
    SECTION_TEXT,
    SettingsViewModel,
)

# ── Section-level isolation tests ──────────────────────────────────────


class TestResetSectionIsolation:
    """Reset of one logical group must not cascade to another group."""

    def test_reset_gui_leaves_output_intact(self, vm: SettingsViewModel) -> None:
        """Resetting General leaves Output intact."""
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.set_field(SECTION_GUI, "md_default_template", "xlsx")
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")
        vm.apply_settings()

        assert vm.config.gui.theme == "dark"
        assert vm.config.output.output_mode == "custom"

        # Reset gui only
        ok = vm.reset_section(SECTION_GUI)
        assert ok is True
        assert vm.config.gui.theme == "light"
        assert vm.config.gui.md_default_template == "xlsx"
        # output must survive
        assert vm.config.output.output_mode == "custom"

    def test_reset_text_restores_template_type_without_resetting_general(self, vm: SettingsViewModel) -> None:
        """Text owns the GUI template preference even though it lives in gui.toml."""
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.set_field(SECTION_GUI, "md_default_template", "xlsx")
        vm.apply_settings()

        assert vm.reset_section(SECTION_TEXT) is True

        assert vm.config.gui.md_default_template == "docx"
        assert vm.config.gui.theme == "dark"

    def test_reset_output_leaves_logger_intact(self, vm: SettingsViewModel) -> None:
        """Resetting Output leaves Logging intact."""
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")
        vm.set_field(SECTION_LOGGING, "level", "warning")
        vm.apply_settings()

        ok = vm.reset_section(SECTION_OUTPUT)
        assert ok is True
        assert vm.config.output.output_mode == "source"
        assert vm.config.logging.level == "warning"

    def test_reset_formatting_leaves_gui_intact(self, vm: SettingsViewModel) -> None:
        """Resetting Formatting leaves General intact."""
        vm.set_field(SECTION_FORMATTING, "body_format", "discard")
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_settings()

        ok = vm.reset_section(SECTION_FORMATTING)
        assert ok is True
        assert vm.config.formatting.body_format == "preserve"  # default
        assert vm.config.gui.theme == "dark"

    def test_reset_logging_leaves_export_intact(self, vm: SettingsViewModel) -> None:
        """Resetting Logging must not touch Export.

        Uses ``base64_compress_enabled``, which is persisted to
        ``conversion.export.base64_compress_enabled`` and correctly
        round-tripped by ``_map_raw_to_config``.
        """
        vm.set_field(SECTION_EXPORT, "base64_compress_enabled", False)
        vm.set_field(SECTION_LOGGING, "level", "warning")
        vm.apply_settings()

        assert vm.config.export.base64_compress_enabled is False
        assert vm.config.logging.level == "warning"

        ok = vm.reset_section(SECTION_LOGGING)
        assert ok is True
        assert vm.config.logging.level == "debug"  # default restored
        # Export must survive because it has a different logical owner.
        assert vm.config.export.base64_compress_enabled is False

    def test_reset_conversion_defaults_leaves_gui_intact(self, vm: SettingsViewModel) -> None:
        """The compatibility aggregate does not cross-reset General."""
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_settings()

        ok = vm.reset_section(SECTION_CONVERSION_DEFAULTS)
        assert ok is True
        assert vm.config.gui.theme == "dark"

    def test_reset_document_group_leaves_image_defaults_intact(
        self,
        vm: SettingsViewModel,
        config_loader: ConfigLoader,
    ) -> None:
        """Tab-level reset must use the precise registry group and its same-tab software keys."""
        loader = config_loader
        loader.set_value("document.to_md_keep_images", False)
        loader.set_value("image.to_md_enable_ocr", False)
        loader.set_value("software.default_priority.word_processors", ["libreoffice", "msoffice_word", "wps_writer"])
        loader.set_value(
            "software.default_priority.spreadsheet_processors",
            ["libreoffice", "msoffice_excel", "wps_spreadsheets"],
        )

        vm.load_from_controller_config()
        assert vm.config.conversion_defaults.document["to_md_keep_images"] is False
        assert vm.config.conversion_defaults.image["to_md_enable_ocr"] is False
        assert vm.config.software_priority.word_processors == ["libreoffice", "msoffice_word", "wps_writer"]
        assert vm.config.software_priority.spreadsheet_processors == [
            "libreoffice",
            "msoffice_excel",
            "wps_spreadsheets",
        ]

        ok = vm.reset_group("document")
        assert ok is True

        assert vm.config.conversion_defaults.document["to_md_keep_images"] is True
        assert vm.config.conversion_defaults.image["to_md_enable_ocr"] is False
        assert vm.config.software_priority.word_processors == ["wps_writer", "msoffice_word", "libreoffice"]
        assert vm.config.software_priority.spreadsheet_processors == [
            "libreoffice",
            "msoffice_excel",
            "wps_spreadsheets",
        ]

    def test_reset_export_group_resets_split_owned_values_only(
        self,
        vm: SettingsViewModel,
        config_loader: ConfigLoader,
    ) -> None:
        """Export owns ``export.toml`` plus two sections in ``conversion.toml``."""
        loader = config_loader
        loader.set_value("export.to_md_image_extraction_mode", "base64")
        loader.set_value("export.to_md_ocr_placement_mode", "image_md")
        loader.set_value("conversion.export.base64_compress_enabled", False)
        loader.set_value("conversion.export.base64_compress_threshold_kb", 512)
        loader.set_value("conversion.ocr_output.show_blockquote_title", False)
        loader.set_value(
            "conversion.ocr_output.blockquote_title_override_by_locale.zh_CN",
            "自定义 OCR 标题",
        )
        loader.set_value("conversion.syntax.bold", "underscore")

        vm.load_from_controller_config()
        assert vm.config.export.image_mode == "base64"
        assert vm.config.export.base64_compress_enabled is False
        assert vm.config.export.base64_compress_threshold_kb == 512
        assert vm.config.export.ocr_title_enabled is False
        assert vm.config.formatting.bold_syntax == "underscore"

        assert vm.reset_group("export") is True

        assert vm.config.export.image_mode == "file"
        assert vm.config.export.ocr_mode == "main_md"
        assert vm.config.export.base64_compress_enabled is True
        assert vm.config.export.base64_compress_threshold_kb == 100
        assert vm.config.export.ocr_title_enabled is True
        assert loader.config.as_dict()["conversion"]["ocr_output"]["blockquote_title_override_by_locale"] == {}
        assert vm.config.formatting.bold_syntax == "underscore"

    def test_reset_formatting_group_preserves_export_and_resets_table_style(
        self,
        vm: SettingsViewModel,
        config_loader: ConfigLoader,
    ) -> None:
        """Formatting owns precise conversion keys and the document table-style keys."""
        loader = config_loader
        loader.set_value("conversion.syntax.bold", "underscore")
        loader.set_value("conversion.md_to_docx.list_separator", ", ")
        loader.set_value("document.style.table.md_to_docx.table_style_mode", "custom")
        loader.set_value("document.style.table.md_to_docx.custom_style_name", "Research Table")
        loader.set_value("conversion.export.base64_compress_enabled", False)
        loader.set_value("conversion.ocr_output.show_blockquote_title", False)

        vm.load_from_controller_config()
        assert vm.config.formatting.bold_syntax == "underscore"
        assert vm.config.formatting.list_separator == ", "
        assert vm.config.formatting.table_style_mode == "custom"
        assert vm.config.formatting.custom_table_style_name == "Research Table"
        assert vm.config.export.base64_compress_enabled is False
        assert vm.config.export.ocr_title_enabled is False

        assert vm.reset_group("formatting") is True

        assert vm.config.formatting.bold_syntax == "asterisk"
        assert vm.config.formatting.list_separator == "、"
        assert vm.config.formatting.table_style_mode == "builtin"
        assert vm.config.formatting.custom_table_style_name == ""
        assert vm.config.export.base64_compress_enabled is False
        assert vm.config.export.ocr_title_enabled is False

    def test_reset_document_group_preserves_non_owned_document_styles(
        self,
        vm: SettingsViewModel,
        config_loader: ConfigLoader,
    ) -> None:
        """Document defaults and Formatting/style overrides share ``document.toml``."""
        loader = config_loader
        loader.set_value("document.to_md_keep_images", False)
        loader.set_value("document.style.code.docx_to_md.fuzzy_match_enabled", False)
        loader.set_value("document.style.table.md_to_docx.table_style_mode", "custom")
        loader.set_value("document.style.table.md_to_docx.custom_style_name", "Research Table")

        vm.load_from_controller_config()
        assert vm.config.conversion_defaults.document["to_md_keep_images"] is False
        assert vm.config.formatting.table_style_mode == "custom"

        assert vm.reset_group("document") is True

        assert vm.config.conversion_defaults.document["to_md_keep_images"] is True
        assert loader.config.document.style.code.docx_to_md.fuzzy_match_enabled is False
        assert vm.config.formatting.table_style_mode == "custom"
        assert vm.config.formatting.custom_table_style_name == "Research Table"

    def test_reset_proofread_group_preserves_curated_dictionaries_and_reloads(
        self,
        vm: SettingsViewModel,
        config_loader: ConfigLoader,
    ) -> None:
        """Protected user dictionaries are excluded without turning reset into a partial failure."""
        loader = config_loader
        loader.set_value("proofread.engine.enable_symbol_pairing", False)
        loader.set_value("proofread.skip.code_blocks", False)
        loader.set_value("proofread.skip.log_skipped", False)
        loader.set_value("proofread.pairs.items", [["<", ">"]])
        loader.set_value("proofread.symbol_map.entries", {"!": ["！"]})
        loader.set_value("proofread.typos.entries", {"teh": ["the"]})
        loader.set_value("proofread.sensitive_words.entries", {"secret": ["allowed"]})
        vm.load_from_controller_config()
        assert vm.config.proofread.symbol_pairing is False
        assert vm.config.proofread.skip_code_blocks is False
        symbol_map_before = loader.config.as_dict()["proofread"]["symbol_map"]["entries"]

        assert vm.reset_group("proofread") is True

        assert vm.config.proofread.symbol_pairing is True
        proofread = loader.config.as_dict()["proofread"]
        assert proofread["skip"]["code_blocks"] is True
        assert proofread["skip"]["log_skipped"] is False
        assert proofread["pairs"]["items"] == [["<", ">"]]
        assert proofread["symbol_map"]["entries"] == symbol_map_before
        assert proofread["typos"]["entries"] == {"teh": ["the"]}
        assert proofread["sensitive_words"]["entries"] == {"secret": ["allowed"]}

    def test_reset_group_uses_injected_config_port_not_unrelated_loader(
        self,
        config_dir: Path,
    ) -> None:
        """A custom composition root resets its own user-config directory."""
        injected_loader = ConfigLoader(
            base_dir=PROJECT_CONFIGS,
            user_dir=config_dir / "injected",
        )
        injected_port = ConfigPortAdapter(injected_loader)
        assert injected_port.set("conversion.syntax.bold", "underscore") is True

        unrelated_loader = ConfigLoader(
            base_dir=PROJECT_CONFIGS,
            user_dir=config_dir / "unrelated",
        )
        assert unrelated_loader.set_value("conversion.syntax.bold", "underscore") is True

        custom_vm = SettingsViewModel(
            controller=ApplicationController(config_port=injected_port),
        )
        assert custom_vm.config.formatting.bold_syntax == "underscore"

        assert custom_vm.reset_group("formatting") is True

        assert custom_vm.config.formatting.bold_syntax == "asterisk"
        assert injected_port.get("conversion.syntax.bold") == "asterisk"
        assert unrelated_loader.config.conversion.syntax.bold == "underscore"

    def test_four_config_toml_sections_are_independent(self, vm: SettingsViewModel) -> None:
        """Reset Output without changing General, Logging, or Formatting."""
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")
        vm.set_field(SECTION_LOGGING, "level", "warning")
        vm.set_field(SECTION_FORMATTING, "body_format", "discard")
        vm.apply_settings()

        # Reset output only — three other sections must survive
        ok = vm.reset_section(SECTION_OUTPUT)
        assert ok is True

        assert vm.config.output.output_mode == "source"  # default restored
        assert vm.config.gui.theme == "dark"
        assert vm.config.logging.level == "warning"
        assert vm.config.formatting.body_format == "discard"


class TestResetSectionCrossFile:
    """Sections in different files must not affect each other.

    Uses the direct loader API to set values because ``_persist_to_controller_config``
    does not yet write all section types (including Link).  The isolation
    contract is between the on-disk sections, not the persist layer.
    """

    def test_reset_formatting_does_not_touch_link(
        self,
        vm: SettingsViewModel,
        config_loader: ConfigLoader,
    ) -> None:
        """Formatting reset leaves the separate Link group intact."""
        loader = config_loader

        # Set values in different files via loader
        loader.set_value("conversion.syntax.bold", "underscore")
        loader.set_value("link.embedding.max_depth", 10)

        # Sync VM from loader
        vm.load_from_controller_config()
        assert vm.config.formatting.bold_syntax == "underscore"
        assert vm.config.link.max_depth == 10

        # Reset Formatting.
        ok = vm.reset_section(SECTION_FORMATTING)
        assert ok is True
        assert vm.config.formatting.bold_syntax == "asterisk"  # default restored
        # Link must be untouched.
        assert vm.config.link.max_depth == 10

    def test_reset_link_does_not_touch_formatting(
        self,
        vm: SettingsViewModel,
        config_loader: ConfigLoader,
    ) -> None:
        """Link reset leaves the separate Formatting group intact."""
        loader = config_loader

        loader.set_value("link.embedding.max_depth", 10)
        loader.set_value("conversion.syntax.italic", "underscore")

        vm.load_from_controller_config()
        assert vm.config.link.max_depth == 10
        assert vm.config.formatting.italic_syntax == "underscore"

        # Reset Link.
        ok = vm.reset_section(SECTION_LINK)
        assert ok is True
        assert vm.config.link.max_depth == 3  # default restored
        # Formatting must be untouched.
        assert vm.config.formatting.italic_syntax == "underscore"


class TestResetEmitsStatus:
    """Section reset must emit human-readable status messages."""

    def test_reset_emits_success_status(self, vm: SettingsViewModel) -> None:
        status_signals: list = []
        vm.status_changed.connect(lambda msg, err: status_signals.append((msg, err)))
        ok = vm.reset_section(SECTION_GUI)
        assert ok is True
        assert len(status_signals) == 1
        assert "reset" in status_signals[0][0].lower()
        assert status_signals[0][1] is False  # not an error

    def test_unknown_section_emits_error(self, vm: SettingsViewModel) -> None:
        status_signals: list = []
        vm.status_changed.connect(lambda msg, err: status_signals.append((msg, err)))
        ok = vm.reset_section("nonexistent")
        assert ok is False
        assert len(status_signals) == 1
        assert status_signals[0][1] is True  # is an error


class TestResetSectionReload:
    """After reset, the VM must reload from disk so widgets see fresh data."""

    def test_reset_emits_reload_signal(self, vm: SettingsViewModel) -> None:
        reload_signals: list = []
        vm.config_reloaded.connect(lambda: reload_signals.append(True))
        vm.reset_section(SECTION_GUI)
        assert len(reload_signals) >= 1

    def test_reset_clears_dirty_state(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert vm.is_dirty is True
        vm.reset_section(SECTION_GUI)
        assert vm.is_dirty is False

    def test_reset_updates_session_baseline_so_cancel_keeps_reset_defaults(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert vm.apply_settings() is True
        assert vm.config.gui.theme == "dark"

        vm.begin_session()
        assert vm.reset_section(SECTION_GUI) is True
        assert vm.config.gui.theme == "light"
        assert vm.is_dirty is False
        assert vm.get_change_summary() == []

        vm.cancel_changes()

        assert vm.config.gui.theme == "light"
        assert vm.is_dirty is False
        assert vm.get_change_summary() == []
