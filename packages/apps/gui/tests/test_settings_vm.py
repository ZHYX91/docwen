"""Model-state tests for SettingsViewModel.

These tests validate that the ViewModel is the source of truth for
settings state, that dirty tracking works, and that signals fire correctly.
No QApplication is needed.
"""

from __future__ import annotations

import inspect
from pathlib import Path
from typing import TYPE_CHECKING

import pytest

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_runtime.config import ConfigLoader

if TYPE_CHECKING:
    pass

from docwen_gui.view_models.settings_vm import (
    SECTION_FORMATTING,
    SECTION_GUI,
    SECTION_LINK,
    SECTION_LOGGING,
    SECTION_OUTPUT,
    SECTION_SOFTWARE_PRIORITY,
    SECTION_TEXT,
    SettingsViewModel,
)

pytestmark = pytest.mark.unit

PROJECT_CONFIGS = Path(__file__).resolve().parent.parent.parent.parent.parent / "configs"


# ── Fixtures ────────────────────────────────────────────────────────────


@pytest.fixture
def config_loader(tmp_path: Path) -> ConfigLoader:
    return ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")


@pytest.fixture
def config_port(config_loader: ConfigLoader) -> ConfigPortAdapter:
    return ConfigPortAdapter(config_loader)


@pytest.fixture
def vm(config_port: ConfigPortAdapter) -> SettingsViewModel:
    controller = ApplicationController(config_port=config_port)
    return SettingsViewModel(controller=controller)


# ── Initial state ───────────────────────────────────────────────────────


class TestInitialState:
    def test_not_dirty_on_start(self, vm: SettingsViewModel) -> None:
        assert vm.is_dirty is False

    def test_default_theme_is_light(self, vm: SettingsViewModel) -> None:
        assert vm.config.gui.theme == "light"

    def test_default_mode_is_single(self, vm: SettingsViewModel) -> None:
        assert vm.config.gui.default_mode == "single"

    def test_default_transparency_disabled(self, vm: SettingsViewModel) -> None:
        assert vm.config.gui.transparency_enabled is False
        assert vm.config.gui.transparency_value == 0.95

    def test_default_proofread_all_enabled(self, vm: SettingsViewModel) -> None:
        proof = vm.config.proofread
        assert proof.symbol_pairing is True
        assert proof.symbol_correction is True
        assert proof.typos_rule is True
        assert proof.sensitive_word is True


class TestTemplateState:
    def test_template_name_selection_is_session_state_only(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "md_default_template", "docx")
        vm.select_template("xlsx", "Quarterly Sheet")

        assert vm.selected_templates == {"xlsx": "Quarterly Sheet"}
        assert vm.get_field(SECTION_GUI, "md_default_template") == "docx"

    def test_settings_vm_template_comments_track_persisted_boundary(self) -> None:
        source_path = Path(inspect.getsourcefile(SettingsViewModel) or "")
        text = source_path.read_text(encoding="utf-8")

        assert "schema extension pending" not in text
        assert "Also sets ``gui.md_default_template``" not in text
        assert "Only the default template type is persisted as gui.md_default_template" in text


# ── Field mutation ──────────────────────────────────────────────────────


class TestFieldMutation:
    def test_set_field_emits_signal(self, vm: SettingsViewModel) -> None:
        signals: list = []
        vm.config_changed.connect(lambda s, k, v: signals.append((s, k, v)))
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert len(signals) == 1
        assert signals[0] == (SECTION_GUI, "theme", "dark")

    def test_set_field_dirty(self, vm: SettingsViewModel) -> None:
        dirty_signals: list[bool] = []
        vm.dirty_state_changed.connect(dirty_signals.append)
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert vm.is_dirty is True
        assert dirty_signals == [True]

    def test_set_same_value_no_signal(self, vm: SettingsViewModel) -> None:
        signals: list = []
        vm.config_changed.connect(lambda s, k, v: signals.append((s, k, v)))
        vm.set_field(SECTION_GUI, "theme", "light")  # already light
        assert len(signals) == 0

    def test_set_field_no_such_key(self, vm: SettingsViewModel) -> None:
        """Setting a non-existent field key should be a no-op."""
        signals: list = []
        vm.config_changed.connect(lambda s, k, v: signals.append((s, k, v)))
        vm.set_field(SECTION_GUI, "nonexistent_key", 123)
        assert len(signals) == 0

    def test_set_field_no_such_section(self, vm: SettingsViewModel) -> None:
        """Setting a field on a non-existent section should be a no-op."""
        signals: list = []
        vm.config_changed.connect(lambda s, k, v: signals.append((s, k, v)))
        vm.set_field("nonexistent_section", "key", 123)
        assert len(signals) == 0

    def test_set_field_normalizes_pdf_to_office_priority(self, vm: SettingsViewModel) -> None:
        signals: list = []
        vm.config_changed.connect(lambda s, k, v: signals.append((s, k, v)))

        vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            "pdf_to_office",
            ["wps_writer", "libreoffice", "unknown_backend", "msoffice_word", "libreoffice"],
        )

        assert vm.config.software_priority.pdf_to_office == ["libreoffice", "msoffice_word"]
        assert signals == [(SECTION_SOFTWARE_PRIORITY, "pdf_to_office", ["libreoffice", "msoffice_word"])]

    def test_set_field_batch_normalizes_pdf_to_office_priority(self, vm: SettingsViewModel) -> None:
        vm.set_field_batch(
            SECTION_SOFTWARE_PRIORITY,
            {
                "pdf_to_office": ["wps_writer"],
                "word_processors": ["libreoffice", "msoffice_word", "wps_writer"],
            },
        )

        assert vm.config.software_priority.pdf_to_office == ["msoffice_word", "libreoffice"]
        assert vm.config.software_priority.word_processors == ["libreoffice", "msoffice_word", "wps_writer"]

    def test_set_field_normalizes_odt_and_ods_special_conversion_priorities(self, vm: SettingsViewModel) -> None:
        vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            "odt_conversion",
            ["wps_writer", "libreoffice", "unknown_backend"],
        )
        vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            "ods_conversion",
            ["wps_spreadsheets", "libreoffice", "msoffice_excel", "libreoffice"],
        )

        assert vm.config.software_priority.odt_conversion == ["libreoffice", "msoffice_word"]
        assert vm.config.software_priority.ods_conversion == ["libreoffice", "msoffice_excel"]


# ── Batch mutation ──────────────────────────────────────────────────────


class TestBatchMutation:
    def test_set_field_batch_emits_reload(self, vm: SettingsViewModel) -> None:
        reload_signals: list = []
        vm.config_reloaded.connect(lambda: reload_signals.append(True))
        vm.set_field_batch(SECTION_GUI, {"theme": "dark", "language": "en_US"})
        assert len(reload_signals) == 1

    def test_set_field_batch_no_change(self, vm: SettingsViewModel) -> None:
        reload_signals: list = []
        vm.config_reloaded.connect(lambda: reload_signals.append(True))
        vm.set_field_batch(SECTION_GUI, {"theme": "light", "language": "zh_CN"})
        assert len(reload_signals) == 0


# ── Apply / Cancel / Reset ──────────────────────────────────────────────


class TestApplyCancel:
    def test_apply_clears_dirty(self, vm: SettingsViewModel) -> None:
        dirty_signals: list[bool] = []
        vm.dirty_state_changed.connect(dirty_signals.append)
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert vm.is_dirty is True
        vm.apply_settings()
        assert vm.is_dirty is False
        assert dirty_signals == [True, False]

    def test_apply_emits_status(self, vm: SettingsViewModel) -> None:
        status_signals: list = []
        vm.status_changed.connect(lambda msg, err: status_signals.append((msg, err)))
        vm.apply_settings()
        assert len(status_signals) == 1
        assert "applied" in status_signals[0][0].lower()
        assert status_signals[0][1] is False  # not error

    def test_apply_persists_logging_console_colorize(
        self, vm: SettingsViewModel, config_port: ConfigPortAdapter
    ) -> None:
        vm.set_field(SECTION_LOGGING, "console_colorize", "never")

        assert vm.apply_settings() is True

        assert config_port.snapshot()["logger"]["console_colorize"] == "never"

    def test_apply_persists_software_priority(self, vm: SettingsViewModel, config_port: ConfigPortAdapter) -> None:
        vm.set_field(SECTION_SOFTWARE_PRIORITY, "word_processors", ["libreoffice", "msoffice_word", "wps_writer"])
        vm.set_field(SECTION_SOFTWARE_PRIORITY, "document_to_pdf", ["msoffice_word", "wps_writer", "libreoffice"])
        vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            "spreadsheet_to_pdf",
            ["msoffice_excel", "wps_spreadsheets", "libreoffice"],
        )

        assert vm.apply_settings() is True

        vm.load_from_controller_config()
        assert vm.config.software_priority.word_processors == ["libreoffice", "msoffice_word", "wps_writer"]
        special_conversions = config_port.snapshot()["software"]["special_conversions"]
        assert special_conversions["document_to_pdf"] == ["msoffice_word", "wps_writer", "libreoffice"]
        assert special_conversions["spreadsheet_to_pdf"] == [
            "msoffice_excel",
            "wps_spreadsheets",
            "libreoffice",
        ]

    def test_apply_persists_normalized_pdf_to_office_priority(
        self, vm: SettingsViewModel, config_port: ConfigPortAdapter
    ) -> None:
        vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            "pdf_to_office",
            ["wps_writer", "libreoffice", "unknown_backend", "msoffice_word"],
        )

        assert vm.apply_settings() is True

        assert config_port.snapshot()["software"]["special_conversions"]["pdf_to_office"] == [
            "libreoffice",
            "msoffice_word",
        ]

    def test_apply_persists_conversion_defaults(self, vm: SettingsViewModel, config_port: ConfigPortAdapter) -> None:
        updates = {
            "document": {
                "to_md_keep_images": False,
                "to_md_enable_ocr": True,
                "to_md_enable_optimization": True,
                "to_md_optimization_type": "contract",
                "to_md_table_merge_export_strategy": "marker",
            },
            "spreadsheet": {
                "to_md_keep_images": False,
                "to_md_enable_ocr": True,
                "to_md_table_merge_export_strategy": "marker",
                "merge_mode": 2,
            },
            "image": {
                "to_md_keep_images": False,
                "to_md_enable_ocr": False,
                "ocr_language": "japanese",
                "compress_mode": "limit_size",
                "size_limit": 1024,
                "size_unit": "MB",
                "pdf_quality": "fit_a4",
                "tiff_mode": "rgb",
            },
            "layout": {
                "to_md_keep_images": False,
                "to_md_enable_ocr": True,
                "to_md_enable_optimization": True,
                "render_dpi": 600,
            },
            "other": {
                "to_md_keep_images": True,
                "to_md_enable_ocr": True,
            },
        }
        for category, category_updates in updates.items():
            for key, value in category_updates.items():
                vm.set_conversion_default(category, key, value)

        assert vm.apply_settings() is True

        snapshot = config_port.snapshot()
        for category, category_updates in updates.items():
            for key, value in category_updates.items():
                assert snapshot[category][key] == value

    def test_apply_persists_md_to_docx_table_style_settings(
        self, vm: SettingsViewModel, config_port: ConfigPortAdapter
    ) -> None:
        vm.set_field(SECTION_FORMATTING, "table_style_mode", "custom")
        vm.set_field(SECTION_FORMATTING, "builtin_table_style", "table_grid")
        vm.set_field(SECTION_FORMATTING, "custom_table_style_name", "Research Table")

        assert vm.apply_settings() is True

        snapshot = config_port.snapshot()
        table_settings = snapshot["document"]["style"]["table"]["md_to_docx"]
        assert table_settings["table_style_mode"] == "custom"
        assert table_settings["builtin_style_key"] == "table_grid"
        assert table_settings["custom_style_name"] == "Research Table"

    def test_apply_persists_document_to_markdown_numbering_defaults(
        self, vm: SettingsViewModel, config_port: ConfigPortAdapter
    ) -> None:
        vm.set_conversion_default("document", "to_md_remove_numbering", False)
        vm.set_conversion_default("document", "to_md_add_numbering", True)
        vm.set_conversion_default("document", "to_md_default_scheme", "legal_standard")

        assert vm.apply_settings() is True

        document = config_port.snapshot()["document"]
        assert document["to_md_remove_numbering"] is False
        assert document["to_md_add_numbering"] is True
        assert document["to_md_default_scheme"] == "legal_standard"

        fresh_vm = SettingsViewModel(controller=ApplicationController(config_port=config_port))
        assert fresh_vm.config.conversion_defaults.document["to_md_remove_numbering"] is False
        assert fresh_vm.config.conversion_defaults.document["to_md_add_numbering"] is True
        assert fresh_vm.config.conversion_defaults.document["to_md_default_scheme"] == "legal_standard"

    @pytest.mark.parametrize("separator", [", ", ""])
    def test_apply_roundtrips_yaml_list_separator_exactly(
        self,
        separator: str,
        vm: SettingsViewModel,
        config_port: ConfigPortAdapter,
    ) -> None:
        vm.set_field(SECTION_FORMATTING, "list_separator", separator)

        assert vm.apply_settings() is True
        assert config_port.snapshot()["conversion"]["md_to_docx"]["list_separator"] == separator

        fresh_vm = SettingsViewModel(controller=ApplicationController(config_port=config_port))
        assert fresh_vm.config.formatting.list_separator == separator

    @pytest.mark.parametrize("punctuation", ["：§", ""])
    def test_apply_roundtrips_heading_merge_punctuation_exactly(
        self,
        punctuation: str,
        vm: SettingsViewModel,
        config_port: ConfigPortAdapter,
    ) -> None:
        vm.set_field(SECTION_FORMATTING, "heading_merge_punctuation", punctuation)

        assert vm.apply_settings() is True
        assert config_port.snapshot()["conversion"]["md_to_docx"]["heading_merge_punctuation"] == punctuation

        fresh_vm = SettingsViewModel(controller=ApplicationController(config_port=config_port))
        assert fresh_vm.config.formatting.heading_merge_punctuation == punctuation

    def test_apply_persists_text_fact_sources(
        self,
        vm: SettingsViewModel,
        config_port: ConfigPortAdapter,
        config_loader: ConfigLoader,
    ) -> None:
        vm.set_field_batch(
            "text",
            {
                "remove_numbering": False,
                "add_numbering": True,
                "default_scheme": "legal_standard",
                "numbering_schemes": {
                    "settings": {"default_scheme": "legal_standard", "order": ["legal_standard"]},
                    "number_styles": {
                        "legal_custom": {
                            "name": "Legal Custom",
                            "description": "custom style",
                        }
                    },
                    "schemes": {
                        "legal_standard": {
                            "name": "Legal",
                            "enabled": True,
                            "is_system": False,
                            "level_1": {"format": "第{1.chinese_lower}章 "},
                        }
                    },
                },
                "numbering_clean_rules": {
                    "settings": {"order": ["r1"]},
                    "rules": [
                        {
                            "id": "r1",
                            "enabled": True,
                            "pattern": r"^[0-9]+\. ",
                            "description": "dot",
                            "level": 1,
                        }
                    ],
                },
            },
        )

        assert vm.apply_settings() is True

        snapshot = config_port.snapshot()
        assert snapshot["text"]["remove_numbering"] is False
        assert snapshot["text"]["add_numbering"] is True
        assert snapshot["text"]["numbering_scheme"] == "legal_standard"
        assert snapshot["numbering"]["add"]["settings"]["default_scheme"] == "legal_standard"
        assert snapshot["numbering"]["add"]["number_styles"]["legal_custom"]["name"] == "Legal Custom"
        # Clean rules were persisted through the same injected loader.
        saved = config_loader.get_file_dict("numbering/cleanup.toml")
        rule_ids = [r.get("id") for r in saved.get("rules", []) if isinstance(r, dict)]
        assert "r1" in rule_ids

    def test_apply_persists_field_processor_enabled_flags(
        self, vm: SettingsViewModel, config_port: ConfigPortAdapter
    ) -> None:
        vm.set_field_processor_enabled("gongwen", False)

        assert vm.apply_settings() is True

        snapshot = config_port.snapshot()
        assert snapshot["field_processors"]["processors"]["gongwen"]["enabled"] is False

    def test_dynamic_non_mapping_text_drafts_degrade_without_apply_failure(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_TEXT, "field_processors", "invalid dynamic value")
        vm.set_field(SECTION_TEXT, "numbering_schemes", ["invalid dynamic value"])

        assert vm.get_available_field_processors() == []
        assert vm.set_field_processor_enabled("gongwen", False) is False
        assert vm.apply_settings() is True

    def test_cancel_restores_snapshot(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert vm.config.gui.theme == "dark"
        vm.cancel()
        assert vm.config.gui.theme == "light"  # back to snapshot
        assert vm.is_dirty is False

    def test_cancel_emits_reload(self, vm: SettingsViewModel) -> None:
        reload_signals: list = []
        vm.config_reloaded.connect(lambda: reload_signals.append(True))
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.cancel()
        assert len(reload_signals) == 1


class TestReset:
    def test_reset_section_reverts_to_defaults(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_FORMATTING, "body_format", "discard")
        assert vm.config.formatting.body_format == "discard"
        vm.reset_section(SECTION_FORMATTING)
        assert vm.config.formatting.body_format == "preserve"  # default

    def test_reset_section_emits_status(self, vm: SettingsViewModel) -> None:
        status_signals: list = []
        vm.status_changed.connect(lambda msg, err: status_signals.append((msg, err)))
        vm.reset_section(SECTION_GUI)
        assert len(status_signals) == 1
        assert "reset" in status_signals[0][0].lower()

    def test_reset_all_clears_all(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.set_field(SECTION_LINK, "max_depth", 10)
        vm.reset_all()
        assert vm.is_dirty is False
        assert vm.config.gui.theme == "light"
        assert vm.config.link.max_depth == 3

    def test_reset_text_preserves_user_authored_editors(self, vm: SettingsViewModel) -> None:
        """Reset Text defaults without discarding editor-owned draft data."""
        # Verify the section maps to a group
        group = SettingsViewModel._SECTION_GROUP_MAP.get(SECTION_TEXT)  # type: ignore[attr-defined]
        assert group is not None

        from docwen_runtime.config.registry import reset_plan_for_group

        plan = reset_plan_for_group(group)
        assert "numbering/cleanup.toml" not in plan.files
        assert "field_processors.toml" not in plan.files

        # Set some custom clean rules to make them non-default
        vm.set_field(
            SECTION_TEXT,
            "numbering_clean_rules",
            {
                "settings": {"order": ["my_rule"]},
                "rules": [{"id": "my_rule", "pattern": "^test", "enabled": True, "description": "custom"}],
            },
        )
        vm.set_field(
            SECTION_TEXT,
            "field_processors",
            {
                "settings": {"order": ["my_processor"]},
                "processors": {"my_processor": {"enabled": False}},
            },
        )

        # Verify before reset
        rules_before = vm.config.text.numbering_clean_rules or {}
        rule_ids_before = {r["id"] for r in rules_before.get("rules", [])}
        assert "my_rule" in rule_ids_before

        # Reset text section
        ok = vm.reset_section(SECTION_TEXT)
        assert ok is True

        # User-authored editor drafts are outside the destructive reset plan.
        rules_after = vm.config.text.numbering_clean_rules or {}
        rule_ids_after = {r["id"] for r in rules_after.get("rules", [])}
        assert "my_rule" in rule_ids_after
        assert vm.config.text.field_processors["processors"]["my_processor"]["enabled"] is False
        assert vm.config.dirty_sections == frozenset({SECTION_TEXT})


# ── Dirty state granularity ─────────────────────────────────────────────


class TestDirtyState:
    def test_dirty_sections_tracks_changes(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")
        assert "output" in vm.config.dirty_sections

    def test_dirty_cleared_after_apply(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_settings()
        assert len(vm.config.dirty_sections) == 0

    def test_dirty_cleared_after_cancel(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.cancel()
        assert len(vm.config.dirty_sections) == 0


# ── Copy semantics ──────────────────────────────────────────────────────


class TestCopySemantics:
    def test_config_is_copy_not_ref(self, vm: SettingsViewModel) -> None:
        """Config property should return a deep copy, not a reference."""
        config1 = vm.config
        config1.gui.theme = "dark"
        # Mutating the copy should not affect the ViewModel
        assert vm.config.gui.theme == "light"

    def test_get_section_returns_copy(self, vm: SettingsViewModel) -> None:
        gui1 = vm.get_section(SECTION_GUI)
        gui1.theme = "dark"
        assert vm.config.gui.theme == "light"


# ── Full config replacement ─────────────────────────────────────────────


class TestLoadFullConfig:
    def test_load_full_config_replaces_entire_state(self, vm: SettingsViewModel) -> None:
        from docwen_gui.models.settings_config import GUIConfig, SettingsConfig

        new_gui = GUIConfig(theme="dark", language="en_US")
        new_config = SettingsConfig(gui=new_gui)
        vm.load_full_config(new_config)
        assert vm.config.gui.theme == "dark"
        assert vm.config.gui.language == "en_US"
        assert vm.is_dirty is False

    def test_load_full_config_emits_reload(self, vm: SettingsViewModel) -> None:
        reload_signals: list = []
        vm.config_reloaded.connect(lambda: reload_signals.append(True))
        vm.load_full_config(vm.config)
        assert len(reload_signals) == 1


# ── Heading Numbering Render Mode ───────────────────────────────────────


class TestHeadingNumberingRenderMode:
    def test_default_value(self, vm: SettingsViewModel) -> None:
        assert vm.config.text.heading_numbering_render_mode == "text"

    def test_set_and_get(self, vm: SettingsViewModel) -> None:
        vm.set_field("text", "heading_numbering_render_mode", "word_native")
        assert vm.config.text.heading_numbering_render_mode == "word_native"

    def test_persist_round_trip(self, vm: SettingsViewModel, config_port: ConfigPortAdapter) -> None:
        vm.set_field("text", "heading_numbering_render_mode", "word_native")
        assert vm.apply_settings() is True

        snapshot = config_port.snapshot()
        assert snapshot["text"]["heading_numbering_render_mode"] == "word_native"
