"""Focused tests split from test_settings_vm_preview_state.py."""

from __future__ import annotations

from ._settings_vm_preview_state_support import (
    SECTION_GUI,
    SECTION_OUTPUT,
    ApplicationController,
    SettingsViewModel,
    pytest,
)
from ._settings_vm_preview_state_support import (
    config_port as config_port,
)

pytestmark = pytest.mark.unit
from ._settings_vm_preview_state_support import (
    vm as vm,
)


class TestPartialPersistenceFailure:
    """A failed multi-file Apply must reconcile its partial source honestly."""

    @pytest.mark.parametrize(
        ("group", "owner_path", "draft_value", "persisted_value"),
        (
            ("general", ("gui", "theme"), "dark", "light"),
            ("text", ("text", "remove_numbering"), False, True),
            ("proofread", ("proofread", "symbol_pairing"), False, True),
            (
                "document",
                ("conversion_defaults", "document", "to_md_keep_images"),
                False,
                True,
            ),
            (
                "spreadsheet",
                ("conversion_defaults", "spreadsheet", "to_md_keep_images"),
                False,
                True,
            ),
            (
                "image",
                ("conversion_defaults", "image", "to_md_keep_images"),
                False,
                True,
            ),
            (
                "layout",
                ("conversion_defaults", "layout", "to_md_keep_images"),
                False,
                True,
            ),
            ("link", ("link", "max_depth"), 10, 3),
            ("formatting", ("formatting", "body_format"), "discard", "preserve"),
            ("output", ("output", "output_mode"), "custom", "source"),
            ("export", ("export", "image_mode"), "base64", "file"),
            ("logging", ("logging", "level"), "warning", "debug"),
            (
                "other",
                ("conversion_defaults", "other", "to_md_keep_images"),
                False,
                True,
            ),
        ),
    )
    def test_successful_builtin_group_reset_discards_only_owner_draft(
        self,
        group: str,
        owner_path: tuple[str, ...],
        draft_value: object,
        persisted_value: object,
    ) -> None:
        from copy import deepcopy

        from docwen_gui.view_models.settings_vm import _read_model_path

        class SuccessfulNoopGroupResetPort:
            def __init__(self) -> None:
                self.raw: dict[str, object] = {
                    "gui": {"theme": {"default_theme": "light"}},
                    "text": {"remove_numbering": True},
                    "proofread": {
                        "engine": {"enable_symbol_pairing": True},
                        "symbol_map": {"entries": {"persisted": ["value"]}},
                    },
                    "document": {"to_md_keep_images": True},
                    "spreadsheet": {"to_md_keep_images": True},
                    "image": {"to_md_keep_images": True},
                    "layout": {"to_md_keep_images": True},
                    "other": {"to_md_keep_images": True},
                    "link": {"embedding": {"max_depth": 3}},
                    "conversion": {"docx_to_md": {"preserve_formatting": True}},
                    "output": {"directory": {"mode": "source"}},
                    "export": {"to_md_image_extraction_mode": "file"},
                    "logger": {"level": "debug"},
                }

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def reset_group(self, requested_group: str) -> bool:
                assert requested_group == group
                return True

            def reload(self) -> None:
                return None

        port = SuccessfulNoopGroupResetPort()
        vm = SettingsViewModel(
            controller=ApplicationController(config_port=port),  # type: ignore[arg-type]
        )
        vm.begin_session()
        if owner_path[0] == "conversion_defaults":
            vm.set_conversion_default(owner_path[1], owner_path[2], draft_value)
        else:
            vm.set_field(owner_path[0], owner_path[1], draft_value)
        if group == "output":
            unrelated_path = ("gui", "theme")
            unrelated_draft = "dark"
            unrelated_persisted = "light"
        else:
            unrelated_path = ("output", "output_mode")
            unrelated_draft = "custom"
            unrelated_persisted = "source"
        vm.set_field(unrelated_path[0], unrelated_path[1], unrelated_draft)

        assert _read_model_path(vm.config, owner_path) == draft_value
        assert _read_model_path(vm.config, unrelated_path) == unrelated_draft

        assert vm.reset_group(group) is True

        assert _read_model_path(vm.config, owner_path) == persisted_value
        assert _read_model_path(vm.config, unrelated_path) == unrelated_draft
        assert _read_model_path(vm.persisted_config, unrelated_path) == unrelated_persisted
        assert {change["field"] for change in vm.get_change_summary()} == {".".join(unrelated_path)}
        assert vm.config.dirty_sections == frozenset({unrelated_path[0]})
        assert vm.is_dirty is True

        vm.cancel_changes()
        assert _read_model_path(vm.config, owner_path) == persisted_value
        assert _read_model_path(vm.config, unrelated_path) == unrelated_persisted
        assert vm.is_dirty is False
        assert vm.get_change_summary() == []

    def test_successful_unknown_group_uses_conservative_full_refresh(self) -> None:
        from copy import deepcopy

        class ExtendedConfigPort:
            def __init__(self) -> None:
                self.raw: dict[str, object] = {
                    "gui": {"theme": {"default_theme": "light"}},
                    "output": {"directory": {"mode": "source"}},
                }

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def reset_group(self, group: str) -> bool:
                assert group == "vendor_group"
                return True

            def reload(self) -> None:
                return None

        port = ExtendedConfigPort()
        vm = SettingsViewModel(
            controller=ApplicationController(config_port=port),  # type: ignore[arg-type]
        )
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")

        assert vm.reset_group("vendor_group") is True

        assert vm.config.gui.theme == "light"
        assert vm.config.output.output_mode == "source"
        assert vm.persisted_config == vm.config
        assert vm.is_dirty is False

    def test_conversion_defaults_aggregate_resets_owned_text_and_export_leaves(self) -> None:
        from copy import deepcopy

        from docwen_gui.view_models.settings_vm import (
            SECTION_CONVERSION_DEFAULTS,
            SECTION_EXPORT,
            SECTION_FORMATTING,
            SECTION_TEXT,
        )

        class SuccessfulAggregateResetPort:
            def __init__(self) -> None:
                self.raw: dict[str, object] = {
                    "text": {
                        "remove_numbering": True,
                        "add_numbering": False,
                        "numbering_scheme": "hierarchical_standard",
                        "heading_numbering_render_mode": "text",
                    },
                    "export": {
                        "to_md_image_extraction_mode": "file",
                        "to_md_ocr_placement_mode": "image_md",
                    },
                    "conversion": {
                        "export": {"base64_compress_enabled": True},
                    },
                }

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def reset_group(self, group: str) -> bool:
                assert group == "conversion_defaults"
                return True

            def reload(self) -> None:
                return None

        port = SuccessfulAggregateResetPort()
        vm = SettingsViewModel(
            controller=ApplicationController(config_port=port),  # type: ignore[arg-type]
        )
        vm.begin_session()
        vm.set_field(SECTION_TEXT, "remove_numbering", False)
        vm.set_field(SECTION_EXPORT, "image_mode", "base64")
        vm.set_field(SECTION_EXPORT, "base64_compress_enabled", False)
        vm.set_field(SECTION_FORMATTING, "table_style_mode", "custom")
        vm.set_field(SECTION_FORMATTING, "custom_table_style_name", "Draft Table")
        vm.set_field(SECTION_FORMATTING, "body_format", "discard")
        vm.set_field(SECTION_GUI, "theme", "dark")

        assert vm.reset_section(SECTION_CONVERSION_DEFAULTS) is True

        assert vm.config.text.remove_numbering is True
        assert vm.config.export.image_mode == "file"
        assert vm.config.export.base64_compress_enabled is False
        assert vm.config.formatting.table_style_mode == "builtin"
        assert vm.config.formatting.custom_table_style_name == ""
        assert vm.config.formatting.body_format == "discard"
        assert vm.config.gui.theme == "dark"
        assert {change["field"] for change in vm.get_change_summary()} == {
            "gui.theme",
            "export.base64_compress_enabled",
            "formatting.body_format",
        }

    def test_software_alias_resets_software_priority_draft(self) -> None:
        from copy import deepcopy

        from docwen_gui.view_models.settings_vm import SECTION_SOFTWARE_PRIORITY

        class SuccessfulSoftwareResetPort:
            def __init__(self) -> None:
                self.raw: dict[str, object] = {
                    "software": {
                        "default_priority": {
                            "word_processors": ["wps_writer", "msoffice_word", "libreoffice"],
                        }
                    }
                }

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def reset_group(self, group: str) -> bool:
                assert group == "software"
                return True

            def reload(self) -> None:
                return None

        port = SuccessfulSoftwareResetPort()
        vm = SettingsViewModel(
            controller=ApplicationController(config_port=port),  # type: ignore[arg-type]
        )
        vm.begin_session()
        vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            "word_processors",
            ["libreoffice", "msoffice_word", "wps_writer"],
        )

        assert vm.reset_group("software") is True

        assert vm.config.software_priority.word_processors == ["wps_writer", "msoffice_word", "libreoffice"]
        assert vm.is_dirty is False
