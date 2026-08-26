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

    def test_batch_persistence_fails_closed_without_set_many(self) -> None:
        set_calls: list[tuple[str, object]] = []

        class LegacySetOnlyPort:
            def set(self, key: str, value: object) -> bool:
                set_calls.append((key, value))
                return True

        port = LegacySetOnlyPort()

        assert (
            SettingsViewModel._persist_config_values(  # pyright: ignore[reportPrivateUsage]
                port,
                {
                    "gui.theme.default_theme": "dark",
                    "output.directory.mode": "custom",
                },
            )
            is False
        )
        assert set_calls == []

    def test_partial_apply_keeps_only_unpersisted_draft_dirty(self) -> None:
        from copy import deepcopy

        class PartialPort:
            def __init__(self) -> None:
                self.raw: dict[str, object] = {
                    "gui": {
                        "window": {
                            "remember_gui_state": True,
                            "auto_center": False,
                            "expand_side_panels": False,
                        }
                    },
                    "output": {"directory": {"mode": "source"}},
                }
                self.reload_count = 0

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def set_many(self, values: dict[str, object]) -> bool:
                window = self.raw["gui"]["window"]  # type: ignore[index]
                for key, value in values.items():
                    if key.startswith("gui.window."):
                        window[key.rsplit(".", 1)[-1]] = value  # type: ignore[index]
                return False

            def reload(self) -> None:
                self.reload_count += 1

        port = PartialPort()
        controller = ApplicationController(config_port=port)  # type: ignore[arg-type]
        vm = SettingsViewModel(controller=controller)
        statuses: list[tuple[str, bool]] = []
        vm.status_changed.connect(lambda message, is_error: statuses.append((message, is_error)))
        vm.begin_session()
        vm.set_field(SECTION_GUI, "remember_gui_state", False)
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")

        assert vm.apply_changes() is False

        assert port.reload_count == 1
        assert vm.config.gui.remember_gui_state is False
        assert vm.config.output.output_mode == "custom"
        assert vm.persisted_config.gui.remember_gui_state is False
        assert vm.persisted_config.output.output_mode == "source"
        assert vm.is_dirty is True
        assert {change["field"] for change in vm.get_change_summary()} == {"output.output_mode"}
        assert statuses[-1][1] is True

        vm.cancel_changes()
        assert vm.config.gui.remember_gui_state is False
        assert vm.config.output.output_mode == "source"

    @pytest.mark.parametrize(
        ("operation", "raises", "mutates"),
        (
            ("reset_group", False, True),
            ("reset_group", True, True),
            ("reset_all", False, True),
            ("reset_all", True, True),
            ("reset_group", False, False),
            ("reset_group", True, False),
            ("reset_all", False, False),
            ("reset_all", True, False),
        ),
    )
    def test_failed_reset_preserves_draft_or_reloads_partial_source(
        self,
        operation: str,
        raises: bool,
        mutates: bool,
    ) -> None:
        from copy import deepcopy

        class PartialResetPort:
            def __init__(self) -> None:
                self.raw: dict[str, object] = {
                    "gui": {
                        "window": {
                            "remember_gui_state": False,
                            "auto_center": True,
                            "expand_side_panels": True,
                        }
                    }
                }
                self.reload_count = 0

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def _partial_reset(self) -> bool:
                if mutates:
                    self.raw["gui"] = {
                        "window": {
                            "remember_gui_state": True,
                            "auto_center": False,
                            "expand_side_panels": False,
                        }
                    }
                if raises:
                    raise OSError("simulated later reset failure")
                return False

            def reset_group(self, _group: str) -> bool:
                return self._partial_reset()

            def reset_all(self) -> bool:
                return self._partial_reset()

            def reload(self) -> None:
                self.reload_count += 1

        port = PartialResetPort()
        controller = ApplicationController(config_port=port)  # type: ignore[arg-type]
        vm = SettingsViewModel(controller=controller)
        statuses: list[tuple[str, bool]] = []
        vm.status_changed.connect(lambda message, is_error: statuses.append((message, is_error)))
        vm.begin_session()
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")
        if not mutates:
            vm.set_field(SECTION_GUI, "remember_gui_state", True)

        result = vm.reset_group("general") if operation == "reset_group" else vm.reset_all()

        assert result is False
        assert port.reload_count == 1
        assert statuses[-1][1] is True
        if mutates:
            assert vm.config.gui.remember_gui_state is True
            assert vm.config.gui.auto_center is False
            assert vm.config.gui.expand_side_panels is False
            if operation == "reset_group":
                assert vm.config.output.output_mode == "custom"
                assert vm.persisted_config.output.output_mode == "source"
                assert vm.is_dirty is True
                assert {change["field"] for change in vm.get_change_summary()} == {"output.output_mode"}
            else:
                assert vm.config.output.output_mode == "source"
                assert vm.is_dirty is False
        else:
            assert vm.config.gui.remember_gui_state is True
            assert vm.config.output.output_mode == "custom"
            assert vm.is_dirty is True
            assert {change["field"] for change in vm.get_change_summary()} == {
                "gui.remember_gui_state",
                "output.output_mode",
            }
            vm.cancel_changes()
            assert vm.config.gui.remember_gui_state is False
            assert vm.config.output.output_mode == "source"

    @pytest.mark.parametrize("source_changes", (False, True))
    def test_successful_group_reset_preserves_non_owner_drafts(
        self,
        source_changes: bool,
    ) -> None:
        from copy import deepcopy

        class SuccessfulGroupResetPort:
            def __init__(self) -> None:
                window = {
                    "remember_gui_state": False,
                    "auto_center": True,
                    "expand_side_panels": True,
                }
                if not source_changes:
                    window = {
                        "remember_gui_state": True,
                        "auto_center": False,
                        "expand_side_panels": False,
                    }
                self.raw: dict[str, object] = {
                    "gui": {
                        "window": window,
                        "template": {"md_default_template": "docx"},
                    },
                    "output": {"directory": {"mode": "source"}},
                }

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def reset_group(self, group: str) -> bool:
                assert group == "general"
                if source_changes:
                    self.raw["gui"]["window"] = {  # type: ignore[index]
                        "remember_gui_state": True,
                        "auto_center": False,
                        "expand_side_panels": False,
                    }
                return True

            def reload(self) -> None:
                return None

        port = SuccessfulGroupResetPort()
        controller = ApplicationController(config_port=port)  # type: ignore[arg-type]
        vm = SettingsViewModel(controller=controller)
        vm.begin_session()
        vm.set_field(SECTION_GUI, "auto_center", True)
        vm.set_field(SECTION_GUI, "md_default_template", "xlsx")
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")

        assert vm.reset_group("general") is True

        assert vm.config.gui.remember_gui_state is True
        assert vm.config.gui.auto_center is False
        assert vm.config.gui.expand_side_panels is False
        assert vm.config.gui.md_default_template == "xlsx"
        assert vm.config.output.output_mode == "custom"
        assert vm.persisted_config.gui.md_default_template == "docx"
        assert vm.persisted_config.output.output_mode == "source"
        assert vm.is_dirty is True
        assert {change["field"] for change in vm.get_change_summary()} == {
            "gui.md_default_template",
            "output.output_mode",
        }
        assert vm.config.dirty_sections == frozenset({SECTION_GUI, SECTION_OUTPUT})

        vm.cancel_changes()
        assert vm.config.gui.auto_center is False
        assert vm.config.gui.md_default_template == "docx"
        assert vm.config.output.output_mode == "source"

    def test_successful_proofread_reset_preserves_protected_dictionary_draft(self) -> None:
        from copy import deepcopy

        from docwen_gui.view_models.settings_vm import SECTION_PROOFREAD

        class SuccessfulProofreadResetPort:
            def __init__(self) -> None:
                self.raw: dict[str, object] = {
                    "proofread": {
                        "engine": {
                            "enable_symbol_pairing": True,
                            "enable_symbol_correction": True,
                            "enable_typos_rule": True,
                            "enable_sensitive_word": True,
                        },
                        "skip": {"code_blocks": True, "quote_blocks": False},
                        "symbol_map": {"entries": {"persisted": ["value"]}},
                    }
                }

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def reset_group(self, group: str) -> bool:
                assert group == "proofread"
                return True

            def reload(self) -> None:
                return None

        port = SuccessfulProofreadResetPort()
        vm = SettingsViewModel(
            controller=ApplicationController(config_port=port),  # type: ignore[arg-type]
        )
        vm.begin_session()
        vm.set_field(SECTION_PROOFREAD, "symbol_pairing", False)
        vm.set_field(SECTION_PROOFREAD, "symbol_mappings", {"draft": ["value"]})

        assert vm.reset_group("proofread") is True

        assert vm.config.proofread.symbol_pairing is True
        assert vm.config.proofread.symbol_mappings == {"draft": ["value"]}
        assert vm.persisted_config.proofread.symbol_mappings == {"persisted": ["value"]}
        assert {change["field"] for change in vm.get_change_summary()} == {"proofread.symbol_mappings"}

    def test_successful_document_reset_preserves_sibling_conversion_drafts(self) -> None:
        from copy import deepcopy

        from docwen_gui.view_models.settings_vm import SECTION_FORMATTING, SECTION_SOFTWARE_PRIORITY

        class SuccessfulDocumentResetPort:
            def __init__(self) -> None:
                self.raw: dict[str, object] = {
                    "document": {"to_md_keep_images": True},
                    "spreadsheet": {"to_md_keep_images": True},
                    "software": {
                        "default_priority": {
                            "word_processors": ["wps_writer", "msoffice_word", "libreoffice"],
                            "spreadsheet_processors": ["wps_spreadsheets", "msoffice_excel", "libreoffice"],
                        }
                    },
                }

            def snapshot(self) -> dict[str, object]:
                return deepcopy(self.raw)

            def reset_group(self, group: str) -> bool:
                assert group == "document"
                return True

            def reload(self) -> None:
                return None

        port = SuccessfulDocumentResetPort()
        vm = SettingsViewModel(
            controller=ApplicationController(config_port=port),  # type: ignore[arg-type]
        )
        vm.begin_session()
        vm.set_conversion_default("document", "to_md_keep_images", False)
        vm.set_conversion_default("spreadsheet", "to_md_keep_images", False)
        vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            "word_processors",
            ["libreoffice", "msoffice_word", "wps_writer"],
        )
        vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            "spreadsheet_processors",
            ["libreoffice", "msoffice_excel", "wps_spreadsheets"],
        )
        vm.set_field(SECTION_FORMATTING, "body_format", "discard")

        assert vm.reset_group("document") is True

        assert vm.config.conversion_defaults.document["to_md_keep_images"] is True
        assert vm.config.conversion_defaults.spreadsheet["to_md_keep_images"] is False
        assert vm.config.software_priority.word_processors == ["wps_writer", "msoffice_word", "libreoffice"]
        assert vm.config.software_priority.spreadsheet_processors == [
            "libreoffice",
            "msoffice_excel",
            "wps_spreadsheets",
        ]
        assert vm.config.formatting.body_format == "discard"
        assert {change["field"] for change in vm.get_change_summary()} == {
            "conversion_defaults.spreadsheet",
            "software_priority.spreadsheet_processors",
            "formatting.body_format",
        }

    def test_reset_tab_draft_ownership_covers_every_dialog_group(self) -> None:
        from docwen_gui.view_models.settings_vm import _RESET_GROUP_DRAFT_PATHS
        from docwen_gui.widgets.settings.dialog import RESET_GROUPS
        from docwen_runtime.config.registry import all_specs, reset_plan_for_group

        dialog_groups = set(RESET_GROUPS.values())
        assert dialog_groups <= set(_RESET_GROUP_DRAFT_PATHS)
        for group in dialog_groups:
            assert _RESET_GROUP_DRAFT_PATHS[group]
            plan = reset_plan_for_group(group)
            assert plan.files or plan.dotted_keys

        runtime_groups = {group for spec in all_specs() for group in spec.groups}
        assert runtime_groups <= set(_RESET_GROUP_DRAFT_PATHS)

        assert ("gui", "md_default_template") not in _RESET_GROUP_DRAFT_PATHS["general"]
        assert ("gui", "md_default_template") in _RESET_GROUP_DRAFT_PATHS["text"]
        assert ("proofread", "symbol_mappings") not in _RESET_GROUP_DRAFT_PATHS["proofread"]
        for group in ("document", "spreadsheet", "layout", "other"):
            assert ("conversion_defaults",) not in _RESET_GROUP_DRAFT_PATHS[group]
