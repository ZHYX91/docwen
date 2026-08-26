"""Focused tests split from test_action_area_vm.py."""

from __future__ import annotations

import pytest

from ._action_area_vm_support import (
    ActionAreaViewModel,
    FakeMainWindowViewModel_for_port,
    _write_minimal_base_config_tree,
)

pytestmark = pytest.mark.unit
from ._action_area_vm_support import (
    vm as vm,
)


class TestSettingsActionAreaNumberingSync:
    """SettingsViewModel persist -> ActionArea reads the persisted default.

    Phase B-GUI real-effect path: when the user changes a numbering
    default in the settings editor and applies it, a subsequently
    constructed ActionAreaViewModel must read that persisted value as
    its default (not the hardcoded fallback). This is the cross-VM
    contract the plan calls out: '编辑器保存后下游 ActionArea 默认值
    同步；reload 后再次打开仍一致'.

    Uses a real ``ConfigPortAdapter`` over a temp config dir, shared
    between a real ``SettingsViewModel`` (which persists) and the
    ActionArea's ``FakeMainWindowViewModel`` double (which reads).
    """

    def test_render_mode_persisted_reaches_action_area(self, tmp_path) -> None:
        """word_native set + applied in settings -> ActionArea default is word_native."""
        from pathlib import Path

        from docwen_application.controller import ApplicationController
        from docwen_bundle.config_port import ConfigPortAdapter
        from docwen_gui.view_models.settings_vm import SettingsViewModel

        config_dir = Path(tmp_path) / "configs"
        _write_minimal_base_config_tree(config_dir)

        # 1. Settings VM: set render_mode to word_native and apply (persist).
        controller = ApplicationController(config_port=ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir))
        settings_vm = SettingsViewModel(controller=controller)
        settings_vm.set_field("text", "heading_numbering_render_mode", "word_native")
        assert settings_vm.apply_settings() is True

        # 2. Fresh ConfigPortAdapter over the same dir, handed to an
        #    ActionArea via the FakeMainWindowViewModel double.
        shared_port = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
        action_vm = ActionAreaViewModel(
            FakeMainWindowViewModel_for_port(shared_port)  # type: ignore[arg-type]
        )
        action_vm.setup_for_md_to_document("/test.md")

        # 3. The ActionArea default must reflect the persisted value,
        #    not the hardcoded 'text' fallback.
        assert action_vm.md_heading_numbering_render_mode == "word_native", (
            "ActionArea render_mode should reflect the settings-persisted "
            "value 'word_native', got "
            f"{action_vm.md_heading_numbering_render_mode!r}"
        )

    def test_numbering_scheme_persisted_reaches_action_area(self, tmp_path) -> None:
        """scheme set + applied in settings -> ActionArea default matches."""
        from pathlib import Path

        from docwen_application.controller import ApplicationController
        from docwen_bundle.config_port import ConfigPortAdapter
        from docwen_gui.view_models.settings_vm import SettingsViewModel

        config_dir = Path(tmp_path) / "configs"
        _write_minimal_base_config_tree(config_dir)

        controller = ApplicationController(config_port=ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir))
        settings_vm = SettingsViewModel(controller=controller)
        settings_vm.set_field("text", "default_scheme", "legal_standard")
        settings_vm.set_field("text", "add_numbering", True)
        assert settings_vm.apply_settings() is True

        shared_port = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
        action_vm = ActionAreaViewModel(
            FakeMainWindowViewModel_for_port(shared_port)  # type: ignore[arg-type]
        )
        action_vm.setup_for_md_to_document("/test.md")

        assert action_vm.md_numbering_scheme == "legal_standard", (
            "ActionArea numbering_scheme should reflect the settings-"
            f"persisted 'legal_standard', got "
            f"{action_vm.md_numbering_scheme!r}"
        )
        assert action_vm.md_add_numbering is True, (
            f"ActionArea add_numbering should reflect the settings-persisted True, got {action_vm.md_add_numbering!r}"
        )

    def test_reload_keeps_persisted_default_consistent(self, tmp_path) -> None:
        """A fresh settings VM on the same config dir reloads the same value.

        This is the 'reload 后再次打开仍一致' half of the contract: the
        persisted default survives a settings re-open, and the
        ActionArea still reads it consistently afterwards.
        """
        from pathlib import Path

        from docwen_application.controller import ApplicationController
        from docwen_bundle.config_port import ConfigPortAdapter
        from docwen_gui.view_models.settings_vm import SettingsViewModel

        config_dir = Path(tmp_path) / "configs"
        _write_minimal_base_config_tree(config_dir)

        # First settings session: persist word_native.
        controller_a = ApplicationController(config_port=ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir))
        vm_a = SettingsViewModel(controller=controller_a)
        vm_a.set_field("text", "heading_numbering_render_mode", "word_native")
        assert vm_a.apply_settings() is True

        # Second settings session: a fresh VM over the same dir must
        # reload the persisted value, not the 'text' default.
        controller_b = ApplicationController(config_port=ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir))
        vm_b = SettingsViewModel(controller=controller_b)
        assert vm_b.config.text.heading_numbering_render_mode == "word_native", (
            "Fresh SettingsViewModel should reload the persisted render_mode 'word_native'"
        )

        # And the ActionArea, reading the same dir, agrees with the
        # reloaded settings VM.
        shared_port = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
        action_vm = ActionAreaViewModel(
            FakeMainWindowViewModel_for_port(shared_port)  # type: ignore[arg-type]
        )
        action_vm.setup_for_md_to_document("/test.md")
        assert action_vm.md_heading_numbering_render_mode == "word_native"
        assert action_vm.md_heading_numbering_render_mode == vm_b.config.text.heading_numbering_render_mode

    def test_export_markdown_modes_are_the_single_action_area_source(self, tmp_path) -> None:
        """Applied Export modes reach the ActionArea without a section fallback."""
        from pathlib import Path

        from docwen_application.controller import ApplicationController
        from docwen_bundle.config_port import ConfigPortAdapter
        from docwen_gui.view_models.settings_vm import SECTION_EXPORT, SettingsViewModel

        config_dir = Path(tmp_path) / "configs"
        _write_minimal_base_config_tree(config_dir)

        controller = ApplicationController(config_port=ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir))
        settings_vm = SettingsViewModel(controller=controller)
        settings_vm.set_field(SECTION_EXPORT, "image_mode", "base64")
        settings_vm.set_field(SECTION_EXPORT, "ocr_mode", "main_md")
        assert settings_vm.apply_settings() is True

        shared_port = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
        action_vm = ActionAreaViewModel(
            FakeMainWindowViewModel_for_port(shared_port)  # type: ignore[arg-type]
        )
        action_vm.setup_for_spreadsheet_file("/test.xlsx")

        opts = action_vm.collect_options()
        assert opts["image_mode"] == "base64"
        assert opts["ocr_placement"] == "main_md"
