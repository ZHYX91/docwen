from __future__ import annotations

from copy import deepcopy
from pathlib import Path

import pytest
from PySide6.QtWidgets import QDialog, QTableWidgetItem

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_runtime.config.loader import ConfigLoader

pytestmark = pytest.mark.gui

PROJECT_CONFIGS = Path(__file__).resolve().parents[4] / "configs"


def _view_model(port: ConfigPortAdapter) -> SettingsViewModel:
    return SettingsViewModel(controller=ApplicationController(config_port=port))


def test_text_numbering_editors_write_the_injected_config_port(
    tmp_path: Path,
) -> None:
    """Immediate Text editor saves must not consult an unrelated loader."""
    injected_loader = ConfigLoader(
        base_dir=PROJECT_CONFIGS,
        user_dir=tmp_path / "injected",
    )
    injected = ConfigPortAdapter(injected_loader)
    unrelated_loader = ConfigLoader(
        base_dir=PROJECT_CONFIGS,
        user_dir=tmp_path / "unrelated",
    )
    vm = _view_model(injected)

    schemes = {
        "settings": {"default_scheme": "isolated_scheme", "order": ["isolated_scheme"]},
        "number_styles": {},
        "schemes": {
            "isolated_scheme": {
                "name": "Isolated",
                "is_system": False,
                "level_1": {"format": "{1.arabic_half}. "},
            }
        },
    }
    clean_rules = {
        "settings": {"order": ["isolated_rule"]},
        "rules": [
            {
                "id": "isolated_rule",
                "enabled": True,
                "pattern": "^isolated",
                "description": "isolated",
                "level": 1,
                "is_system": False,
            }
        ],
    }

    assert vm.persist_numbering_schemes_source(schemes) is True
    assert vm.persist_numbering_clean_rules_source(clean_rules) is True

    injected_snapshot = injected.snapshot()
    assert injected_snapshot["numbering"]["add"]["settings"]["default_scheme"] == "isolated_scheme"
    assert injected_snapshot["numbering"]["cleanup"]["settings"]["order"] == ["isolated_rule"]
    unrelated_snapshot = unrelated_loader.config.as_dict()
    assert "isolated_scheme" not in unrelated_snapshot["numbering"]["add"]["schemes"]
    assert "isolated_rule" not in unrelated_snapshot["numbering"]["cleanup"]["settings"]["order"]


def test_text_settings_apply_writes_only_the_injected_config_port(
    tmp_path: Path,
) -> None:
    injected_loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "injected")
    injected = ConfigPortAdapter(injected_loader)
    unrelated_loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "unrelated")
    vm = _view_model(injected)
    schemes = deepcopy(vm.config.text.numbering_schemes)
    schemes["settings"]["default_scheme"] = "apply_scheme"
    schemes["settings"]["order"].append("apply_scheme")
    schemes["schemes"]["apply_scheme"] = {
        "name": "Apply Scheme",
        "is_system": False,
        "level_1": {"format": "{1.arabic_half}. "},
    }
    vm.set_field_batch(
        "text",
        {
            "add_numbering": True,
            "default_scheme": "apply_scheme",
            "numbering_schemes": schemes,
        },
    )

    assert vm.apply_settings() is True

    assert injected.get("numbering.add.schemes.apply_scheme.name") == "Apply Scheme"
    assert injected.get("text.numbering_scheme") == "apply_scheme"
    assert unrelated_loader.config.as_dict()["numbering"]["add"]["schemes"].get("apply_scheme") is None


def test_custom_numbering_scheme_deletion_round_trips(
    tmp_path: Path,
) -> None:
    port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    vm = _view_model(port)
    with_custom = deepcopy(vm.config.text.numbering_schemes)
    with_custom["settings"]["order"].append("deletable_scheme")
    with_custom["schemes"]["deletable_scheme"] = {
        "name": "Deletable",
        "is_system": False,
        "level_1": {"format": "{1.arabic_half}. "},
    }
    assert vm.persist_numbering_schemes_source(with_custom) is True
    assert port.get("numbering.add.schemes.deletable_scheme.name") == "Deletable"

    without_custom = deepcopy(vm.config.text.numbering_schemes)
    without_custom["settings"]["order"].remove("deletable_scheme")
    del without_custom["schemes"]["deletable_scheme"]

    assert vm.persist_numbering_schemes_source(without_custom) is True
    assert port.get("numbering.add.schemes.deletable_scheme") is None


def test_proofread_editor_reads_effective_base_data_through_the_injected_port(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A sparse user layer must not make shipped pairing rows appear empty."""
    from docwen_gui.widgets.settings import proofread_tab

    injected_loader = ConfigLoader(
        base_dir=PROJECT_CONFIGS,
        user_dir=tmp_path / "injected",
    )
    injected = ConfigPortAdapter(injected_loader)
    tab = proofread_tab.ProofreadTab(_view_model(injected))
    row_counts: list[int] = []

    def inspect_without_modal(dialog: proofread_tab._SymbolMappingEditor) -> int:
        row_counts.append(dialog._table.rowCount())  # pyright: ignore[reportPrivateUsage]
        dialog.reject()
        return int(QDialog.DialogCode.Rejected)

    monkeypatch.setattr(proofread_tab._SymbolMappingEditor, "exec", inspect_without_modal)

    tab._open_symbol_mapping_editor()  # pyright: ignore[reportPrivateUsage]

    assert row_counts and row_counts[0] > 0
    tab.close()


def test_proofread_editor_writes_the_injected_port_not_an_unrelated_loader(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings import proofread_tab

    injected_loader = ConfigLoader(
        base_dir=PROJECT_CONFIGS,
        user_dir=tmp_path / "injected",
    )
    injected = ConfigPortAdapter(injected_loader)
    unrelated_loader = ConfigLoader(
        base_dir=PROJECT_CONFIGS,
        user_dir=tmp_path / "unrelated",
    )
    tab = proofread_tab.ProofreadTab(_view_model(injected))

    def edit_without_modal(dialog: proofread_tab._TyposDictionaryEditor) -> int:
        dialog._add_row()  # pyright: ignore[reportPrivateUsage]
        row = dialog._table.rowCount() - 1  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 0, QTableWidgetItem("isolated_typo"))  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 1, QTableWidgetItem("isolated_fix"))  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 2, QTableWidgetItem("isolated remark"))  # pyright: ignore[reportPrivateUsage]
        dialog._on_accept()  # pyright: ignore[reportPrivateUsage]
        return int(QDialog.DialogCode.Accepted)

    monkeypatch.setattr(proofread_tab._TyposDictionaryEditor, "exec", edit_without_modal)

    tab._open_typos_editor()  # pyright: ignore[reportPrivateUsage]

    assert injected.get("proofread.typos.entries.isolated_typo") == ["isolated_fix"]
    assert unrelated_loader.config.as_dict()["proofread"]["typos"]["entries"].get("isolated_typo") is None
    tab.close()


def test_proofread_editor_save_survives_later_settings_apply(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Applying stale dialog draft state must not erase an immediate dictionary save."""
    from docwen_gui.widgets.settings import proofread_tab

    port = ConfigPortAdapter(
        base_dir=PROJECT_CONFIGS,
        user_dir=tmp_path / "configs",
    )
    vm = _view_model(port)
    tab = proofread_tab.ProofreadTab(vm)

    def edit_without_modal(dialog: proofread_tab._TyposDictionaryEditor) -> int:
        dialog._add_row()  # pyright: ignore[reportPrivateUsage]
        row = dialog._table.rowCount() - 1  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 0, QTableWidgetItem("surviving_typo"))  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 1, QTableWidgetItem("surviving_fix"))  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 2, QTableWidgetItem("surviving remark"))  # pyright: ignore[reportPrivateUsage]
        dialog._on_accept()  # pyright: ignore[reportPrivateUsage]
        return int(QDialog.DialogCode.Accepted)

    monkeypatch.setattr(proofread_tab._TyposDictionaryEditor, "exec", edit_without_modal)

    tab._open_typos_editor()  # pyright: ignore[reportPrivateUsage]
    assert port.get("proofread.typos.entries.surviving_typo") == ["surviving_fix"]
    user_file = tmp_path / "configs" / "proofread" / "typos.toml"
    assert "surviving remark" in user_file.read_text(encoding="utf-8")

    assert vm.apply_settings() is True

    assert port.get("proofread.typos.entries.surviving_typo") == ["surviving_fix"]
    assert "surviving remark" in user_file.read_text(encoding="utf-8")
    tab.close()


def test_immediate_proofread_save_updates_cancel_baseline(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings import proofread_tab

    port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    vm = _view_model(port)
    vm.begin_session()
    original_theme = vm.config.gui.theme
    vm.set_field("gui", "theme", "dark" if original_theme != "dark" else "light")
    tab = proofread_tab.ProofreadTab(vm)

    def edit_without_modal(dialog: proofread_tab._TyposDictionaryEditor) -> int:
        dialog._add_row()  # pyright: ignore[reportPrivateUsage]
        row = dialog._table.rowCount() - 1  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 0, QTableWidgetItem("cancel_safe_typo"))  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 1, QTableWidgetItem("cancel_safe_fix"))  # pyright: ignore[reportPrivateUsage]
        dialog._table.setItem(row, 2, QTableWidgetItem("cancel-safe remark"))  # pyright: ignore[reportPrivateUsage]
        dialog._on_accept()  # pyright: ignore[reportPrivateUsage]
        return int(QDialog.DialogCode.Accepted)

    monkeypatch.setattr(proofread_tab._TyposDictionaryEditor, "exec", edit_without_modal)
    tab._open_typos_editor()  # pyright: ignore[reportPrivateUsage]

    vm.cancel_changes()

    assert vm.config.gui.theme == original_theme
    assert vm.config.proofread.typos_dict["cancel_safe_typo"] == ["cancel_safe_fix"]
    assert port.get("proofread.typos.entries.cancel_safe_typo") == ["cancel_safe_fix"]
    tab.close()


def test_proofread_editor_can_delete_a_shipped_dictionary_entry(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings import proofread_tab

    port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    tab = proofread_tab.ProofreadTab(_view_model(port))

    def delete_without_modal(dialog: proofread_tab._SymbolErrorEditor) -> int:
        for row in range(dialog._table.rowCount()):  # pyright: ignore[reportPrivateUsage]
            item = dialog._table.item(row, 0)  # pyright: ignore[reportPrivateUsage]
            if item is not None and item.text() == "0":
                dialog._table.removeRow(row)  # pyright: ignore[reportPrivateUsage]
                break
        dialog._on_accept()  # pyright: ignore[reportPrivateUsage]
        return int(QDialog.DialogCode.Accepted)

    monkeypatch.setattr(proofread_tab._SymbolErrorEditor, "exec", delete_without_modal)
    tab._open_symbol_error_editor()  # pyright: ignore[reportPrivateUsage]

    assert port.get("proofread.symbol_map.entries.0") is None
    assert port.get("proofread.symbol_map.entries.1") is not None
    tab.close()


def test_editor_save_failures_keep_child_dialogs_open(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings.numbering_add_editor import NumberingAddDialog
    from docwen_gui.widgets.settings.numbering_clean_editor import NumberingCleanDialog
    from docwen_gui.widgets.settings.proofread_tab import _TyposDictionaryEditor

    monkeypatch.setattr("PySide6.QtWidgets.QMessageBox.warning", lambda *_args, **_kwargs: None)

    typo = _TyposDictionaryEditor(
        "proofread/typos.toml",
        source_text="[entries]\n",
        save_callback=lambda _name, _content: False,
    )
    typo._add_row()  # pyright: ignore[reportPrivateUsage]
    typo._table.setItem(0, 0, QTableWidgetItem("failed_typo"))  # pyright: ignore[reportPrivateUsage]
    typo._table.setItem(0, 1, QTableWidgetItem("failed_fix"))  # pyright: ignore[reportPrivateUsage]
    typo._on_accept()  # pyright: ignore[reportPrivateUsage]
    assert typo.result() != int(QDialog.DialogCode.Accepted)

    add = NumberingAddDialog(
        config_data={
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "schemes": {"s1": {"name": "S1", "is_system": False}},
        },
        on_save=lambda _data: False,
    )
    assert add._save_to_disk() is False  # pyright: ignore[reportPrivateUsage]
    assert add.result() != int(QDialog.DialogCode.Accepted)

    clean = NumberingCleanDialog(
        config_data={
            "settings": {"order": ["r1"]},
            "rules": [{"id": "r1", "enabled": True, "pattern": "^x", "is_system": True}],
        },
        on_save=lambda _data: False,
    )
    assert clean._save() is False  # pyright: ignore[reportPrivateUsage]
    assert clean.result() != int(QDialog.DialogCode.Accepted)

    typo.close()
    add.close()
    clean.close()


def test_settings_vm_rejects_unowned_editor_file_before_writing(tmp_path: Path) -> None:
    user_dir = tmp_path / "configs"
    port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    vm = _view_model(port)

    assert vm.read_config_file_text("gui.toml") is None
    assert vm.save_config_file_text("gui.toml", '[window]\ndefault_mode = "batch"\n') is False
    assert not (user_dir / "gui.toml").exists()


def test_proofread_ad_hoc_fallback_uses_runtime_atomic_writer(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core.toml_tools import read_toml_text
    from docwen_gui.widgets.settings import proofread_tab

    target = tmp_path / "typos.toml"
    writes: list[tuple[Path, str]] = []

    def atomic_write(path: str | Path, content: str) -> None:
        writes.append((Path(path), content))

    monkeypatch.setattr(proofread_tab, "atomic_write_text", atomic_write)
    doc = read_toml_text('[entries]\ntypo = ["fix"]\n')

    assert (
        proofread_tab._save_toml_document(  # pyright: ignore[reportPrivateUsage]
            doc,
            config_name="proofread/typos.toml",
            fallback_path=target,
            save_callback=None,
        )
        is True
    )
    assert writes == [(target, '[entries]\ntypo = ["fix"]\n')]
    assert not target.exists()
