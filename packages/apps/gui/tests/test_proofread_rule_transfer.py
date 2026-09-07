"""Import preview acceptance, cancellation, concurrency, and complete export."""

from pathlib import Path

import pytest
from PySide6.QtWidgets import QDialog, QWidget

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings import proofread_transfer

pytestmark = pytest.mark.gui
PROJECT_CONFIGS = Path(__file__).resolve().parents[4] / "configs"


@pytest.mark.parametrize("outcome", ["cancel", "accept", "changed", "invalid"])
def test_rule_import_changes_disk_only_after_a_valid_unchanged_preview(
    qapp, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, outcome: str
) -> None:
    name = "proofread/typos.toml"
    port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    assert port.save_file_text(name, '[entries]\noriginal = ["keep"]\n')
    vm = SettingsViewModel(controller=ApplicationController(config_port=port))
    parent = QWidget()
    transfer = proofread_transfer.ProofreadRuleTransfer(parent, vm)
    original = vm.read_config_file_text(name)
    incoming = tmp_path / "import.toml"
    incoming.write_text('[entries]\nadded = ["new"]\n' if outcome != "invalid" else "[broken", encoding="utf-8")
    errors: list[Exception] = []
    monkeypatch.setattr(proofread_transfer.QFileDialog, "getOpenFileName", lambda *args: (str(incoming), ""))
    monkeypatch.setattr(transfer, "_error", lambda title, error: errors.append(error))
    monkeypatch.setattr(proofread_transfer.QMessageBox, "information", lambda *args: None)

    def preview(dialog: proofread_transfer.RuleImportDialog) -> int:
        assert dialog.plan.source_text == original
        assert dialog.strategy.currentData() == "merge"
        if outcome == "changed":
            assert port.save_file_text(name, '[entries]\nother_editor = ["retain"]\n')
        return int(QDialog.DialogCode.Rejected if outcome == "cancel" else QDialog.DialogCode.Accepted)

    monkeypatch.setattr(proofread_transfer.RuleImportDialog, "exec", preview)
    transfer.import_rules(name)
    entries = port.snapshot()["proofread"]["typos"]["entries"]
    expected = (
        {"original": ["keep"], "added": ["new"]}
        if outcome == "accept"
        else {"other_editor": ["retain"]}
        if outcome == "changed"
        else {"original": ["keep"]}
    )
    assert entries == expected
    assert bool(errors) == (outcome in {"changed", "invalid"})
    parent.close()


def test_rule_export_writes_complete_effective_source_with_comments(
    qapp, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    name = "proofread/typos.toml"
    port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    assert port.save_file_text(name, '# My rules\n[entries]\nfirst = ["a"] # note\nsecond = ["b"]\n')
    vm = SettingsViewModel(controller=ApplicationController(config_port=port))
    parent = QWidget()
    transfer = proofread_transfer.ProofreadRuleTransfer(parent, vm)
    destination = tmp_path / "all-rules.toml"
    monkeypatch.setattr(proofread_transfer.QFileDialog, "getSaveFileName", lambda *args: (str(destination), ""))
    monkeypatch.setattr(proofread_transfer.QMessageBox, "information", lambda *args: None)
    transfer.export_rules(name)
    assert destination.read_text(encoding="utf-8") == vm.read_config_file_text(name)
    assert "second" in destination.read_text(encoding="utf-8")
    parent.close()
