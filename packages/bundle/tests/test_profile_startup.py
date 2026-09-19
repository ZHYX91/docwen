"""Unusable profile selection stops bootstrap with an actionable error."""

import pytest


@pytest.mark.contract
def test_cli_rejects_an_invalid_profile_without_loading_default_config(tmp_path, monkeypatch, capsys):
    from docwen_bundle.cli_entry import main

    selected = tmp_path / "selected"
    selected.write_bytes(b"file")
    monkeypatch.setenv("DOCWEN_DATA_DIR", str(selected))
    monkeypatch.delenv("DOCWEN_CONFIG_DIR", raising=False)
    assert main(["--version"]) == 4
    captured = capsys.readouterr()
    assert not captured.out
    assert "Cannot open the selected DocWen profile" in captured.err
    assert str(selected) in captured.err
    assert selected.read_bytes() == b"file"


@pytest.mark.gui
def test_gui_shows_invalid_profile_before_composing_the_window(qapp, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QMessageBox

    from docwen_bundle.gui_entry import main

    selected = tmp_path / "selected"
    selected.write_bytes(b"file")
    monkeypatch.setenv("DOCWEN_DATA_DIR", str(selected))
    monkeypatch.delenv("DOCWEN_CONFIG_DIR", raising=False)
    messages = []
    monkeypatch.setattr(QMessageBox, "critical", lambda _parent, _title, message: messages.append(message))
    assert main([]) == 4
    assert len(messages) == 1 and str(selected) in messages[0]
    assert selected.read_bytes() == b"file"
