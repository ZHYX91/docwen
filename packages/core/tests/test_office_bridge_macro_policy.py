"""Real inputs must follow an owned, explicitly macro-disabled backend."""

from collections import deque
from pathlib import Path
from typing import cast
from unittest.mock import MagicMock
from xml.etree import ElementTree

import pytest

from docwen_core import office_bridge
from docwen_core.windows_process import WindowsProcessIdentity

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize("app_type", ["word", "excel", "powerpoint"])
@pytest.mark.parametrize("policy", ["supported", "setter_error", "getter_error", "ignored"])
def test_owned_macro_policy_precedes_real_input_and_cleanup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, app_type: str, policy: str
) -> None:
    events: list[str] = []
    source = tmp_path / "input.csv"
    source.write_text("name,value\nalpha,1\n", encoding="utf-8")
    output = tmp_path / "out.bin"
    document = MagicMock()
    document.SaveAs.side_effect = lambda *args, **kwargs: (events.append("save"), output.write_bytes(b"output"))
    document.Close.side_effect = lambda *args, **kwargs: events.append("close")

    class App:
        def __init__(self) -> None:
            self.Documents = self.Workbooks = self.Presentations = self

        @property
        def AutomationSecurity(self) -> int:
            events.append("get_policy")
            if policy == "getter_error":
                raise AttributeError("getter unavailable")
            return 1 if policy == "ignored" else 3

        @AutomationSecurity.setter
        def AutomationSecurity(self, value: int) -> None:
            assert value == 3
            assert events[-1] == "owned"
            events.append("set_policy")
            if policy == "setter_error":
                raise AttributeError("setter unavailable")

        def Open(self, path: str, *args: object, **kwargs: object) -> MagicMock:
            assert path == str(source)
            assert events[-1] == "get_policy"
            events.append("open")
            return document

        def Quit(self) -> None:
            events.append("quit")

    client = MagicMock()
    client.DispatchEx.return_value = App()
    pythoncom = MagicMock()
    pythoncom.CoUninitialize.side_effect = lambda: events.append("uninitialize")
    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (pythoncom, client))
    monkeypatch.setattr(office_bridge, "_get_com_app_pid", lambda app: 400)
    identity = WindowsProcessIdentity(400, 123)
    monkeypatch.setattr(WindowsProcessIdentity, "capture", lambda *args, **kwargs: identity)
    monkeypatch.setattr(WindowsProcessIdentity, "terminate_if_running", lambda *args, **kwargs: events.append("reap"))
    result = office_bridge._try_com_conversion(
        str(source),
        str(output),
        prog_id="Test.Application",
        save_format=1,
        app_type=cast(office_bridge.AppType, app_type),
        on_process_owned=lambda owned: events.append("owned"),
    )
    assert events[-3:] == ["quit", "reap", "uninitialize"]
    if policy == "supported":
        assert result == str(output)
        assert events[:6] == ["owned", "set_policy", "get_policy", "open", "save", "close"]
    else:
        assert isinstance(result, office_bridge._OfficeSecurityFailure)
        assert result.error_code == "OFFICE_AUTOMATION_SECURITY_UNAVAILABLE"
        assert "open" not in events
        assert not output.exists()


def test_excel_uninspectable_input_is_rejected_before_com_import(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "input.xls"
    source.write_bytes(b"unknown")
    import_com = MagicMock(side_effect=AssertionError("must not start COM"))
    monkeypatch.setattr(office_bridge, "_import_win32", import_com)
    result = office_bridge._try_com_conversion(
        str(source), str(tmp_path / "out.xlsx"), prog_id="KET.Application", save_format=51, app_type="excel"
    )
    assert isinstance(result, office_bridge._OfficeSecurityFailure)
    assert result.error_code == "OFFICE_INPUT_MACRO_POLICY_REJECTED"
    import_com.assert_not_called()


def test_security_refusal_survives_worker_pipe(monkeypatch: pytest.MonkeyPatch) -> None:
    refusal = office_bridge._OfficeSecurityFailure("OFFICE_AUTOMATION_SECURITY_UNAVAILABLE", "unsupported policy")
    messages = deque([refusal])
    connection = MagicMock()
    connection.poll.side_effect = lambda *args: bool(messages)
    connection.recv.side_effect = messages.popleft
    process = MagicMock()
    process.is_alive.return_value = False
    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge, "_start_com_conversion_process", lambda *args, **kwargs: (process, connection))
    assert (
        office_bridge._try_com_conversion_bounded(
            "in.doc", "out.docx", prog_id="Word.Application", save_format=16, app_type="word"
        )
        == refusal
    )
    connection.close.assert_called_once()


@pytest.mark.parametrize("safe_fallback_succeeds", [False, True])
def test_security_refusal_preserves_diagnostic_or_uses_gated_fallback(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, safe_fallback_succeeds: bool
) -> None:
    refusal = office_bridge._OfficeSecurityFailure("OFFICE_AUTOMATION_SECURITY_UNAVAILABLE", "policy unsupported")
    source = tmp_path / "input.doc"
    source.write_bytes(b"source")
    output = tmp_path / "output.docx"
    monkeypatch.setattr(office_bridge, "_try_com_conversion_bounded", lambda *args, **kwargs: refusal)
    monkeypatch.setattr(office_bridge, "_is_com_candidate_available", lambda candidate: True)
    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")

    def safe_fallback(input_path: str, output_path: str, **kwargs: object) -> str | None:
        if safe_fallback_succeeds:
            Path(output_path).write_bytes(b"converted")
            return output_path
        return None

    monkeypatch.setattr(office_bridge, "_try_libreoffice_conversion", safe_fallback)
    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(output),
        source_format="doc",
        backend_priority=["word", "libreoffice"],
        com_candidates={"word": office_bridge.BridgeCandidate("Word", "Word.Application", 16, "word")},
        libreoffice_format="docx",
    )
    assert result.attempted_backend_ids == ("word", "libreoffice")
    assert result.success is safe_fallback_succeeds
    if safe_fallback_succeeds:
        assert result.backend == "LibreOffice"
        assert output.read_bytes() == b"converted"
    else:
        assert result.error_code == "OFFICE_AUTOMATION_SECURITY_UNAVAILABLE"
        assert "word:" in result.message and "policy unsupported" in result.message
        assert not output.exists()
    assert not list(tmp_path.glob(".dw-office-*"))


@pytest.mark.parametrize("times_formula_fallback", [False, True])
def test_libreoffice_private_profile_disables_macros_and_preserves_font_policy(
    tmp_path: Path, times_formula_fallback: bool
) -> None:
    office_bridge._prepare_libreoffice_profile(tmp_path, times_formula_fallback=times_formula_fallback)
    root = ElementTree.parse(tmp_path / "user" / "registrymodifications.xcu").getroot()
    ns = "{http://openoffice.org/2001/registry}"
    properties = {
        (item.get(ns + "path"), prop.get(ns + "name")): prop.findtext("value")
        for item in root
        for prop in item.findall("prop")
    }
    assert properties[("/org.openoffice.Office.Common/Security/Scripting", "DisableMacrosExecution")] == "true"
    assert (("/org.openoffice.Office.Common/Font/Substitution", "Replacement") in properties) is times_formula_fallback
    assert ("Times New Roman" in ElementTree.tostring(root, encoding="unicode")) is times_formula_fallback


def test_word_probe_is_closed_and_removed_when_ownership_lookup_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    app = MagicMock()
    probe = app.Documents.Open.return_value
    probe.Windows.Item.side_effect = RuntimeError("window unavailable")
    opened_probe_paths: list[Path] = []

    def opened(path: str, **kwargs: object) -> MagicMock:
        opened_probe_paths.append(Path(path))
        assert Path(path).read_bytes() == b"{\\rtf1\\ansi DocWen ownership probe.}"
        return probe

    app.Documents.Open.side_effect = opened
    client = MagicMock()
    client.DispatchEx.return_value = app
    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (MagicMock(), client))
    monkeypatch.setattr(office_bridge, "_get_com_app_pid", lambda obj: None)
    result = office_bridge._try_com_conversion(
        str(tmp_path / "input.doc"),
        str(tmp_path / "out.docx"),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
    )
    assert result is None
    assert len(opened_probe_paths) == 1
    assert not opened_probe_paths[0].exists()
    assert "AutomationSecurity" not in app.__dict__
    probe.Close.assert_called_once_with(SaveChanges=False)
    app.Quit.assert_not_called()


def test_libreoffice_profile_failure_does_not_launch_process_and_is_diagnostic(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")
    monkeypatch.setattr(
        office_bridge, "_prepare_libreoffice_profile", MagicMock(side_effect=OSError("profile not writable"))
    )
    run = MagicMock(side_effect=AssertionError("unsafe launch"))
    monkeypatch.setattr(office_bridge, "_run_libreoffice_process", run)
    result = office_bridge._try_libreoffice_conversion(
        str(tmp_path / "in.doc"), str(tmp_path / "out.pdf"), convert_to="pdf"
    )
    assert isinstance(result, office_bridge._OfficeSecurityFailure)
    assert result.error_code == "OFFICE_MACRO_PROFILE_UNAVAILABLE"
    assert "profile not writable" in result.message
    run.assert_not_called()
    assert not list(tmp_path.glob(".dw-lo-*"))


def test_word_probe_close_failure_never_opens_real_input_and_still_cleans_up(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    app = MagicMock()
    probe = app.Documents.Open.return_value
    window = probe.Windows.Item.return_value
    probe.Close.side_effect = [RuntimeError("first close failed"), None]
    client = MagicMock()
    client.DispatchEx.return_value = app
    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (MagicMock(), client))
    monkeypatch.setattr(office_bridge, "_get_com_app_pid", lambda obj: 500 if obj is window else None)
    identity = WindowsProcessIdentity(500, 123)
    monkeypatch.setattr(WindowsProcessIdentity, "capture", lambda *args, **kwargs: identity)
    cleanup = MagicMock()
    monkeypatch.setattr(WindowsProcessIdentity, "terminate_if_running", lambda *args, **kwargs: cleanup())
    result = office_bridge._try_com_conversion(
        str(tmp_path / "input.doc"),
        str(tmp_path / "out.docx"),
        prog_id="Word.Application",
        save_format=16,
        app_type="word",
    )
    assert result is None
    app.Documents.Open.assert_called_once()
    assert Path(app.Documents.Open.call_args.args[0]).suffix == ".rtf"
    assert probe.Close.call_count == 2
    assert "AutomationSecurity" not in app.__dict__
    app.Quit.assert_called_once()
    cleanup.assert_called_once()
    assert not list(tmp_path.glob(".dw-word-probe-*"))
