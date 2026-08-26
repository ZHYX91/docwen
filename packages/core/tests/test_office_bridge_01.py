"""Focused tests split from test_office_bridge.py."""

from __future__ import annotations

from ._office_bridge_support import (
    Path,
    pytest,
)

pytestmark = pytest.mark.unit


def test_low_level_conversion_bypasses_are_not_public() -> None:
    from docwen_core import office_bridge

    assert not hasattr(office_bridge, "try_com_conversion")
    assert not hasattr(office_bridge, "try_com_conversion_bounded")
    assert not hasattr(office_bridge, "try_libreoffice_conversion")
    assert not hasattr(office_bridge, "convert_with_fallback")


def test_find_soffice_path_prefers_an_existing_path_candidate(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    path_candidate = tmp_path / "path" / "soffice.exe"
    path_candidate.parent.mkdir()
    path_candidate.write_bytes(b"exe")
    monkeypatch.setattr(office_bridge.shutil, "which", lambda _name: str(path_candidate))
    monkeypatch.setattr(
        office_bridge,
        "_find_windows_registered_soffice",
        lambda: pytest.fail("registered discovery must not run after PATH success"),
        raising=False,
    )

    assert office_bridge.find_soffice_path() == str(path_candidate)


def test_find_soffice_path_uses_a_valid_windows_app_path(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    registered = tmp_path / "registered" / "soffice.exe"
    registered.parent.mkdir()
    registered.write_bytes(b"exe")
    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge.shutil, "which", lambda _name: None)
    monkeypatch.setattr(
        office_bridge,
        "_read_windows_soffice_app_path",
        lambda: str(registered),
        raising=False,
    )

    assert office_bridge.find_soffice_path() == str(registered)


def test_find_soffice_path_uses_a_bounded_windows_standard_location(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    program_files = tmp_path / "Program Files"
    standard = program_files / "LibreOffice" / "program" / "soffice.exe"
    standard.parent.mkdir(parents=True)
    standard.write_bytes(b"exe")
    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge.shutil, "which", lambda _name: None)
    monkeypatch.setattr(
        office_bridge,
        "_read_windows_soffice_app_path",
        lambda: str(tmp_path / "missing" / "soffice.exe"),
        raising=False,
    )
    monkeypatch.setenv("ProgramFiles", str(program_files))
    monkeypatch.delenv("ProgramW6432", raising=False)
    monkeypatch.delenv("ProgramFiles(x86)", raising=False)

    assert office_bridge.find_soffice_path() == str(standard)


def test_find_soffice_path_rejects_missing_and_directory_candidates(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    directory = tmp_path / "soffice.exe"
    directory.mkdir()
    monkeypatch.setattr(office_bridge.sys, "platform", "win32")
    monkeypatch.setattr(office_bridge.shutil, "which", lambda _name: str(directory))
    monkeypatch.setattr(
        office_bridge,
        "_read_windows_soffice_app_path",
        lambda: str(tmp_path / "missing.exe"),
        raising=False,
    )
    monkeypatch.setenv("ProgramFiles", str(tmp_path / "missing-program-files"))
    monkeypatch.delenv("ProgramW6432", raising=False)
    monkeypatch.delenv("ProgramFiles(x86)", raising=False)

    assert office_bridge.find_soffice_path() is None


def test_find_soffice_path_does_not_run_windows_discovery_on_other_platforms(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    monkeypatch.setattr(office_bridge.sys, "platform", "linux")
    monkeypatch.setattr(office_bridge.shutil, "which", lambda _name: None)
    monkeypatch.setattr(
        office_bridge,
        "_find_windows_registered_soffice",
        lambda: pytest.fail("Windows discovery must remain platform-gated"),
        raising=False,
    )

    assert office_bridge.find_soffice_path() is None


def test_try_com_conversion_continues_when_visible_cannot_be_hidden(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """PowerPoint may reject ``Visible = False`` but can still SaveAs."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.ppt"
    output_path = tmp_path / "output.pptx"
    input_path.write_bytes(b"legacy ppt")

    class _PythonCom:
        initialized = False
        uninitialized = False

        @classmethod
        def CoInitialize(cls) -> None:
            cls.initialized = True

        @classmethod
        def CoUninitialize(cls) -> None:
            cls.uninitialized = True

    class _Presentation:
        closed = False

        def SaveAs(self, output: str, save_format: int) -> None:
            assert save_format == 24
            Path(output).write_bytes(b"converted pptx")

        def Close(self) -> None:
            self.closed = True

    class _Presentations:
        def __init__(self) -> None:
            self.opened: str | None = None

        def Open(self, input_file: str, *args: object) -> _Presentation:
            self.opened = input_file
            return _Presentation()

    class _PowerPointApp:
        def __init__(self) -> None:
            self.Presentations = _Presentations()
            self.quit_called = False

        @property
        def Visible(self) -> bool:
            return True

        @Visible.setter
        def Visible(self, value: bool) -> None:
            raise RuntimeError("PowerPoint cannot be hidden")

        def Quit(self) -> None:
            self.quit_called = True

    class _Win32Client:
        app = _PowerPointApp()

        @classmethod
        def Dispatch(cls, prog_id: str) -> _PowerPointApp:
            assert prog_id == "PowerPoint.Application"
            return cls.app

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="PowerPoint.Application",
        save_format=24,
        app_type="powerpoint",
    )

    assert result == str(output_path.resolve())
    assert output_path.read_bytes() == b"converted pptx"
    assert _PythonCom.initialized is True
    assert _PythonCom.uninitialized is True
    assert _Win32Client.app.Presentations.opened == str(input_path.resolve())
    assert _Win32Client.app.quit_called is True


def test_powerpoint_open_fallback_remains_explicitly_read_only(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    input_path = tmp_path / "input.ppt"
    output_path = tmp_path / "output.pptx"
    input_path.write_bytes(b"legacy ppt")
    open_attempts: list[tuple[object, ...]] = []

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None: ...

        @staticmethod
        def CoUninitialize() -> None: ...

    class _Presentation:
        def SaveAs(self, output: str, _save_format: int) -> None:
            Path(output).write_bytes(b"pptx")

        def Close(self) -> None: ...

    class _Presentations:
        def Open(self, _path: str, *args: object) -> _Presentation:
            open_attempts.append(args)
            if len(args) == 3:
                raise TypeError("legacy COM signature")
            assert args == (True,)
            return _Presentation()

    class _App:
        def __init__(self) -> None:
            self.Presentations = _Presentations()

        def Quit(self) -> None: ...

    class _Win32Client:
        @staticmethod
        def Dispatch(_prog_id: str) -> _App:
            return _App()

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="PowerPoint.Application",
        save_format=24,
        app_type="powerpoint",
    )

    assert result == str(output_path.resolve())
    assert open_attempts == [(True, False, False), (True,)]


def test_try_com_conversion_uses_fixed_format_for_excel_pdf(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Excel-family PDF export should use ExportAsFixedFormat."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.pdf"
    input_path.write_bytes(b"xlsx")

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None:
            return None

        @staticmethod
        def CoUninitialize() -> None:
            return None

    class _Workbook:
        save_as_called = False
        export_called = False

        def SaveAs(self, output: str, *, FileFormat: int) -> None:
            self.save_as_called = True
            raise AssertionError("PDF export should not prefer SaveAs")

        def ExportAsFixedFormat(self, export_type: int, output: str) -> None:
            assert export_type == 0
            self.export_called = True
            Path(output).write_bytes(b"%PDF-1.4")

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False

    class _Workbooks:
        workbook = _Workbook()

        def Open(self, input_file: str, **_kwargs: object) -> _Workbook:
            assert input_file == str(input_path.resolve())
            return self.workbook

    class _ExcelApp:
        def __init__(self) -> None:
            self.Workbooks = _Workbooks()

        def Quit(self) -> None:
            return None

    class _Win32Client:
        @staticmethod
        def Dispatch(prog_id: str) -> _ExcelApp:
            assert prog_id == "KET.Application"
            return _ExcelApp()

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="KET.Application",
        save_format=57,
        app_type="excel",
    )

    assert result == str(output_path.resolve())
    assert output_path.read_bytes() == b"%PDF-1.4"
    assert _Workbooks.workbook.export_called is True
    assert _Workbooks.workbook.save_as_called is False


def test_try_com_conversion_opens_spreadsheets_read_only_without_updating_links(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Spreadsheet bridges must not refresh external targets or mutate input."""
    from docwen_core import office_bridge

    input_path = tmp_path / "linked.xlsx"
    output_path = tmp_path / "linked.xls"
    input_path.write_bytes(b"xlsx")
    open_options: dict[str, object] = {}

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None:
            return None

        @staticmethod
        def CoUninitialize() -> None:
            return None

    class _Workbook:
        def SaveAs(self, output: str, *, FileFormat: int) -> None:
            assert FileFormat == 56
            Path(output).write_bytes(b"xls")

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False

    class _Workbooks:
        def Open(self, input_file: str, **kwargs: object) -> _Workbook:
            assert input_file == str(input_path.resolve())
            open_options.update(kwargs)
            return _Workbook()

    class _ExcelApp:
        def __init__(self) -> None:
            self.Workbooks = _Workbooks()

        def Quit(self) -> None:
            return None

    class _Win32Client:
        @staticmethod
        def Dispatch(prog_id: str) -> _ExcelApp:
            assert prog_id == "Excel.Application"
            return _ExcelApp()

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="Excel.Application",
        save_format=56,
        app_type="excel",
    )

    assert result == str(output_path.resolve())
    assert open_options == {
        "UpdateLinks": 0,
        "ReadOnly": True,
        "IgnoreReadOnlyRecommended": True,
        "AddToMru": False,
    }


def test_try_com_conversion_prefers_dispatch_ex_for_isolated_com_instance(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """COM conversions should prefer a new application instance when available."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.odt"
    input_path.write_bytes(b"docx")

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None:
            return None

        @staticmethod
        def CoUninitialize() -> None:
            return None

    class _Document:
        def SaveAs(self, output: str, *, FileFormat: int) -> None:
            assert FileFormat == 23
            Path(output).write_bytes(b"odt")

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False

    class _Documents:
        def Open(self, input_file: str, **kwargs: object) -> _Document:
            assert input_file == str(input_path.resolve())
            assert kwargs["ReadOnly"] is True
            return _Document()

    class _WordApp:
        def __init__(self) -> None:
            self.Documents = _Documents()
            self.quit_called = False

        def Quit(self) -> None:
            self.quit_called = True

    class _Win32Client:
        dispatch_called = False
        dispatch_ex_called = False
        app = _WordApp()

        @classmethod
        def DispatchEx(cls, prog_id: str) -> _WordApp:
            assert prog_id == "Word.Application"
            cls.dispatch_ex_called = True
            return cls.app

        @classmethod
        def Dispatch(cls, prog_id: str) -> _WordApp:
            cls.dispatch_called = True
            raise AssertionError(f"Dispatch fallback should not be used for {prog_id}")

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="Word.Application",
        save_format=23,
        app_type="word",
    )

    assert result == str(output_path.resolve())
    assert output_path.read_bytes() == b"odt"
    assert _Win32Client.dispatch_ex_called is True
    assert _Win32Client.dispatch_called is False
    assert _Win32Client.app.quit_called is True
