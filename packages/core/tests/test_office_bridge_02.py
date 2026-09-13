"""Focused tests split from test_office_bridge.py."""

from __future__ import annotations

from ._office_bridge_support import (
    Path,
    _mock_owned_process,
    pytest,
)

pytestmark = pytest.mark.unit


def test_try_com_conversion_can_suppress_converter_created_word_revisions(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A route-owned private copy may disable recording SaveAs normalization."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.rtf"
    input_path.write_bytes(b"docx")

    class _PythonCom:
        @staticmethod
        def CoInitialize() -> None:
            return None

        @staticmethod
        def CoUninitialize() -> None:
            return None

    class _Document:
        TrackRevisions = True

        def SaveAs(self, output: str, *, FileFormat: int) -> None:
            assert FileFormat == 6
            assert self.TrackRevisions is False
            Path(output).write_bytes(b"{\\rtf1 output}")

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False

    class _Documents:
        document = _Document()

        def Open(self, input_file: str, **kwargs: object) -> _Document:
            assert input_file == str(input_path.resolve())
            assert kwargs["ReadOnly"] is True
            return self.document

    class _WordApp:
        def __init__(self) -> None:
            self.Documents = _Documents()

        def Quit(self) -> None:
            return None

    class _Win32Client:
        @staticmethod
        def DispatchEx(prog_id: str) -> _WordApp:
            assert prog_id == "Word.Application"
            return _WordApp()

    monkeypatch.setattr(office_bridge, "_import_win32", lambda: (_PythonCom, _Win32Client))
    _mock_owned_process(monkeypatch)

    result = office_bridge._try_com_conversion(
        str(input_path),
        str(output_path),
        prog_id="Word.Application",
        save_format=6,
        app_type="word",
        suppress_new_revisions=True,
    )

    assert result == str(output_path.resolve())
