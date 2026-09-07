"""Windows Shell selection and COM ownership, without opening desktop windows."""

import ctypes
from pathlib import Path
from types import SimpleNamespace

import pytest

from docwen_gui import windows_shell

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "initialized,parse_result,select_result", [(0, 0, 0), (1, 0, 0), (-2147417850, 0, 0), (0, -1, 0), (0, 0, -1)]
)
def test_shell_selects_unicode_item_and_releases_only_owned_resources(
    monkeypatch, tmp_path, initialized, parse_result, select_result
):
    selected = []
    released = []
    target = tmp_path / "中文 name, data.docx"

    def parse(name, unused, pointer, flags, attributes):
        assert Path(name) == target.absolute()
        if parse_result >= 0:
            ctypes.cast(pointer, ctypes.POINTER(ctypes.c_void_p))[0] = ctypes.c_void_p(123)
        return parse_result

    def select(item, count, children, flags):
        selected.append((item.value, count, children, flags))
        return select_result

    ole = SimpleNamespace(
        CoInitializeEx=lambda *_: initialized,
        CoUninitialize=lambda: released.append("com"),
        CoTaskMemFree=lambda item: released.append(item.value),
    )
    shell = SimpleNamespace(SHParseDisplayName=parse, SHOpenFolderAndSelectItems=select)
    monkeypatch.setattr(windows_shell, "_load_libraries", lambda: (ole, shell))
    if parse_result < 0 or select_result < 0:
        with pytest.raises(OSError):
            windows_shell.reveal_file(target)
    else:
        windows_shell.reveal_file(target)
    assert selected == ([(123, 0, None, 0)] if parse_result >= 0 else [])
    assert released == ([123] if parse_result >= 0 else []) + (["com"] if initialized >= 0 else [])
