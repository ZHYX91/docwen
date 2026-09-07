"""Reveal a filesystem item with the Windows Shell's explicit selection API."""

from __future__ import annotations

import ctypes
from pathlib import Path
from typing import Any

_RPC_E_CHANGED_MODE = -2147417850


def _load_libraries() -> tuple[Any, Any]:
    ole = ctypes.WinDLL("ole32", use_last_error=True)
    shell = ctypes.WinDLL("shell32", use_last_error=True)
    ole.CoInitializeEx.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
    ole.CoInitializeEx.restype = ctypes.c_long
    ole.CoUninitialize.argtypes = []
    ole.CoUninitialize.restype = None
    ole.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    ole.CoTaskMemFree.restype = None
    shell.SHParseDisplayName.argtypes = [
        ctypes.c_wchar_p,
        ctypes.c_void_p,
        ctypes.POINTER(ctypes.c_void_p),
        ctypes.c_ulong,
        ctypes.POINTER(ctypes.c_ulong),
    ]
    shell.SHParseDisplayName.restype = ctypes.c_long
    shell.SHOpenFolderAndSelectItems.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_ulong]
    shell.SHOpenFolderAndSelectItems.restype = ctypes.c_long
    return ole, shell


def reveal_file(path: Path) -> None:
    """Open the parent and select this item, without Explorer exit-code guessing.

    https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shopenfolderandselectitems
    """
    ole, shell = _load_libraries()
    initialized = ole.CoInitializeEx(None, 2)
    if initialized < 0 and initialized != _RPC_E_CHANGED_MODE:
        raise OSError(initialized, "Windows Shell initialization failed")
    item = ctypes.c_void_p()
    try:
        parsed = shell.SHParseDisplayName(str(path.absolute()), None, ctypes.byref(item), 0, None)
        if parsed < 0 or not item.value:
            raise OSError(parsed, "Windows Shell could not resolve the file")
        result = shell.SHOpenFolderAndSelectItems(item, 0, None, 0)
        if result < 0:
            raise OSError(result, "Windows Shell could not select the file")
    finally:
        if item.value:
            ole.CoTaskMemFree(item)
        if initialized >= 0:
            ole.CoUninitialize()
