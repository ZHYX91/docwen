"""Windows handle-pinned staging cleanup. No path-based deletion after validation."""

from __future__ import annotations

import ctypes
from contextlib import ExitStack, contextmanager
from ctypes import wintypes
from pathlib import Path


class _FileInfo(ctypes.Structure):
    _fields_ = (
        ("attributes", wintypes.DWORD),
        ("creation", wintypes.FILETIME),
        ("access", wintypes.FILETIME),
        ("write", wintypes.FILETIME),
        ("volume", wintypes.DWORD),
        ("size_high", wintypes.DWORD),
        ("size_low", wintypes.DWORD),
        ("links", wintypes.DWORD),
        ("index_high", wintypes.DWORD),
        ("index_low", wintypes.DWORD),
    )


class _Disposition(ctypes.Structure):
    _fields_ = (("delete", wintypes.BOOL),)


class _Handles:
    def __init__(self) -> None:
        self.api = vars(ctypes)["WinDLL"]("kernel32", use_last_error=True)
        self.api.CreateFileW.argtypes = (
            wintypes.LPCWSTR,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.LPVOID,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.HANDLE,
        )
        self.api.CreateFileW.restype = wintypes.HANDLE
        self.api.GetFileInformationByHandle.argtypes = (wintypes.HANDLE, ctypes.POINTER(_FileInfo))
        self.api.GetFileInformationByHandle.restype = wintypes.BOOL
        self.api.SetFileInformationByHandle.argtypes = (
            wintypes.HANDLE,
            ctypes.c_int,
            wintypes.LPVOID,
            wintypes.DWORD,
        )
        self.api.SetFileInformationByHandle.restype = wintypes.BOOL
        self.api.CloseHandle.argtypes = (wintypes.HANDLE,)
        self.api.CloseHandle.restype = wintypes.BOOL

    @staticmethod
    def error() -> OSError:
        return vars(ctypes)["WinError"](vars(ctypes)["get_last_error"]())

    @contextmanager
    def open(self, path: Path, *, delete: bool = False):
        # Deny write/delete sharing: ancestors cannot be renamed or made reparse points.
        handle = self.api.CreateFileW(
            str(path), 0x80 | (0x10000 if delete else 0), 1, None, 3, 0x02000000 | 0x00200000, None
        )
        if handle == ctypes.c_void_p(-1).value:
            raise self.error()
        try:
            metadata = _FileInfo()
            if not self.api.GetFileInformationByHandle(handle, ctypes.byref(metadata)):
                raise self.error()
            yield handle, metadata.attributes
        finally:
            self.api.CloseHandle(handle)

    def remove(self, handle) -> None:
        disposition = _Disposition(True)
        if not self.api.SetFileInformationByHandle(handle, 4, ctypes.byref(disposition), ctypes.sizeof(disposition)):
            raise self.error()


def discard_windows(root: Path, relative: Path) -> None:
    api = _Handles()
    with ExitStack() as stack:
        current = Path(root.anchor)
        # Pin every ancestor before opening its child. Handle attributes are authoritative.
        for part in root.parts[1:]:
            current /= part
            _, attributes = stack.enter_context(api.open(current))
            if attributes & 0x400 or not attributes & 0x10:
                return
        _remove_descendant(api, current, relative.parts)


def _remove_descendant(api: _Handles, parent: Path, parts: tuple[str, ...]) -> bool:
    with api.open(parent / parts[0], delete=True) as (handle, attributes):
        if len(parts) > 1:
            if attributes & 0x400 or not attributes & 0x10:
                return False
            if not _remove_descendant(api, parent / parts[0], parts[1:]):
                return False
        elif attributes & 0x10 and not attributes & 0x400:
            return False
        # Children are already closed; deletion uses the same handle we validated.
        api.remove(handle)
        return True
