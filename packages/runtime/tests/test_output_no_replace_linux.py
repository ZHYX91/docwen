"""Linux publication uses atomic no-replace rename without hard-link support."""

import ctypes
import errno
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from docwen_runtime.output import finalizer

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("error", [0, errno.EEXIST, errno.EOPNOTSUPP])
def test_linux_no_replace_preserves_system_error(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, error: int) -> None:
    library = MagicMock()
    library.renameat2.return_value = -1 if error else 0
    monkeypatch.setattr(finalizer.sys, "platform", "linux")
    monkeypatch.setattr(ctypes, "CDLL", lambda *args, **kwargs: library)
    monkeypatch.setattr(ctypes, "get_errno", lambda: error)
    monkeypatch.setattr(finalizer.os, "link", lambda *args: pytest.fail("publication must not require hard links"))
    source = str(tmp_path / "staged")
    destination = str(tmp_path / "published")
    if error:
        with pytest.raises(OSError) as raised:
            finalizer.OutputFinalizer._publish_no_clobber(source, destination)
        assert raised.value.errno == error
        if error == errno.EEXIST:
            assert isinstance(raised.value, FileExistsError)
    else:
        finalizer.OutputFinalizer._publish_no_clobber(source, destination)
    arguments = library.renameat2.call_args.args
    assert arguments[0] == arguments[2] == -100
    assert arguments[-1] == 1
