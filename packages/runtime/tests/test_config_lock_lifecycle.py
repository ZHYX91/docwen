"""An unsuccessful lock entry must release every acquired OS resource."""

import os

import pytest

from docwen_runtime.config.transaction import _ProcessFileLock

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("stage", ["initialize", "lock"])
def test_failed_entry_closes_descriptor_and_can_retry(tmp_path, monkeypatch, stage):
    opened = []
    real_open = os.open

    def capture(*args, **kwargs):
        descriptor = real_open(*args, **kwargs)
        opened.append(descriptor)
        return descriptor

    def fail(*args):
        raise PermissionError("Injected ordinary I/O failure")

    lock = _ProcessFileLock(tmp_path / "config.lock")
    with monkeypatch.context() as patch:
        patch.setattr(os, "open", capture)
        if stage == "initialize":
            patch.setattr(os, "write", fail)
        elif os.name == "nt":
            patch.setattr("msvcrt.locking", fail)
        else:
            patch.setattr("fcntl.flock", fail)
        with pytest.raises(PermissionError), lock:
            pytest.fail("Unacquired lock entered")
    assert lock._descriptor is None
    with pytest.raises(OSError):
        os.fstat(opened[-1])
    with lock:
        assert lock._descriptor is not None
    assert lock._descriptor is None
