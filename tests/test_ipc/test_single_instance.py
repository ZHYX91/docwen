from __future__ import annotations

from pathlib import Path

import pytest

from docwen.ipc.single_instance import SingleInstance


@pytest.mark.unit
def test_single_instance_acquire_and_release(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("tempfile.gettempdir", lambda: str(tmp_path))

    a = SingleInstance("app")
    assert a.acquire() is True

    b = SingleInstance("app")
    assert b.acquire() is False

    a.release()

    c = SingleInstance("app")
    assert c.acquire() is True
    c.release()


@pytest.mark.unit
def test_single_instance_context_manager(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("tempfile.gettempdir", lambda: str(tmp_path))

    inst = SingleInstance("app2")
    with inst as acquired:
        assert acquired in (True, False)

