"""ipc 单元测试。"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docwen.ipc.file_ipc import CommandHandler, FileIPC

pytestmark = pytest.mark.unit


class _DummyEvent:
    def __init__(self, src_path: str, is_directory: bool = False) -> None:
        self.src_path = src_path
        self.is_directory = is_directory


@pytest.mark.unit
def test_send_command_writes_json(tmp_path: Path) -> None:
    ipc_dir = str(tmp_path)
    commands_dir = tmp_path / "commands"
    commands_dir.mkdir()

    assert FileIPC.send_command(ipc_dir, {"action": "ping"}) is True
    files = list(commands_dir.glob("cmd_*.json"))
    assert len(files) == 1
    data = json.loads(files[0].read_text(encoding="utf-8"))
    assert data["action"] == "ping"


@pytest.mark.unit
def test_command_handler_processes_and_deletes_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("time.sleep", lambda *_args, **_kwargs: None)

    received = []

    def callback(cmd):
        received.append(cmd)

    cmd_file = tmp_path / "cmd.json"
    cmd_file.write_text(json.dumps({"action": "test", "data": 1}, ensure_ascii=False), encoding="utf-8")

    handler = CommandHandler(callback)
    handler.on_created(_DummyEvent(str(cmd_file)))

    assert received == [{"action": "test", "data": 1}]
    assert cmd_file.exists() is False


@pytest.mark.unit
def test_check_instance_running_uses_status_file(tmp_path: Path) -> None:
    ipc_dir = str(tmp_path)
    assert FileIPC.check_instance_running(ipc_dir) is False
    (tmp_path / "status.json").write_text("{}", encoding="utf-8")
    assert FileIPC.check_instance_running(ipc_dir) is True
