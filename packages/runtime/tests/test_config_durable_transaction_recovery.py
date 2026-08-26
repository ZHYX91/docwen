"""Durable journal, process-death, and metadata contracts for DEBT-04."""

from __future__ import annotations

import hashlib
import json
import os
import stat
import subprocess
import sys
import textwrap
import time
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

_PROCESS_COORDINATION_TIMEOUT_SECONDS = 60.0


def _write_minimal_base_config_tree(base_dir: Path) -> None:
    from docwen_runtime.config.registry import CONFIG_FILES

    for spec in CONFIG_FILES:
        path = base_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


def _config_tree(tmp_path: Path) -> tuple[Path, Path, dict[str, bytes]]:
    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    _write_minimal_base_config_tree(base_dir)

    base_contents = {
        "gui.toml": '[window]\ndefault_mode = "single"\n',
        "output.toml": '[directory]\nmode = "source"\n',
    }
    user_contents = {
        "gui.toml": '# gui preimage\n[window]\ndefault_mode = "batch"\n',
        "output.toml": '# output preimage\n[directory]\nmode = "custom"\n',
    }
    for rel_path, content in base_contents.items():
        (base_dir / rel_path).write_text(content, encoding="utf-8")
    for rel_path, content in user_contents.items():
        (user_dir / rel_path).write_text(content, encoding="utf-8")
    return base_dir, user_dir, {rel_path: (user_dir / rel_path).read_bytes() for rel_path in user_contents}


def _run_child(script: str, base_dir: Path, user_dir: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, "-c", textwrap.dedent(script), str(base_dir), str(user_dir)],
        check=False,
        capture_output=True,
        text=True,
        timeout=20,
    )


def _wait_for(path: Path, *, timeout: float = _PROCESS_COORDINATION_TIMEOUT_SECONDS) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if path.exists():
            return
        time.sleep(0.02)
    raise AssertionError(f"timed out waiting for {path}")


def _canonical_json(value: object) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _write_valid_checksum_journal(user_dir: Path, payload: dict[str, object]) -> Path:
    envelope = {
        "checksum": hashlib.sha256(_canonical_json(payload)).hexdigest(),
        "payload": payload,
    }
    path = user_dir / ".docwen-config.transaction.json"
    path.write_bytes(_canonical_json(envelope) + b"\n")
    return path


def test_atomic_write_flushes_parent_directory_and_preserves_mode(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io

    target = tmp_path / "gui.toml"
    target.write_bytes(b"old = true\n")
    target.chmod(0o640)
    expected_mode = stat.S_IMODE(target.stat().st_mode)
    synced: list[Path] = []

    monkeypatch.setattr(toml_io, "_sync_directory", lambda path: synced.append(Path(path)))
    toml_io.atomic_write_bytes(target, b"new = true\n")

    assert target.read_bytes() == b"new = true\n"
    assert stat.S_IMODE(target.stat().st_mode) == expected_mode
    assert synced == [tmp_path]


def test_handled_multifile_failure_restores_regular_file_metadata(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.config import loader as loader_module
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, preimages = _config_tree(tmp_path)
    gui_path = user_dir / "gui.toml"
    gui_path.chmod(0o640)
    fixed_ns = 1_700_000_000_123_456_700
    os.utime(gui_path, ns=(fixed_ns, fixed_ns))
    before = gui_path.stat()
    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    original_write = loader_module.write_toml_file

    def fail_second(path: Path, data) -> None:
        if Path(path).name == "output.toml":
            raise OSError("simulated second-file failure")
        original_write(path, data)

    monkeypatch.setattr(loader_module, "write_toml_file", fail_second)

    assert (
        loader.set_values(
            {
                "gui.window.default_mode": "single",
                "output.directory.mode": "source",
            }
        )
        is False
    )
    after = gui_path.stat()
    assert gui_path.read_bytes() == preimages["gui.toml"]
    assert stat.S_IMODE(after.st_mode) == stat.S_IMODE(before.st_mode)
    assert after.st_mtime_ns == before.st_mtime_ns


def test_process_death_after_first_mutation_recovers_prepared_generation(
    tmp_path: Path,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, preimages = _config_tree(tmp_path)
    result = _run_child(
        """
        import os
        import sys
        from pathlib import Path
        from docwen_runtime.config import loader as loader_module
        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=Path(sys.argv[1]), user_dir=Path(sys.argv[2]))
        original_write = loader_module.write_toml_file
        writes = 0

        def crash_after_first(path, data):
            global writes
            original_write(path, data)
            writes += 1
            if writes == 1:
                os._exit(97)

        loader_module.write_toml_file = crash_after_first
        loader.set_values({
            "gui.window.default_mode": "single",
            "output.directory.mode": "source",
        })
        """,
        base_dir,
        user_dir,
    )
    assert result.returncode == 97, result.stderr

    recovered = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    assert recovered.config_state_trusted is True
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]


def test_process_death_before_commit_marker_recovers_old_generation(
    tmp_path: Path,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, preimages = _config_tree(tmp_path)
    result = _run_child(
        """
        import os
        import sys
        from pathlib import Path
        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=Path(sys.argv[1]), user_dir=Path(sys.argv[2]))
        loader.reload = lambda: os._exit(98)
        loader.set_values({
            "gui.window.default_mode": "single",
            "output.directory.mode": "source",
        })
        """,
        base_dir,
        user_dir,
    )
    assert result.returncode == 98, result.stderr

    recovered = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    assert recovered.config_state_trusted is True
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]


def test_process_death_after_commit_marker_keeps_new_generation(
    tmp_path: Path,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, _preimages = _config_tree(tmp_path)
    result = _run_child(
        """
        import os
        import sys
        from pathlib import Path
        from docwen_runtime.config import transaction
        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=Path(sys.argv[1]), user_dir=Path(sys.argv[2]))
        transaction.remove_transaction_journal = lambda *args, **kwargs: os._exit(99)
        loader.set_values({
            "gui.window.default_mode": "single",
            "output.directory.mode": "source",
        })
        """,
        base_dir,
        user_dir,
    )
    assert result.returncode == 99, result.stderr

    recovered = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    assert recovered.config_state_trusted is True
    assert recovered.config.gui.window.default_mode == "single"
    assert recovered.config.output.directory.mode == "source"


def test_corrupt_journal_fails_loader_closed_without_deleting_evidence(
    tmp_path: Path,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, _preimages = _config_tree(tmp_path)
    journal = user_dir / ".docwen-config.transaction.json"
    journal.write_bytes(b'{"version": 1, "payload": "truncated"')

    with pytest.raises(RuntimeError, match="configuration transaction journal"):
        ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    assert journal.read_bytes() == b'{"version": 1, "payload": "truncated"'


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("version", 2),
        ("state", "UNKNOWN"),
        ("path", "../escape.toml"),
    ],
)
def test_valid_checksum_but_invalid_journal_schema_fails_closed(
    tmp_path: Path,
    field: str,
    value: object,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, _preimages = _config_tree(tmp_path)
    payload: dict[str, object] = {
        "version": 1,
        "operation": "test",
        "state": "PREPARED",
        "preimages": [
            {
                "path": "gui.toml",
                "content": None,
                "symlink_target": None,
                "mode": None,
                "atime_ns": None,
                "mtime_ns": None,
            }
        ],
    }
    if field == "path":
        preimages = payload["preimages"]
        assert isinstance(preimages, list)
        assert isinstance(preimages[0], dict)
        preimages[0]["path"] = value
    else:
        payload[field] = value
    journal = _write_valid_checksum_journal(user_dir, payload)

    with pytest.raises(RuntimeError, match="configuration transaction journal"):
        ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    assert journal.exists()


def test_real_local_directory_flush_primitive_succeeds(tmp_path: Path) -> None:
    from docwen_runtime import toml_io

    toml_io._sync_directory(tmp_path)


def test_journal_parent_barrier_failure_precedes_every_user_mutation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, preimages = _config_tree(tmp_path)
    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    original_sync = toml_io._sync_directory
    failed = False

    def fail_prepare(directory: Path) -> None:
        nonlocal failed
        if Path(directory) == user_dir and not failed:
            failed = True
            raise OSError("simulated journal parent barrier failure")
        original_sync(directory)

    monkeypatch.setattr(toml_io, "_sync_directory", fail_prepare)

    assert (
        loader.set_values(
            {
                "gui.window.default_mode": "single",
                "output.directory.mode": "source",
            }
        )
        is False
    )
    assert failed is True
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]

    monkeypatch.setattr(toml_io, "_sync_directory", original_sync)
    loader.reload()
    assert not (user_dir / ".docwen-config.transaction.json").exists()


def test_user_file_parent_barrier_failure_compensates_prepared_transaction(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, preimages = _config_tree(tmp_path)
    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    original_sync = toml_io._sync_directory
    user_sync_count = 0

    def fail_first_user_file_barrier(directory: Path) -> None:
        nonlocal user_sync_count
        if Path(directory) == user_dir:
            user_sync_count += 1
            if user_sync_count == 2:
                raise OSError("simulated user-file parent barrier failure")
        original_sync(directory)

    monkeypatch.setattr(toml_io, "_sync_directory", fail_first_user_file_barrier)

    assert (
        loader.set_values(
            {
                "gui.window.default_mode": "single",
                "output.directory.mode": "source",
            }
        )
        is False
    )
    assert user_sync_count >= 2
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]
    assert not (user_dir / ".docwen-config.transaction.json").exists()


def test_delete_parent_barrier_failure_restores_deleted_override(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, preimages = _config_tree(tmp_path)
    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    original_sync = toml_io._sync_directory
    user_sync_count = 0

    def fail_delete_barrier(directory: Path) -> None:
        nonlocal user_sync_count
        if Path(directory) == user_dir:
            user_sync_count += 1
            if user_sync_count == 2:
                raise OSError("simulated delete parent barrier failure")
        original_sync(directory)

    monkeypatch.setattr(toml_io, "_sync_directory", fail_delete_barrier)

    assert loader.reset_file("gui.toml") is False
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert loader.config.gui.window.default_mode == "batch"


def test_commit_marker_failure_rolls_back_old_generation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.config import transaction
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, preimages = _config_tree(tmp_path)
    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

    def fail_commit(*_args, **_kwargs) -> None:
        raise OSError("simulated committed-marker barrier failure")

    monkeypatch.setattr(transaction, "mark_transaction_committed", fail_commit)

    assert (
        loader.set_values(
            {
                "gui.window.default_mode": "single",
                "output.directory.mode": "source",
            }
        )
        is False
    )
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]


def test_committed_cleanup_failure_is_safe_success_and_retried_on_reload(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.config import transaction
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, _preimages = _config_tree(tmp_path)
    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    original_remove = transaction.remove_transaction_journal

    def fail_cleanup(_user_dir: Path) -> None:
        raise OSError("simulated committed journal cleanup failure")

    monkeypatch.setattr(transaction, "remove_transaction_journal", fail_cleanup)
    assert loader.set_value("gui.window.default_mode", "single") is True
    journal = user_dir / ".docwen-config.transaction.json"
    assert journal.exists()
    assert loader.config.gui.window.default_mode == "single"

    monkeypatch.setattr(transaction, "remove_transaction_journal", original_remove)
    reloaded = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    assert reloaded.config_state_trusted is True
    assert reloaded.config.gui.window.default_mode == "single"
    assert not journal.exists()


def test_prepared_recovery_failure_leaves_journal_for_idempotent_retry(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io
    from docwen_runtime.config import transaction
    from docwen_runtime.config.loader import ConfigLoader

    base_dir, user_dir, preimages = _config_tree(tmp_path)
    gui_path = user_dir / "gui.toml"
    before = transaction.capture_user_file_preimage(gui_path)
    transaction.write_transaction_journal(
        user_dir,
        "test-recovery-retry",
        [before],
        state="PREPARED",
    )
    toml_io.atomic_write_bytes(gui_path, b'[window]\ndefault_mode = "single"\n')
    original_restore = transaction.restore_user_file_preimage

    def fail_restore(*_args, **_kwargs) -> None:
        raise OSError("simulated recovery barrier failure")

    monkeypatch.setattr(transaction, "restore_user_file_preimage", fail_restore)
    with pytest.raises(OSError, match="simulated recovery barrier failure"):
        ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    journal = user_dir / ".docwen-config.transaction.json"
    assert journal.exists()

    monkeypatch.setattr(transaction, "restore_user_file_preimage", original_restore)
    recovered = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    assert recovered.config_state_trusted is True
    assert gui_path.read_bytes() == preimages["gui.toml"]
    assert not journal.exists()


def test_two_real_processes_preserve_disjoint_same_file_updates(tmp_path: Path) -> None:
    from docwen_runtime.toml_io import read_toml_file

    base_dir, user_dir, _preimages = _config_tree(tmp_path)
    start = tmp_path / "start"
    script = textwrap.dedent(
        """
        import sys
        import time
        from pathlib import Path
        from docwen_runtime.config.loader import ConfigLoader

        base, user, ready, start, result, key, value = map(Path, sys.argv[1:8])
        loader = ConfigLoader(base_dir=base, user_dir=user)
        ready.write_text("ready", encoding="utf-8")
        while not start.exists():
            time.sleep(0.01)
        ok = loader.set_value(str(key), str(value))
        result.write_text(str(ok), encoding="utf-8")
        """
    )
    processes: list[subprocess.Popen[str]] = []
    results: list[Path] = []
    for index, (key, value) in enumerate(
        (
            ("gui.window.default_mode", "single"),
            ("gui.window.startup_mode", "batch"),
        )
    ):
        ready = tmp_path / f"ready-{index}"
        result = tmp_path / f"result-{index}"
        results.append(result)
        processes.append(
            subprocess.Popen(
                [
                    sys.executable,
                    "-c",
                    script,
                    str(base_dir),
                    str(user_dir),
                    str(ready),
                    str(start),
                    str(result),
                    key,
                    value,
                ],
                text=True,
            )
        )
        _wait_for(ready)
    start.write_text("go", encoding="utf-8")
    for process in processes:
        assert process.wait(timeout=_PROCESS_COORDINATION_TIMEOUT_SECONDS) == 0
    assert [path.read_text(encoding="utf-8") for path in results] == ["True", "True"]
    gui = read_toml_file(user_dir / "gui.toml")
    assert gui["window"]["default_mode"] == "single"
    assert gui["window"]["startup_mode"] == "batch"


def test_real_reload_blocks_until_multifile_writer_releases_process_lock(
    tmp_path: Path,
) -> None:
    base_dir, user_dir, _preimages = _config_tree(tmp_path)
    paused = tmp_path / "writer-paused"
    resume = tmp_path / "resume-writer"
    observed = tmp_path / "reader-observed"
    attempted = tmp_path / "reader-attempted"
    writer_script = textwrap.dedent(
        """
        import sys
        import time
        from pathlib import Path
        from docwen_runtime.config import loader as loader_module
        from docwen_runtime.config.loader import ConfigLoader

        base, user, paused, resume = map(Path, sys.argv[1:5])
        loader = ConfigLoader(base_dir=base, user_dir=user)
        original_write = loader_module.write_toml_file
        writes = 0
        def pause_after_first(path, data):
            global writes
            original_write(path, data)
            writes += 1
            if writes == 1:
                paused.write_text("paused", encoding="utf-8")
                while not resume.exists():
                    time.sleep(0.01)
        loader_module.write_toml_file = pause_after_first
        ok = loader.set_values({
            "gui.window.default_mode": "single",
            "output.directory.mode": "source",
        })
        raise SystemExit(0 if ok else 2)
        """
    )
    reader_script = textwrap.dedent(
        """
        import json
        import sys
        from pathlib import Path
        from docwen_runtime.config.loader import ConfigLoader

        base, user, attempted, observed = map(Path, sys.argv[1:5])
        attempted.write_text("attempted", encoding="utf-8")
        loader = ConfigLoader(base_dir=base, user_dir=user)
        observed.write_text(json.dumps({
            "gui": loader.config.gui.window.default_mode,
            "output": loader.config.output.directory.mode,
        }), encoding="utf-8")
        """
    )
    writer = subprocess.Popen(
        [sys.executable, "-c", writer_script, str(base_dir), str(user_dir), str(paused), str(resume)]
    )
    _wait_for(paused)
    reader = subprocess.Popen(
        [
            sys.executable,
            "-c",
            reader_script,
            str(base_dir),
            str(user_dir),
            str(attempted),
            str(observed),
        ]
    )
    _wait_for(attempted)
    time.sleep(0.35)
    assert not observed.exists()
    resume.write_text("resume", encoding="utf-8")
    assert writer.wait(timeout=20) == 0
    assert reader.wait(timeout=20) == 0
    assert json.loads(observed.read_text(encoding="utf-8")) == {
        "gui": "single",
        "output": "source",
    }
