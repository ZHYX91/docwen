"""Shared config journal and process coordination fixtures."""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import textwrap
import time
from pathlib import Path

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
