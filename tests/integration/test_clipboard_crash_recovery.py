"""A failed crash recovery must preserve authority for the next startup."""

import os
import sys
from pathlib import Path

import pytest

from docwen_gui.clipboard_inputs import ClipboardInputStore
from tests.support.subprocess_runner import run_subprocess

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def test_partial_crash_recovery_keeps_marker_until_body_deleted(tmp_path, monkeypatch):
    root = tmp_path / "managed"
    child = run_subprocess(
        [
            sys.executable,
            "-c",
            "import os,sys; from docwen_gui.clipboard_inputs import ClipboardInputStore; "
            "s=ClipboardInputStore(sys.argv[1]); "
            "s.create('# crashed body', display_name_template='Clipboard {index}.md'); os._exit(23)",
            str(root),
        ],
        timeout=20,
    )
    assert child.returncode == 23, child.stderr
    stale = next(root.glob("session-*"))
    source = next(stale.glob("clipboard-*.md"))
    marker = stale / ".owner.lock"
    marker_bytes = marker.read_bytes()
    legacy = root / "session-unmarked"
    legacy.mkdir()
    (legacy / "body.md").write_text("not owned", encoding="utf-8")
    original_iter = Path.iterdir
    original_unlink = os.unlink
    attempted = []

    def marker_first(path):
        entries = list(original_iter(path))
        return iter(sorted(entries, key=lambda item: item.name != ".owner.lock")) if path == stale else iter(entries)

    def deny_body(path, *args, **kwargs):
        if os.fspath(path) in {str(source), source.name}:
            attempted.append(source.name)
            raise PermissionError("controlled body removal failure")
        return original_unlink(path, *args, **kwargs)

    monkeypatch.setattr(Path, "iterdir", marker_first)
    first = None
    second = None
    try:
        with monkeypatch.context() as fault:
            fault.setattr(os, "unlink", deny_body)
            first = ClipboardInputStore(root)
        assert attempted == [source.name]
        assert source.read_text(encoding="utf-8") == "# crashed body"
        assert marker.read_bytes() == marker_bytes
        second = ClipboardInputStore(root)
        assert not stale.exists()
        assert first.session_root.is_dir()
        assert (legacy / "body.md").read_text(encoding="utf-8") == "not owned"
    finally:
        if second is not None:
            second.close()
        if first is not None:
            first.close()
