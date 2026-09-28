"""Contract tests for managed clipboard Markdown snapshot ownership."""

from pathlib import Path

import pytest

from docwen_gui.clipboard_inputs import ClipboardInputStore, bounded_plaintext_preview

pytestmark = pytest.mark.contract


def test_clipboard_snapshot_preserves_exact_utf8_bytes_and_bounded_preview(tmp_path: Path) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    text = "---\ntitle: Test\n---\n```text\n  keep  \n```\n" + ("x" * 400)

    snapshot = store.create(text, display_name_template="Clipboard Markdown {index}.md")

    assert Path(snapshot.path).read_bytes() == text.encode("utf-8")
    assert snapshot.display_name == "Clipboard Markdown 1.md"
    assert snapshot.size_bytes == len(text.encode("utf-8"))
    assert len(snapshot.preview) <= 241
    assert not snapshot.preview.endswith(snapshot.path)
    store.sync_visible([snapshot.path])
    store.close()


def test_clipboard_snapshot_ownership_survives_active_and_failed_retry_then_cleans(tmp_path: Path) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    snapshot = store.create("# Snapshot\n", display_name_template="Clipboard Markdown {index}.md")
    path = Path(snapshot.path)

    store.sync_visible([snapshot.path])
    store.retain_active("task", [snapshot.path])
    store.sync_visible([])
    assert path.is_file()

    store.retain_history("task", [snapshot.path])
    store.release_active("task")
    assert path.is_file()

    store.release_history("task")
    assert not path.exists()
    store.close()


def test_clipboard_snapshot_empty_and_whitespace_are_rejected_without_files(tmp_path: Path) -> None:
    store = ClipboardInputStore(tmp_path / "managed")

    for text in ("", " \t\r\n "):
        with pytest.raises(ValueError, match="empty"):
            store.create(text, display_name_template="Clipboard Markdown {index}.md")

    assert list(store.session_root.iterdir()) == []
    store.close()


def test_plaintext_preview_replaces_controls_without_mutating_content_contract() -> None:
    assert bounded_plaintext_preview("A\x00B\nC") == "A�B\nC"

def test_create_fsync_failure_compensates_partial_snapshot(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    store = ClipboardInputStore(tmp_path / "managed")

    def fail_fsync(_descriptor: int) -> None:
        raise OSError("test fsync failure")

    monkeypatch.setattr("docwen_gui.clipboard_inputs.os.fsync", fail_fsync)
    with pytest.raises(OSError, match="fsync failure"):
        store.create("# not committed\n", display_name_template="Clipboard {index}.md")

    assert not list(store.session_root.glob("clipboard-*.md"))
    store.close()


def test_create_cleanup_failure_remains_retryable_until_delete_succeeds(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    original_unlink = Path.unlink

    def fail_fsync(_descriptor: int) -> None:
        raise OSError("test fsync failure")

    def fail_clipboard_unlink(path: Path, *args, **kwargs):
        if path.name.startswith("clipboard-"):
            raise PermissionError("test delete denied")
        return original_unlink(path, *args, **kwargs)

    monkeypatch.setattr("docwen_gui.clipboard_inputs.os.fsync", fail_fsync)
    monkeypatch.setattr(Path, "unlink", fail_clipboard_unlink)
    with pytest.raises(OSError, match="fsync failure"):
        store.create("# partial\n", display_name_template="Clipboard {index}.md")

    partials = list(store.session_root.glob("clipboard-*.md"))
    assert len(partials) == 1
    store.close()
    assert store.cleanup_pending is True
    assert partials[0].exists()

    monkeypatch.setattr(Path, "unlink", original_unlink)
    store.close()
    assert store.cleanup_pending is False
    assert not store.session_root.exists()


def test_close_delete_failure_does_not_forget_snapshot_tracking(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    snapshot = store.create("# retained on cleanup failure\n", display_name_template="Clipboard {index}.md")
    store.sync_visible([snapshot.path])
    original_unlink = Path.unlink

    def fail_snapshot_unlink(path: Path, *args, **kwargs):
        if path == Path(snapshot.path):
            raise PermissionError("test close delete denied")
        return original_unlink(path, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", fail_snapshot_unlink)
    store.close()
    assert store.cleanup_pending is True
    assert Path(snapshot.path).exists()

    monkeypatch.setattr(Path, "unlink", original_unlink)
    store.close()
    assert store.cleanup_pending is False
    assert not store.session_root.exists()


def test_session_lock_preserves_live_instance_and_recovers_provably_stale_session(tmp_path: Path) -> None:
    root = tmp_path / "managed"
    live = ClipboardInputStore(root)
    live_snapshot = live.create("# live\n", display_name_template="Clipboard {index}.md")
    live_marker = (live.session_root / ".owner.lock").read_bytes()

    peer = ClipboardInputStore(root)
    assert Path(live_snapshot.path).is_file()
    assert live.session_root.is_dir()

    stale = root / "session-crashed-fixture"
    stale.mkdir()
    (stale / ".owner.lock").write_bytes(live_marker)
    (stale / "clipboard-orphan.md").write_text("# orphan\n", encoding="utf-8")

    recovery = ClipboardInputStore(root)
    assert not stale.exists()
    assert Path(live_snapshot.path).is_file()
    assert live.session_root.is_dir()

    recovery.close()
    peer.close()
    live.close()

