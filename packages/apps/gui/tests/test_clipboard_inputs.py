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
