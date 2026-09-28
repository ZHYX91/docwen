"""Real session locks serialize initialization and retired-directory deletion."""

from __future__ import annotations

import threading
from pathlib import Path

import pytest

import docwen_gui.clipboard_inputs as clipboard

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("barrier", ["owner-lock", "retirement"])
def test_session_namespace_excludes_concurrent_publication_and_recovery(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, barrier: str
) -> None:
    root = tmp_path / "managed"
    root.mkdir()
    stale = root / "session-stale"
    if barrier == "retirement":
        stale.mkdir()
        (stale / ".owner.lock").write_bytes(clipboard._OWNER_MARKER)
        (stale / "orphan.md").write_text("old body", encoding="utf-8")
    entered = threading.Event()
    release = threading.Event()
    peer_attempted = threading.Event()
    peer_acquired = threading.Event()
    lock = clipboard._lock_owner_stream
    remove = clipboard.shutil.rmtree

    def observed_lock(stream, *, blocking: bool) -> None:
        name = Path(stream.name).name
        thread_name = threading.current_thread().name
        if thread_name == "clipboard-peer" and name == ".sessions.lock":
            peer_attempted.set()
        if thread_name == "clipboard-first" and name == ".owner.lock" and blocking and barrier == "owner-lock":
            entered.set()
            assert release.wait(5), "initialization barrier not released"
        lock(stream, blocking=blocking)
        if thread_name == "clipboard-peer" and name == ".sessions.lock":
            peer_acquired.set()

    def observed_remove(path, *args, **kwargs) -> None:
        if Path(path) == stale and threading.current_thread().name == "clipboard-first":
            entered.set()
            assert release.wait(5), "retirement barrier not released"
        remove(path, *args, **kwargs)

    monkeypatch.setattr(clipboard, "_lock_owner_stream", observed_lock)
    monkeypatch.setattr(clipboard.shutil, "rmtree", observed_remove)
    stores: list[clipboard.ClipboardInputStore] = []
    snapshots: list[str] = []
    errors: list[BaseException] = []

    def create_store() -> None:
        try:
            store = clipboard.ClipboardInputStore(root)
            stores.append(store)
            snapshot = store.create("# retained body\n", display_name_template="Clipboard {index}.md")
            store.sync_visible([snapshot.path])
            store.retain_active("worker", [snapshot.path])
            snapshots.append(snapshot.path)
        except BaseException as error:
            errors.append(error)

    first = threading.Thread(target=create_store, name="clipboard-first")
    peer = threading.Thread(target=create_store, name="clipboard-peer")
    first.start()
    try:
        assert entered.wait(5)
        # This is a real OS nonblocking exclusion probe, not a timing delay.
        with (root / ".sessions.lock").open("r+b") as stream, pytest.raises(OSError):
            lock(stream, blocking=False)
        peer.start()
        assert peer_attempted.wait(5)
        assert not peer_acquired.is_set()
    finally:
        release.set()
        first.join(10)
        if peer.ident is not None:
            peer.join(10)
    try:
        assert not first.is_alive() and not peer.is_alive()
        assert errors == []
        assert len(snapshots) == 2
        assert all(Path(path).read_text(encoding="utf-8") == "# retained body\n" for path in snapshots)
        assert all(store.session_root.is_dir() for store in stores)
        assert not stale.exists()
    finally:
        for store in stores:
            store.release_active("worker")
            store.close()
