"""Managed clipboard SnapshotBundle ownership and integrity."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_gui.clipboard_inputs import ClipboardInputStore

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def test_bundle_owns_main_and_resources_as_one_lifecycle(tmp_path: Path) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    bundle = store.create_bundle(
        b'{"schema":"docwen.clipboard_document.v1","blocks":[]}',
        display_name_template="Clipboard Document {index}.dwclip",
        preview="[Table 1x1]",
        resources=(("image-a", "resources/a.png", "image/png", b"png-a"),),
    )
    main = Path(bundle.main.path)
    resource = Path(bundle.resources[0].path)
    marker = Path(bundle.marker_path)
    assert main.is_file() and resource.is_file() and marker.is_file()
    assert store.snapshot_available(main)
    assert store.snapshot_available(resource)

    store.sync_visible([main.as_posix()])
    store.retain_active("task", [resource.as_posix()])
    store.sync_visible([])
    assert main.exists() and resource.exists()

    store.retain_history("task", [main.as_posix()])
    store.release_active("task")
    assert main.exists() and resource.exists()
    store.release_history("task")
    assert not Path(bundle.root_path).exists()
    store.close()


def test_bundle_tamper_or_missing_marker_invalidates_whole_snapshot(tmp_path: Path) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    bundle = store.create_bundle(
        b'{"schema":"docwen.clipboard_document.v1","blocks":[]}',
        display_name_template="Clipboard Document {index}.dwclip",
        preview="",
        resources=(("resource-a", "resources/a.bin", "application/octet-stream", b"original"),),
    )
    resource = Path(bundle.resources[0].path)
    resource.write_bytes(b"tampered")
    assert not store.snapshot_available(bundle.main.path)
    resource.write_bytes(b"original")
    Path(bundle.marker_path).unlink()
    assert not store.snapshot_available(bundle.main.path)
    store.close()


def test_bundle_marker_is_not_published_when_creation_fails(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    original = store._write_fsynced
    calls = 0

    def fail_second(path: Path, payload: bytes) -> None:
        nonlocal calls
        calls += 1
        if calls == 2:
            raise OSError("resource write failed")
        original(path, payload)

    monkeypatch.setattr(store, "_write_fsynced", fail_second)
    with pytest.raises(OSError, match="resource write failed"):
        store.create_bundle(
            b"main",
            display_name_template="Clipboard Document {index}.dwclip",
            preview="",
            resources=(("resource-a", "resources/a.bin", "application/octet-stream", b"resource"),),
        )
    assert not list(store.session_root.glob("bundle-*/.bundle.marker"))
    store.close()
