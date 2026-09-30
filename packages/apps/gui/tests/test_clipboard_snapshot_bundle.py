"""Managed clipboard SnapshotBundle ownership and integrity."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from docwen_core.models.clipboard_document import CLIPBOARD_DOCUMENT_SCHEMA
from docwen_gui.clipboard_inputs import ClipboardInputStore

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def _payload(
    resource_id: str = "resource-a",
    logical_path: str = "resources/a.bin",
    media_type: str = "application/octet-stream",
    resource_bytes: bytes = b"original",
    *,
    include_resource: bool = True,
) -> bytes:
    resources = []
    if include_resource:
        resources.append(
            {
                "resourceId": resource_id,
                "logicalPath": logical_path,
                "mediaType": media_type,
                "sizeBytes": len(resource_bytes),
                "sha256": hashlib.sha256(resource_bytes).hexdigest(),
            }
        )
    return json.dumps(
        {"schema": CLIPBOARD_DOCUMENT_SCHEMA, "blocks": [], "resources": resources},
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")


def test_bundle_owns_main_and_resources_as_one_lifecycle(tmp_path: Path) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    resource_bytes = b"png-a"
    bundle = store.create_bundle(
        _payload("image-a", "resources/a.png", "image/png", resource_bytes),
        display_name_template="Clipboard Document {index}.dwclip",
        preview="[Table 1x1]",
        resources=(("image-a", "resources/a.png", "image/png", resource_bytes),),
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


@pytest.mark.parametrize(
    ("payload", "resources"),
    [
        (_payload(), ()),
        (
            _payload(include_resource=False),
            (("resource-a", "resources/a.bin", "application/octet-stream", b"original"),),
        ),
        (
            _payload(),
            (("resource-a", "resources/other.bin", "application/octet-stream", b"original"),),
        ),
        (
            _payload(),
            (("resource-a", "resources/a.bin", "image/png", b"original"),),
        ),
        (
            _payload(resource_bytes=b"expected"),
            (("resource-a", "resources/a.bin", "application/octet-stream", b"tampered"),),
        ),
    ],
)
def test_bundle_rejects_resource_declaration_member_mismatches(
    tmp_path: Path,
    payload: bytes,
    resources: tuple[tuple[str, str, str, bytes], ...],
) -> None:
    store = ClipboardInputStore(tmp_path / "managed")

    with pytest.raises(ValueError, match="resource"):
        store.create_bundle(
            payload,
            display_name_template="Clipboard Document {index}.dwclip",
            preview="",
            resources=resources,
        )

    assert not list(store.session_root.glob("bundle-*"))
    store.close()


def test_bundle_tamper_or_missing_marker_invalidates_whole_snapshot(tmp_path: Path) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    bundle = store.create_bundle(
        _payload(),
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
            _payload(resource_bytes=b"resource"),
            display_name_template="Clipboard Document {index}.dwclip",
            preview="",
            resources=(("resource-a", "resources/a.bin", "application/octet-stream", b"resource"),),
        )
    assert not list(store.session_root.glob("bundle-*/.bundle.marker"))
    assert not list(store.session_root.glob("bundle-*"))
    store.close()



def test_bundle_delete_failure_keeps_tracking_until_retry_succeeds(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    bundle = store.create_bundle(
        _payload(),
        display_name_template="Clipboard Document {index}.dwclip",
        preview="",
        resources=(("resource-a", "resources/a.bin", "application/octet-stream", b"original"),),
    )
    bundle_root = Path(bundle.root_path)
    store.sync_visible([bundle.main.path])
    original_rmtree = __import__("shutil").rmtree

    def deny_bundle_delete(path, *args, **kwargs):
        if Path(path) == bundle_root:
            raise PermissionError("test bundle delete denied")
        return original_rmtree(path, *args, **kwargs)

    monkeypatch.setattr("docwen_gui.clipboard_inputs.shutil.rmtree", deny_bundle_delete)
    store.sync_visible([])
    assert bundle_root.exists()
    assert store.bundle(bundle.main.path) == bundle
    assert store.snapshot_available(bundle.main.path)

    monkeypatch.setattr("docwen_gui.clipboard_inputs.shutil.rmtree", original_rmtree)
    store.discard_if_unowned(bundle.main.path)
    assert not bundle_root.exists()
    store.close()
