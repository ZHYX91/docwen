"""A failed crash recovery must preserve authority for the next startup."""

import json
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


def test_partial_resource_bundle_crash_recovery_preserves_authority(tmp_path, monkeypatch):
    root = tmp_path / "managed"
    program = """
import hashlib, json, os, sys
from docwen_core.models.clipboard_document import CLIPBOARD_DOCUMENT_SCHEMA
from docwen_gui.clipboard_inputs import ClipboardInputStore
store = ClipboardInputStore(sys.argv[1])
resources = tuple((f'asset-{i}', f'resources/asset-{i}.bin', 'application/octet-stream',
                   f'resource-bytes-{i}'.encode()) for i in range(2))
payload = json.dumps({
    'schema': CLIPBOARD_DOCUMENT_SCHEMA,
    'blocks': [{'type': 'paragraph', 'inlines': [
        {'type': 'image', 'resourceId': item[0], 'alt': item[0], 'missingReason': ''}
        for item in resources]}],
    'resources': [{'resourceId': item[0], 'logicalPath': item[1], 'mediaType': item[2],
                   'sizeBytes': len(item[3]), 'sha256': hashlib.sha256(item[3]).hexdigest()}
                  for item in resources]
}).encode()
bundle = store.create_bundle(payload, display_name_template='Clipboard {index}.dwclip',
                             preview='crashed resources', resources=resources)
store.sync_visible([bundle.main.path])
print(json.dumps({'session': str(store.session_root), 'main': bundle.main.path,
                  'resources': [item.path for item in bundle.resources]}), flush=True)
os._exit(23)
"""
    child = run_subprocess([sys.executable, "-c", program, str(root)], timeout=20)
    assert child.returncode == 23, child.stderr
    saved = json.loads(child.stdout)
    stale = Path(saved["session"])
    assert stale.parent == root
    first_resource, blocked_resource = map(Path, saved["resources"])
    assert first_resource.is_relative_to(stale) and blocked_resource.is_relative_to(stale)
    marker = stale / ".owner.lock"
    marker_bytes = marker.read_bytes()
    legacy = root / "session-unmarked"
    legacy.mkdir()
    (legacy / "body.md").write_text("not owned", encoding="utf-8")
    original_unlink = os.unlink
    attempted = []

    def deny_resource(path, *args, **kwargs):
        if os.fspath(path) in {str(blocked_resource), blocked_resource.name}:
            attempted.append(blocked_resource.name)
            raise PermissionError("controlled resource removal failure")
        return original_unlink(path, *args, **kwargs)

    first = None
    second = None
    try:
        with monkeypatch.context() as fault:
            fault.setattr(os, "unlink", deny_resource)
            first = ClipboardInputStore(root)
        assert attempted == [blocked_resource.name]
        assert not first_resource.exists(), "The recovery must really have removed part of the bundle"
        assert blocked_resource.read_bytes() == b"resource-bytes-1"
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
