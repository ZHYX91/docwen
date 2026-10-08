"""Transfer recovery preserves identity and rejects ambiguous byte ranges."""

from __future__ import annotations

import hashlib
import io
import json
from pathlib import Path
from types import SimpleNamespace

import pytest
from scripts.release import publication_download as download
from scripts.release.publication_contract import PublicationError
from scripts.release.publication_http import env_seconds

pytestmark = pytest.mark.unit


class Response(io.BytesIO):
    def __init__(self, content: bytes, offset: int, total: int, *, status: int | None = None):
        super().__init__(content)
        self.status = status or (206 if offset else 200)
        self.headers = {"Content-Length": str(total - offset), "Content-Range": f"bytes {offset}-{total - 1}/{total}"}

    def getheader(self, name, default=None):
        return self.headers.get(name, default)


@pytest.fixture
def transfer(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    content = b"independent immutable artifact bytes"
    artifact = {"id": 42, "size_in_bytes": len(content), "digest": "sha256:" + hashlib.sha256(content).hexdigest()}
    calls = []
    now = [0.0]
    monkeypatch.setattr(download.time, "monotonic", lambda: now[0])
    monkeypatch.setattr(download.time, "sleep", lambda delay: now.__setitem__(0, now[0] + delay))
    api = SimpleNamespace(
        repository="owner/repo", artifact_url=lambda artifact_id, timeout: "https://storage.invalid/signed-secret"
    )

    def opener(url, offset, **kwargs):
        calls.append((offset, kwargs))
        return (
            SimpleNamespace(close=lambda: None),
            Response(content[offset:], offset, len(content)),
            SimpleNamespace(settimeout=lambda value: None),
        )

    monkeypatch.setattr(download, "open_response", opener)
    return api, artifact, tmp_path / "download-42.zip", content, calls, now, opener


def test_interrupted_transfer_resumes_range_and_hashes_complete_bytes(transfer, monkeypatch, capsys):
    api, artifact, path, content, calls, _, opener = transfer
    first = True

    def interrupted(url, offset, **kwargs):
        nonlocal first
        connection, response, sock = opener(url, offset, **kwargs)
        if first:
            first = False
            response = Response(content[:7], 0, len(content))
        return connection, response, sock

    monkeypatch.setattr(download, "open_response", interrupted)
    download.download_artifact(api, artifact, path)
    assert path.read_bytes() == content
    assert [offset for offset, _ in calls] == [0, 7]
    assert "signed-secret" not in capsys.readouterr().err


def test_exhausted_transfer_can_resume_in_a_later_invocation(transfer, monkeypatch):
    api, artifact, path, content, calls, _, opener = transfer

    def truncated(url, offset, **kwargs):
        connection, _, sock = opener(url, offset, **kwargs)
        return connection, Response(content[offset : offset + 1], offset, len(content)), sock

    monkeypatch.setattr(download, "open_response", truncated)
    with pytest.raises(PublicationError, match="recovery exhausted"):
        download.download_artifact(api, artifact, path)
    assert path.read_bytes() == content[:5]
    monkeypatch.setattr(download, "open_response", opener)
    download.download_artifact(api, artifact, path)
    assert calls[-1][0] == 5 and path.read_bytes() == content


@pytest.mark.parametrize("fault", ["identity", "range", "ignored-range", "digest", "oversized"])
def test_resume_rejects_different_identity_or_untrusted_response(transfer, monkeypatch, fault):
    api, artifact, path, content, calls, _, opener = transfer
    download.download_artifact(api, artifact, path)
    calls.clear()
    path.write_bytes(content[:5])
    if fault == "identity":
        artifact = {**artifact, "id": 43}
    elif fault == "digest":
        path.write_bytes(b"WRONG")
    elif fault == "oversized":
        path.write_bytes(content + b"extra")
    else:

        def invalid(url, offset, **kwargs):
            connection, response, sock = opener(url, offset, **kwargs)
            if fault == "range":
                response.headers["Content-Range"] = f"bytes 0-{len(content) - 1}/{len(content)}"
            else:
                response.status = 200
            return connection, response, sock

        monkeypatch.setattr(download, "open_response", invalid)
    with pytest.raises(PublicationError):
        download.download_artifact(api, artifact, path)
    if fault in {"identity", "oversized"}:
        assert not calls


def test_slow_progress_can_exceed_old_ten_minute_limit(transfer, monkeypatch):
    api, artifact, path, content, _, now, opener = transfer

    def slow(url, offset, **kwargs):
        connection, response, sock = opener(url, offset, **kwargs)

        def read1(size):
            now[0] += 50
            return response.read(1)

        response.read1 = read1
        return connection, response, sock

    monkeypatch.setattr(download, "open_response", slow)
    download.download_artifact(api, artifact, path)
    assert path.read_bytes() == content and now[0] > 600


def test_total_deadline_retains_partial_instead_of_restarting(transfer, monkeypatch):
    api, artifact, path, content, calls, now, opener = transfer
    monkeypatch.setenv("DOCWEN_PUBLICATION_ARTIFACT_TIMEOUT", "3")

    def slow(url, offset, **kwargs):
        connection, response, sock = opener(url, offset, **kwargs)

        def read1(size):
            now[0] += 1
            return response.read(1)

        response.read1 = read1
        return connection, response, sock

    monkeypatch.setattr(download, "open_response", slow)
    with pytest.raises(PublicationError, match="total download deadline"):
        download.download_artifact(api, artifact, path)
    assert path.read_bytes() == content[:3] and len(calls) == 1
    assert json.loads(path.with_suffix(".json").read_text())["id"] == 42


@pytest.mark.parametrize("value", ["nan", "inf", "-inf", "0", "-1"])
def test_timeout_overrides_must_be_finite_positive(monkeypatch, value):
    monkeypatch.setenv("DOCWEN_PUBLICATION_ARTIFACT_TIMEOUT", value)
    with pytest.raises(PublicationError):
        env_seconds("DOCWEN_PUBLICATION_ARTIFACT_TIMEOUT", 7200)


@pytest.mark.parametrize(
    "url",
    [
        "http://storage.invalid/file",
        "https://user:secret@storage.invalid/file",
        "https://storage.invalid/file#fragment",
    ],
)
def test_invalid_storage_url_rejected_before_connection(url):
    with pytest.raises(PublicationError, match="storage URL"):
        download.open_response(url, 0, connect_timeout=30, idle_timeout=120)
