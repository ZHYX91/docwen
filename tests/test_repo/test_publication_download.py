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
    download.download_artifact_in_process(api, artifact, path)
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
        download.download_artifact_in_process(api, artifact, path)
    assert path.read_bytes() == content[:5]
    monkeypatch.setattr(download, "open_response", opener)
    download.download_artifact_in_process(api, artifact, path)
    assert calls[-1][0] == 5 and path.read_bytes() == content


@pytest.mark.parametrize("fault", ["identity", "range", "ignored-range", "digest", "oversized"])
def test_resume_rejects_different_identity_or_untrusted_response(transfer, monkeypatch, fault):
    api, artifact, path, content, calls, _, opener = transfer
    download.download_artifact_in_process(api, artifact, path)
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
        download.download_artifact_in_process(api, artifact, path)
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
    download.download_artifact_in_process(api, artifact, path)
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
        download.download_artifact_in_process(api, artifact, path)
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


@pytest.mark.parametrize("bypass", [False, True])
def test_storage_http_connect_proxy_auth_is_not_forwarded_to_origin(monkeypatch, bypass):
    calls = {}
    socket = SimpleNamespace(settimeout=lambda value: None)

    class Connection:
        sock = socket

        def __init__(self, host, port, timeout):
            calls["destination"] = (host, port)

        def set_tunnel(self, host, port, headers):
            calls["tunnel"] = (host, port, headers)

        def connect(self):
            pass

        def request(self, method, path, headers):
            calls["headers"] = headers

        def getresponse(self):
            return SimpleNamespace()

        def close(self):
            pass

    monkeypatch.setattr(download, "getproxies", lambda: {"https": "http://user:p%40ss@proxy.invalid:8080"})
    monkeypatch.setattr(download, "proxy_bypass", lambda host: bypass)
    monkeypatch.setattr(download.http.client, "HTTPSConnection", Connection)
    download.open_response("https://storage.invalid/archive?secret", 7, connect_timeout=30, idle_timeout=120)
    assert calls["destination"] == (("storage.invalid", None) if bypass else ("proxy.invalid", 8080))
    assert calls["headers"] == {"Accept-Encoding": "identity", "User-Agent": "DocWen-release", "Range": "bytes=7-"}
    if not bypass:
        assert calls["tunnel"] == ("storage.invalid", 443, {"Proxy-Authorization": "Basic dXNlcjpwQHNz"})


def test_storage_retry_after_is_honored_inside_budget(transfer, monkeypatch):
    api, artifact, path, content, calls, now, opener = transfer

    def limited(url, offset, **kwargs):
        connection, response, sock = opener(url, offset, **kwargs)
        if len(calls) == 1:
            response.status = 429
            response.headers["Retry-After"] = "120"
        return connection, response, sock

    monkeypatch.setattr(download, "open_response", limited)
    download.download_artifact_in_process(api, artifact, path)
    assert now[0] == 120 and path.read_bytes() == content


def test_worker_absolute_deadline_bounds_blocked_transport(tmp_path, monkeypatch):
    import subprocess
    import sys

    real_run = subprocess.run

    def blocked_worker(command, **kwargs):
        assert kwargs["timeout"] == 1
        # DNS and slow HTTP headers are contained in this same child boundary.
        return real_run([sys.executable, "-c", "import time; time.sleep(60)"], **kwargs)

    monkeypatch.setenv("DOCWEN_PUBLICATION_ARTIFACT_TIMEOUT", "1")
    monkeypatch.setattr(download.subprocess, "run", blocked_worker)
    with pytest.raises(PublicationError, match="total download deadline"):
        download.download_artifact(
            SimpleNamespace(repository="owner/repo", _token="not-real"), {}, tmp_path / "partial.zip"
        )


@pytest.mark.parametrize(
    "proxy_key,no_proxy,expected", [("all", "", "proxy.invalid"), ("https", "storage.invalid:443", "storage.invalid")]
)
def test_proxy_fallback_and_effective_port_bypass(monkeypatch, proxy_key, no_proxy, expected):
    from urllib.request import proxy_bypass_environment

    hosts = []

    class Connection:
        sock = SimpleNamespace(settimeout=lambda value: None)

        def __init__(self, host, port, timeout):
            hosts.append(host)

        def set_tunnel(self, *args, **kwargs):
            pass

        def connect(self):
            pass

        def request(self, *args, **kwargs):
            pass

        def getresponse(self):
            return SimpleNamespace()

        def close(self):
            pass

    monkeypatch.setattr(download, "getproxies", lambda: {proxy_key: "http://proxy.invalid:8080"})
    monkeypatch.setattr(download, "proxy_bypass", lambda host: proxy_bypass_environment(host, {"no": no_proxy}))
    monkeypatch.setattr(download.http.client, "HTTPSConnection", Connection)
    download.open_response("https://storage.invalid/file", 0, connect_timeout=30, idle_timeout=120)
    assert hosts == [expected]


def test_worker_preserves_original_leaf_path_before_validation(tmp_path, monkeypatch):
    leaf = tmp_path / "download.zip"
    expected = str(leaf.absolute())
    original_resolve = Path.resolve

    def checked_resolve(self, *args, **kwargs):
        assert self != leaf, "download leaf must not be resolved before worker validation"
        return original_resolve(self, *args, **kwargs)

    def worker(command, **kwargs):
        assert json.loads(kwargs["input"])["archive"] == expected
        return SimpleNamespace(returncode=0)

    monkeypatch.setattr(Path, "resolve", checked_resolve)
    monkeypatch.setattr(download.subprocess, "run", worker)
    download.download_artifact(SimpleNamespace(repository="owner/repo", _token="not-real"), {}, leaf)


@pytest.mark.parametrize("stage", ["partial", "promoted"])
def test_atomic_record_recovers_interruption(tmp_path, stage):
    final = tmp_path / "receipt.json"
    content = b'{"identity":"fixed"}\n'
    pending = download.pending_record(final, content)
    pending.write_bytes(content[:7] if stage == "partial" else content)
    if stage == "promoted":
        final.write_bytes(content)
    download.atomic_record(final, content)
    assert final.read_bytes() == content and not pending.exists()


def test_atomic_record_rejects_foreign_pending_and_final_bytes(tmp_path):
    final = tmp_path / "receipt.json"
    content = b'{"identity":"fixed"}\n'
    pending = download.pending_record(final, content)
    pending.write_bytes(b"foreign")
    with pytest.raises(PublicationError, match="pending record identity"):
        download.atomic_record(final, content)
    assert pending.read_bytes() == b"foreign" and not final.exists()
    final.write_bytes(b"foreign final")
    with pytest.raises(PublicationError, match="record identity"):
        download.atomic_record(final, content)
    assert final.read_bytes() == b"foreign final"
