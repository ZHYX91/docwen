"""Bounded, identity-bound artifact transfers; signed URLs are never persisted."""

from __future__ import annotations

import base64
import hashlib
import http.client
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from urllib.parse import unquote, urlsplit
from urllib.request import getproxies, proxy_bypass

from scripts.release.publication_contract import PublicationError, canonical_json, file_identity, read_object, require
from scripts.release.publication_http import ApiError, GitHub, env_seconds, retry_after_seconds


class StorageRetry(OSError):
    def __init__(self, retry_after: float) -> None:
        super().__init__("temporary artifact storage response")
        self.retry_after = retry_after


def regular_file(path: Path) -> None:
    require(
        path.is_file() and not path.is_symlink() and not getattr(path.stat(), "st_file_attributes", 0) & 0x400,
        "download file must be regular and not a reparse point",
    )


def pending_record(path: Path, content: bytes) -> Path:
    return path.with_name(path.name + ".writing-" + hashlib.sha256(content).hexdigest())


def atomic_record(path: Path, content: bytes) -> None:
    """Commit an exact record without exposing partial final bytes or replacing foreign data."""
    temporary = pending_record(path, content)
    if path.exists() or path.is_symlink():
        regular_file(path)
        require(path.read_bytes() == content, "record identity mismatch")
    else:
        if temporary.exists() or temporary.is_symlink():
            regular_file(temporary)
            require(temporary.stat().st_size <= len(content), "unexpected pending record size")
            require(content.startswith(temporary.read_bytes()), "pending record identity mismatch")
            temporary.unlink()
        with temporary.open("xb") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        os.link(temporary, path)
    if temporary.exists() or temporary.is_symlink():
        regular_file(temporary)
        require(temporary.stat().st_size <= len(content), "unexpected pending record size")
        require(content.startswith(temporary.read_bytes()), "pending record identity mismatch")
        temporary.unlink()


def open_response(url: str, offset: int, *, connect_timeout: float, idle_timeout: float, deadline: float | None = None):
    """Only the API request has credentials; storage receives no Authorization header."""
    parsed = urlsplit(url)
    require(
        parsed.scheme == "https"
        and bool(parsed.hostname)
        and not parsed.username
        and not parsed.password
        and not parsed.fragment,
        "invalid artifact storage URL",
    )
    proxies = getproxies()
    authority = parsed.netloc if parsed.port is not None else parsed.netloc + ":443"
    proxy_url = (proxies.get("https") or proxies.get("all")) if not proxy_bypass(authority) else None
    if proxy_url:
        proxy = urlsplit(proxy_url if "://" in proxy_url else "http://" + proxy_url)
        require(proxy.scheme == "http" and bool(proxy.hostname), "HTTPS artifacts require an HTTP CONNECT proxy")
        connection = http.client.HTTPSConnection(proxy.hostname, proxy.port or 80, timeout=connect_timeout)
        tunnel_headers = {}
        if proxy.username is not None:
            credentials = unquote(proxy.username) + ":" + unquote(proxy.password or "")
            tunnel_headers["Proxy-Authorization"] = "Basic " + base64.b64encode(credentials.encode()).decode("ascii")
        connection.set_tunnel(parsed.hostname, parsed.port or 443, headers=tunnel_headers)
    else:
        connection = http.client.HTTPSConnection(parsed.hostname, parsed.port, timeout=connect_timeout)
    try:
        connection.connect()
        assert connection.sock is not None
        transport_socket = connection.sock
        if deadline is not None:
            remaining = deadline - time.monotonic()
            require(remaining > 0, "artifact total download deadline exceeded; partial retained")
            idle_timeout = min(idle_timeout, remaining)
        transport_socket.settimeout(idle_timeout)
        headers = {"Accept-Encoding": "identity", "User-Agent": "DocWen-release"}
        if offset:
            headers["Range"] = f"bytes={offset}-"
        path = parsed.path or "/"
        connection.request("GET", path + ("?" + parsed.query if parsed.query else ""), headers=headers)
        response = connection.getresponse()
        return connection, response, transport_socket
    except Exception:
        connection.close()
        raise


def download_artifact_in_process(api: GitHub, artifact: dict, archive: Path) -> None:
    """Resume only a recorded immutable artifact, with bounded retries and a final full hash."""
    size = artifact.get("size_in_bytes")
    require(type(size) is int and 0 < size <= 8 * 1024**3, "invalid artifact size")
    identity = {"repository": api.repository, "id": artifact["id"], "digest": artifact["digest"], "bytes": size}
    state = archive.with_suffix(".json")
    if state.exists():
        regular_file(state)
        require(read_object(state) == identity, "partial download identity mismatch")
    else:
        require(not archive.exists() and not archive.is_symlink(), "unowned partial download")
    atomic_record(state, canonical_json(identity))
    if not archive.exists():
        with archive.open("xb"):
            pass
    regular_file(archive)
    require(archive.stat().st_size <= size, "partial download exceeds artifact size")
    deadline = time.monotonic() + env_seconds("DOCWEN_PUBLICATION_ARTIFACT_TIMEOUT", 7200)
    connect = env_seconds("DOCWEN_PUBLICATION_CONNECT_TIMEOUT", 30)
    idle = env_seconds("DOCWEN_PUBLICATION_IDLE_TIMEOUT", 120)
    last_progress = 0.0
    for attempt in range(5):
        offset = archive.stat().st_size
        if offset == size:
            break
        remaining = deadline - time.monotonic()
        require(remaining > 0, "artifact total download deadline exceeded; partial retained")
        connection = response = None
        try:
            url = api.artifact_url(artifact["id"], timeout=min(connect, remaining))
            remaining = deadline - time.monotonic()
            require(remaining > 0, "artifact total download deadline exceeded; partial retained")
            connection, response, transport_socket = open_response(
                url,
                offset,
                connect_timeout=min(connect, remaining),
                idle_timeout=min(idle, remaining),
                deadline=deadline,
            )
            if response.status in {403, 408, 429, 500, 502, 503, 504}:
                # Refresh expiring signed URLs on the next attempt; never log them.
                raise StorageRetry(retry_after_seconds(response.getheader("Retry-After")))
            require(response.status == (206 if offset else 200), "storage did not honor artifact byte range")
            if offset:
                match = re.fullmatch(r"bytes (\d+)-(\d+)/(\d+)", response.getheader("Content-Range", ""))
                require(
                    match is not None and tuple(map(int, match.groups())) == (offset, size - 1, size),
                    "artifact Content-Range mismatch",
                )
            length = response.getheader("Content-Length")
            require(length is None or length == str(size - offset), "artifact response size mismatch")
            require(response.getheader("Content-Encoding", "identity") == "identity", "encoded artifact response")
            with archive.open("ab") as stream:
                while offset < size:
                    remaining = deadline - time.monotonic()
                    require(remaining > 0, "artifact total download deadline exceeded; partial retained")
                    # read1 returns available data instead of waiting for an entire large buffer.
                    transport_socket.settimeout(min(idle, remaining))
                    chunk = response.read1(min(1024 * 1024, size - offset))
                    if not chunk:
                        raise OSError("artifact stream ended early")
                    stream.write(chunk)
                    offset += len(chunk)
                    now = time.monotonic()
                    if now - last_progress >= 30 or offset == size:
                        print(
                            f"Artifact {artifact['id']}: {offset}/{size} bytes (attempt {attempt + 1}/5)",
                            file=sys.stderr,
                            flush=True,
                        )
                        last_progress = now
            require(time.monotonic() <= deadline, "artifact total download deadline exceeded; partial retained")
            break
        except (OSError, http.client.HTTPException, ApiError) as error:
            if isinstance(error, ApiError) and not error.transient:
                raise
            remaining = deadline - time.monotonic()
            delay = max(min(2 ** (attempt + 1), 16), getattr(error, "retry_after", 0))
            require(attempt < 4 and remaining > delay, "artifact recovery exhausted; partial retained for retry")
            print(
                f"Artifact {artifact['id']}: transfer interrupted; resuming verified identity",
                file=sys.stderr,
                flush=True,
            )
            time.sleep(delay)
        finally:
            if response is not None:
                response.close()
            if connection is not None:
                connection.close()
    actual = file_identity(archive)
    require(
        actual["bytes"] == size and "sha256:" + actual["sha256"] == artifact["digest"],
        "downloaded artifact digest mismatch",
    )


def finish_download(archive: Path) -> None:
    """Remove only transfer scratch after its verified contents have been extracted."""
    regular_file(archive)
    state = archive.with_suffix(".json")
    regular_file(state)
    archive.unlink()
    state.unlink()


def download_artifact(api: GitHub, artifact: dict, archive: Path) -> None:
    """A killable worker bounds DNS, TLS, slow headers, body, and retry waits."""
    budget = env_seconds("DOCWEN_PUBLICATION_ARTIFACT_TIMEOUT", 7200)
    payload = json.dumps(
        {"repository": api.repository, "token": api._token, "artifact": artifact, "archive": str(archive.absolute())}
    )
    try:
        result = subprocess.run(
            [sys.executable, "-m", "scripts.release.publication_download"],
            input=payload,
            text=True,
            stdout=subprocess.PIPE,
            cwd=Path(__file__).resolve().parents[2],
            timeout=budget,
            check=False,
        )
    except subprocess.TimeoutExpired:
        # subprocess.run kills and waits for this single worker; there are no
        # background readers still writing the partial archive after return.
        raise PublicationError("artifact total download deadline exceeded; partial retained") from None
    require(result.returncode == 0, "artifact transfer failed; identity-bound partial retained")


def _worker_main() -> None:
    request = json.load(sys.stdin)
    api = object.__new__(GitHub)
    api.repository = request["repository"]
    api._token = request["token"]
    try:
        download_artifact_in_process(api, request["artifact"], Path(request["archive"]))
    except Exception as error:
        # Never render transport exceptions: they can contain signed URLs.
        print(str(error) if isinstance(error, PublicationError) else "artifact transport failed", file=sys.stderr)
        raise SystemExit(1) from None


if __name__ == "__main__":
    _worker_main()
