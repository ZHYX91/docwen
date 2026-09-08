from __future__ import annotations

import copy
import hashlib
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from scripts.release.publication_contract import assemble, canonical_json, file_identity, package_names
from scripts.release.publication_http import ApiError, GitHub

REPOSITORY = "example/docwen"
VERSION = "0.10.0"
COMMIT = "a" * 40
DIGEST = "sha256:" + "b" * 64


def candidate(root: Path) -> tuple[Path, dict]:
    source = root / "source"
    for name in ("uv.lock", "release/windows-production-manifest.v1.json", "release/linux-production-manifest.v1.json"):
        path = source / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(name, encoding="utf-8")
    builds = root / "builds"
    for name in package_names(VERSION):
        platform = "windows" if name.endswith(".zip") else "linux"
        for replica in ("a", "b"):
            path = builds / f"{platform}-{replica}" / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(name.encode("utf-8"))
    for replica in ("a", "b"):
        package = builds / f"windows-{replica}" / "DocWen-windows-x64.msix"
        package.write_bytes(b"msix")
        record = {
            **file_identity(package),
            "sourceVersion": VERSION,
            "packageVersion": "1.0.3.0",
            "contentSha256": "c" * 64,
        }
        package.with_suffix(".msix.json").write_bytes(canonical_json(record))
    output = root / "candidate"
    manifest = assemble(
        builds, output, repo=REPOSITORY, version=VERSION, commit=COMMIT, run_id=10, attempt=1, source=source
    )
    return output, manifest


class Clock:
    def __init__(self) -> None:
        self.now = 0.0
        self.delays: list[float] = []

    def time(self) -> float:
        return self.now

    def sleep(self, duration: float) -> None:
        self.delays.append(duration)
        self.now += duration


class FakeGitHub(GitHub):
    def __init__(self) -> None:
        self.repository = REPOSITORY
        self.release: dict | None = None
        self.writes: list[tuple[str, str]] = []
        self.lost: set[str] = set()
        self.hidden_reads = 0
        self.bytes: dict[int, bytes] = {}
        self.bad_download = False
        self.run = {
            "id": 10,
            "run_attempt": 1,
            "status": "completed",
            "conclusion": "success",
            "head_sha": COMMIT,
            "path": ".github/workflows/release.yml",
            "event": "workflow_dispatch",
            "repository": {"full_name": REPOSITORY},
        }
        self.artifact = {
            "id": 20,
            "expired": False,
            "digest": DIGEST,
            "workflow_run": {"id": 10},
            "name": "docwen-publication-10-1",
        }
        names = [
            "source-checks",
            "pytest_windows_push",
            "Required checks",
            "Source gate on windows-latest",
            "Source gate on ubuntu-24.04",
            "Source gate on macos-14",
            "Clean Windows build a",
            "Clean Windows build b",
            "Clean Ubuntu build a",
            "Clean Ubuntu build b",
            "verify-release",
        ]
        self.jobs = {"total_count": len(names), "jobs": [{"name": name, "conclusion": "success"} for name in names]}

    def get(self, path: str, *, allow_missing: bool = False):
        return self.request("GET", path, timeout=60)

    def request(self, method: str, path: str, *, timeout: float, body: object = None, data: bytes | None = None):
        if method == "GET":
            if path.endswith("/immutable-releases"):
                return {"enabled": True}
            if "/jobs?" in path:
                return copy.deepcopy(self.jobs)
            if "/actions/runs/" in path:
                return copy.deepcopy(self.run)
            if "/actions/artifacts/" in path:
                return copy.deepcopy(self.artifact)
            if "/git/ref/" in path:
                return {"object": {"type": "commit", "sha": COMMIT}}
            if self.release is None or self.hidden_reads:
                self.hidden_reads = max(0, self.hidden_reads - 1)
                raise ApiError(404)
            return copy.deepcopy(self.release)
        self.writes.append((method, path))
        if method == "POST" and path.endswith("/releases"):
            assert isinstance(body, dict)
            self.release = {
                **body,
                "id": 30,
                "assets": [],
                "immutable": False,
                "published_at": None,
                "html_url": "https://example.invalid/release",
            }
            operation = "create"
        elif method == "POST":
            assert self.release is not None and data is not None
            name = parse_qs(urlparse(path).query)["name"][0]
            asset_id = 40 + len(self.bytes)
            self.bytes[asset_id] = data
            self.release["assets"].append(
                {
                    "id": asset_id,
                    "name": name,
                    "state": "uploaded",
                    "size": len(data),
                    "digest": "sha256:" + hashlib.sha256(data).hexdigest(),
                }
            )
            operation = "upload"
        else:
            assert self.release is not None
            self.release.update(draft=False, immutable=True, published_at="2026-09-08T00:00:00Z")
            operation = "publish"
        if operation in self.lost:
            self.lost.remove(operation)
            self.hidden_reads = 2
            raise TimeoutError("response lost after remote write")
        return copy.deepcopy(self.release)

    def download_identity(self, path: str, *, timeout: float) -> dict:
        data = self.bytes[int(path.rsplit("/", 1)[1])]
        return {"bytes": len(data), "sha256": "0" * 64 if self.bad_download else hashlib.sha256(data).hexdigest()}
