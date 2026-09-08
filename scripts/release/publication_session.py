"""Draft, upload, publish, and independent readback of one exact candidate."""

from __future__ import annotations

import hashlib
import http.client
import json
import os
import subprocess
from pathlib import Path
from typing import Any
from urllib.parse import quote

from scripts.release.publication_contract import (
    MANIFEST_NAME,
    WORKFLOW,
    PublicationError,
    canonical_json,
    file_identity,
    publication_assets,
    read_object,
    require,
    verify_inventory,
    verify_origin,
)
from scripts.release.publication_http import ApiError, GitHub, PendingRead, read_with_retry


def verify_preflight_jobs(payload: dict[str, Any]) -> None:
    jobs = payload.get("jobs", [])
    require(payload.get("total_count") == len(jobs), "incomplete preflight job response")
    expected = {
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
    }
    for name in expected:
        matching = [job for job in jobs if job.get("name", "").split(" / ")[-1] == name]
        require(
            len(matching) == 1 and matching[0].get("conclusion") == "success",
            f"required preflight job did not pass: {name}",
        )


class ReleaseSession:
    def __init__(
        self,
        api: GitHub,
        directory: Path,
        receipt: Path,
        *,
        repository: str,
        version: str,
        commit: str,
        artifact_id: int,
        artifact_digest: str,
    ) -> None:
        self.api = api
        self.directory = directory
        self.receipt = receipt
        self.repository, self.version, self.commit = repository, version, commit
        self.prefix = f"/repos/{repository}"
        self.manifest = verify_inventory(directory, repository=repository, version=version, commit=commit)
        self.assets = publication_assets(directory, self.manifest)
        self.identity = {
            "schema": "docwen-publication-progress-v1",
            "repository": repository,
            "version": version,
            "sourceCommit": commit,
            "manifestSha256": file_identity(directory / MANIFEST_NAME)["sha256"],
            "artifactId": artifact_id,
            "artifactDigest": artifact_digest,
        }
        self.state = (
            read_object(receipt) if receipt.exists() else {**self.identity, "stage": "prepared", "pending": None}
        )
        require(
            all(self.state.get(key) == value for key, value in self.identity.items()),
            "publication receipt identity mismatch",
        )
        require(receipt.parent.is_dir() and not receipt.is_symlink(), "receipt must be in an existing owned directory")

    def save(self, **fields: Any) -> None:
        self.state.update(fields)
        replacement = self.receipt.with_suffix(".next")
        require(not replacement.exists(), "publication receipt replacement already exists")
        with replacement.open("xb") as stream:
            stream.write(canonical_json(self.state))
            stream.flush()
            os.fsync(stream.fileno())
        replacement.replace(self.receipt)

    def verify_source(self) -> None:
        origin = self.manifest["origin"]
        run = self.api.get(f"{self.prefix}/actions/runs/{origin['runId']}")
        artifact = self.api.get(f"{self.prefix}/actions/artifacts/{self.identity['artifactId']}")
        require(artifact.get("id") == self.identity["artifactId"], "artifact ID mismatch")
        verify_origin(self.manifest, run, artifact, digest=self.identity["artifactDigest"])
        verify_preflight_jobs(
            self.api.get(
                f"{self.prefix}/actions/runs/{origin['runId']}/attempts/{origin['runAttempt']}/jobs?per_page=100"
            )
        )
        self.verify_tag()

    def verify_tag(self) -> None:
        record = self.api.get(f"{self.prefix}/git/ref/tags/{quote(self.version, safe='')}")
        target = record.get("object", {})
        for _ in range(4):
            if target.get("type") == "commit":
                require(target.get("sha") == self.commit, "release tag source mismatch")
                return
            require(target.get("type") == "tag", "release tag is not a commit")
            target = self.api.get(f"{self.prefix}/git/tags/{target['sha']}").get("object", {})
        raise PublicationError("release tag nesting exceeds limit")

    def verify_provenance(self) -> None:
        for name in (*self.assets, MANIFEST_NAME):
            result = subprocess.run(
                [
                    "gh",
                    "attestation",
                    "verify",
                    str(self.directory / name),
                    "--repo",
                    self.repository,
                    "--signer-workflow",
                    f"{self.repository}/{WORKFLOW}",
                    "--source-digest",
                    self.commit,
                    "--source-ref",
                    "refs/heads/main",
                    "--deny-self-hosted-runners",
                ],
                capture_output=True,
                text=True,
                timeout=60,
            )
            require(result.returncode == 0, f"build provenance verification failed: {name}")
        self.save(provenance="verified")

    def validate_release(self, release: dict[str, Any]) -> None:
        require(type(release.get("id")) is int and release["id"] > 0, "release ID missing")
        require(
            release.get("tag_name") == self.version and release.get("prerelease") is False, "release identity mismatch"
        )
        known_id = self.state.get("releaseId")
        require(known_id is None or release["id"] == known_id, "release ID changed")
        if release.get("draft"):
            require(self.binding() in release.get("body", ""), "draft belongs to another candidate")
        else:
            require(release.get("draft") is False, "release state missing")

    def binding(self) -> str:
        return f"<!-- docwen-candidate:{self.identity['manifestSha256']} -->"

    def load_release(self, *, wait: bool = False) -> dict[str, Any]:
        release_id = self.state.get("releaseId")
        suffix = str(release_id) if release_id else f"tags/{quote(self.version, safe='')}"

        def read(timeout: float) -> dict[str, Any]:
            result = self.api.request("GET", f"{self.prefix}/releases/{suffix}", timeout=timeout)
            self.validate_release(result)
            return result

        return read_with_retry(read, allow_missing=wait)

    def write_once(
        self, operation: str, method: str, path: str, *, body: object = None, data: bytes | None = None
    ) -> None:
        pending = self.state.get("pending")
        require(pending in (None, operation), "another publication write requires reconciliation")
        if pending is not None:
            return  # The caller must resolve the previous remote outcome by reading.
        self.save(pending=operation)
        try:
            result = self.api.request(method, path, timeout=600 if data is not None else 60, body=body, data=data)
            if operation == "create-draft":
                self.validate_release(result)
                self.save(releaseId=result["id"], stage="draft-created", pending=None)
            elif operation == "publish":
                self.validate_release(result)
                require(result.get("draft") is False, "publish response still describes a draft")
                self.save(stage="published-awaiting-readback", pending=None)
        except ApiError as error:
            if not error.transient and error.status != 422:
                self.save(pending=None)
                raise
            # A write is never automatically repeated, including after a lost response.
        except (TimeoutError, OSError, http.client.HTTPException, json.JSONDecodeError):
            pass

    def remote_assets(self, release: dict[str, Any], *, complete: bool) -> dict[str, dict[str, Any]]:
        records = release.get("assets")
        if not isinstance(records, list):
            raise PendingRead("release assets are not visible yet")
        result: dict[str, dict[str, Any]] = {}
        for asset in records:
            name = asset.get("name")
            require(name in self.assets and name not in result, "unexpected or duplicate remote asset")
            if asset.get("state") != "uploaded" or not asset.get("digest"):
                raise PendingRead("asset upload is not visible yet")
            expected = self.assets[name]
            require(
                asset.get("size") == expected["bytes"] and asset.get("digest") == f"sha256:{expected['sha256']}",
                f"remote asset mismatch: {name}",
            )
            require(type(asset.get("id")) is int and asset["id"] > 0, "asset ID missing")
            result[name] = asset
        if complete and set(result) != set(self.assets):
            raise PendingRead("release inventory is not complete yet")
        return result

    def observe(self, *, published: bool, complete: bool) -> tuple[dict[str, Any], dict[str, Any]]:
        release_id = self.state["releaseId"]

        def read(timeout: float):
            release = self.api.request("GET", f"{self.prefix}/releases/{release_id}", timeout=timeout)
            self.validate_release(release)
            if published and (
                release.get("draft") is not False
                or release.get("immutable") is not True
                or not release.get("published_at")
            ):
                raise PendingRead("immutable published state is not visible yet")
            if not published:
                require(release.get("draft") is True, "draft was published before verification completed")
            return release, self.remote_assets(release, complete=complete)

        return read_with_retry(read, allow_missing=True)

    def publish(self, *, notes: str) -> dict[str, Any]:
        self.verify_source()
        self.verify_provenance()
        admin_token = os.environ.get("DOCWEN_IMMUTABILITY_READ_TOKEN")
        settings_api = GitHub(self.repository, token=admin_token) if admin_token else self.api
        require(
            settings_api.get(f"{self.prefix}/immutable-releases").get("enabled") is True,
            "immutable releases must be enabled",
        )
        self.save()
        try:
            release = self.load_release(wait=self.state.get("pending") == "create-draft")
        except ApiError as error:
            if error.status != 404 or self.state.get("releaseId") is not None:
                raise
            self.write_once(
                "create-draft",
                "POST",
                f"{self.prefix}/releases",
                body={
                    "tag_name": self.version,
                    "target_commitish": self.commit,
                    "name": self.version,
                    "body": f"{notes.rstrip()}\n\n{self.binding()}",
                    "draft": True,
                    "prerelease": False,
                },
            )
            release = self.load_release(wait=True)
        self.save(releaseId=release["id"])
        if self.state.get("pending") == "create-draft":
            self.save(pending=None, stage="draft-created")
        if release.get("draft") is False:
            self.save(stage="published-awaiting-readback", pending=None)
            return self.verify_published()
        self.validate_release(release)
        for name in self.assets:
            release, remote = self.observe(published=False, complete=False)
            operation = f"upload:{name}"
            if name not in remote:
                require(file_identity(self.directory / name) == self.assets[name], f"local asset changed: {name}")
                content = (self.directory / name).read_bytes()
                require(
                    len(content) == self.assets[name]["bytes"]
                    and hashlib.sha256(content).hexdigest() == self.assets[name]["sha256"],
                    f"upload bytes changed: {name}",
                )
                self.write_once(
                    operation,
                    "POST",
                    f"{self.prefix}/releases/{release['id']}/assets?name={quote(name, safe='')}",
                    data=content,
                )

                def read_asset(timeout: float, name=name, release_id=release["id"]):
                    observed = self.api.request("GET", f"{self.prefix}/releases/{release_id}", timeout=timeout)
                    self.validate_release(observed)
                    assets = self.remote_assets(observed, complete=False)
                    if name not in assets:
                        raise PendingRead("uploaded asset is not visible yet")
                    return assets[name]

                read_with_retry(read_asset, allow_missing=True)
            if self.state.get("pending") == operation:
                self.save(pending=None)
        self.observe(published=False, complete=True)
        self.save(stage="draft-verified")
        self.verify_tag()
        self.write_once(
            "publish",
            "PATCH",
            f"{self.prefix}/releases/{self.state['releaseId']}",
            body={"draft": False, "make_latest": "true"},
        )
        return self.verify_published()

    def verify_published(self) -> dict[str, Any]:
        self.verify_tag()
        release, assets = self.observe(published=True, complete=True)
        self.save(stage="published-awaiting-readback", pending=None)
        for name, record in assets.items():
            observed = read_with_retry(
                lambda timeout, record=record: self.api.download_identity(
                    f"{self.prefix}/releases/assets/{record['id']}",
                    timeout=timeout,
                )
            )
            require(observed == self.assets[name], f"remote bytes mismatch: {name}")
        self.verify_tag()
        self.save(
            stage="verified",
            releaseUrl=release.get("html_url"),
            assets={name: {"id": record["id"], **self.assets[name]} for name, record in assets.items()},
        )
        return self.state
