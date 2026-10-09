"""Single-platform MSIX provenance and encoded SDK member regression cases."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
from scripts.release.msix_candidate import extract_candidate
from scripts.release.msix_layout import verify_layout
from scripts.release.publication_contract import PublicationError, file_identity

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("failure", [None, "linux", "missing-origin", "missing-platform", "unverified"])
def test_single_platform_receipt_is_bound_before_unpacking(tmp_path: Path, failure: str | None) -> None:
    archive = tmp_path / "portable.zip"
    with zipfile.ZipFile(archive, "w") as bundle:
        for name in ("DocWen.exe", "DocWenCLI.exe"):
            info = zipfile.ZipInfo(name)
            info.external_attr = 0o100644 << 16
            bundle.writestr(info, b"payload")
    artifact = {"id": 42, "digest": "sha256:" + "a" * 64, "size_in_bytes": 123}
    record = {
        "schema": "docwen-platform-candidate-v1",
        "stage": "platform-candidate-verified",
        "platform": "windows",
        "provenance": "verified",
        "version": "0.17.0",
        "repository": "owner/repo",
        "sourceCommit": "b" * 40,
        "manifestSha256": "c" * 64,
        "artifactId": 11,
        "artifactDigest": "sha256:" + "d" * 64,
        "origin": {"runId": 10, "runAttempt": 1},
        "metadataArtifact": artifact,
        "platformArtifact": artifact,
        "assets": {"DocWen-windows-x64.zip": file_identity(archive)},
    }
    if failure == "linux":
        record["platform"] = "linux"
    if failure == "missing-origin":
        record.pop("origin")
    if failure == "missing-platform":
        record.pop("platformArtifact")
    if failure == "unverified":
        record["provenance"] = "unknown"
    receipt = tmp_path / "receipt.json"
    receipt.write_text(json.dumps(record))
    output = tmp_path / "payload"
    if failure:
        with pytest.raises(PublicationError):
            extract_candidate(archive, receipt, output, version="0.17.0")
        assert not output.exists()
    else:
        result = extract_candidate(archive, receipt, output, version="0.17.0")
        assert result["platformArtifact"] == artifact
        assert (output / "DocWen.exe").read_bytes() == b"payload"


@pytest.mark.parametrize("failure", [None, "collision", "changed", "extra"])
def test_msix_readback_decodes_opc_paths_and_rejects_changes(tmp_path: Path, failure: str | None) -> None:
    staging = tmp_path / "staging"
    staging.mkdir()
    (staging / "中文 name.txt").write_bytes(b"verified")
    package = tmp_path / "candidate.msix"
    with zipfile.ZipFile(package, "w") as bundle:
        bundle.writestr("%E4%B8%AD%E6%96%87%20name.txt", b"changed" if failure == "changed" else b"verified")
        for name in ("AppxBlockMap.xml", "[Content_Types].xml"):
            bundle.writestr(name, b"sdk")
        if failure == "collision":
            bundle.writestr("中文 name.txt", b"verified")
        if failure == "extra":
            bundle.writestr("unknown.txt", b"extra")
    if failure:
        with pytest.raises(PublicationError):
            verify_layout(package, staging)
    else:
        assert verify_layout(package, staging) == 1
