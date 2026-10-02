"""A local host consumes one platform without weakening the full publication gate."""

from __future__ import annotations

import copy
import hashlib
import io
import json
import subprocess
import zipfile
from pathlib import Path

import pytest
from scripts.release import publication
from scripts.release.publication_contract import CHECKSUM_NAME, MANIFEST_NAME, PublicationError, verify_inventory
from tests.support.publication import COMMIT, REPOSITORY, VERSION, FakeGitHub, candidate

pytestmark = pytest.mark.unit


class PlatformGitHub(FakeGitHub):
    def __init__(self, directory: Path) -> None:
        super().__init__()
        self.archives: dict[int, bytes] = {}
        self.platform_artifacts: list[dict] = []
        self.provenance: list[str] = []
        self.fail_provenance = False
        groups = {
            21: ("publication-metadata", {MANIFEST_NAME, CHECKSUM_NAME}),
            22: ("windows", {"DocWen-windows-x64.zip"}),
            23: ("linux", {f"DocWen-{VERSION}-linux-x64.tar.gz", f"DocWenCLI-{VERSION}-linux-x64.tar.gz"}),
        }
        for artifact_id, (label, names) in groups.items():
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w") as archive:
                for name in sorted(names):
                    archive.writestr(name, (directory / name).read_bytes())
            self.set_archive(artifact_id, buffer.getvalue(), label)

    def set_archive(self, artifact_id: int, content: bytes, label: str) -> None:
        self.archives[artifact_id] = content
        self.platform_artifacts[:] = [item for item in self.platform_artifacts if item["id"] != artifact_id]
        self.platform_artifacts.append(
            {
                "id": artifact_id,
                "name": f"docwen-{label}-10-1",
                "expired": False,
                "workflow_run": {"id": 10},
                "size_in_bytes": len(content),
                "digest": "sha256:" + hashlib.sha256(content).hexdigest(),
            }
        )

    def get(self, path: str, *, allow_missing: bool = False):
        if path.endswith("/artifacts?per_page=100"):
            self.reads.append(path)
            return {"total_count": len(self.platform_artifacts), "artifacts": copy.deepcopy(self.platform_artifacts)}
        return super().get(path, allow_missing=allow_missing)

    def execute(self, args, **kwargs):
        if args[:2] == ["gh", "api"]:
            artifact_id = int(args[2].split("/")[-2])
            self.downloads.append(str(artifact_id))
            # There is deliberately no whole-publication archive in this fake.
            kwargs["stdout"].write(self.archives[artifact_id])
            return subprocess.CompletedProcess(args, 0, stdout="", stderr="")
        assert args[:3] == ["gh", "attestation", "verify"]
        assert args[args.index("--source-digest") + 1] == COMMIT
        assert args[args.index("--source-ref") + 1] == f"refs/tags/{VERSION}"
        assert args[args.index("--signer-workflow") + 1] == f"{REPOSITORY}/.github/workflows/release.yml"
        assert "--deny-self-hosted-runners" in args
        self.provenance.append(Path(args[3]).name)
        return subprocess.CompletedProcess(args, int(self.fail_provenance), stdout="", stderr="")


@pytest.fixture
def platform_candidate(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    directory, _ = candidate(tmp_path)
    api = PlatformGitHub(directory)
    monkeypatch.setattr(publication, "GitHub", lambda _repository: api)
    monkeypatch.setattr(publication.subprocess, "run", api.execute)
    output = tmp_path / "selected"
    receipt = tmp_path / "receipt.json"
    args = [
        "--repository",
        REPOSITORY,
        "--version",
        VERSION,
        "--commit",
        COMMIT,
        "--artifact-id",
        "20",
        "--directory",
        str(output),
        "--receipt",
        str(receipt),
    ]
    return api, output, receipt, args


@pytest.mark.parametrize("platform,artifact_id,count", [("windows", "22", 1), ("linux", "23", 2)])
def test_selected_platform_download_and_verified_cache_never_fetch_other_platforms(
    platform_candidate, platform, artifact_id, count
) -> None:
    api, output, receipt, args = platform_candidate
    assert publication.main(["fetch-platform", *args, "--platform", platform]) == 0
    assert api.downloads == ["21", artifact_id]
    result = json.loads(receipt.read_text("utf-8"))
    assert result["stage"] == "platform-candidate-verified"
    assert result["platform"] == platform and len(result["assets"]) == count
    assert result["artifactId"] == 20 and result["platformArtifact"]["id"] == int(artifact_id)
    assert set(api.provenance) == {MANIFEST_NAME, CHECKSUM_NAME, *result["assets"]}
    assert len(list(output.iterdir())) == count + 2
    # A partial directory is never accepted as a complete publication candidate.
    with pytest.raises(PublicationError, match="unexpected candidate files"):
        verify_inventory(output, repository=REPOSITORY, version=VERSION, commit=COMMIT)
    assert publication.main(["fetch-platform", *args, "--platform", platform]) == 0
    assert api.downloads == ["21", artifact_id]
    assert not api.writes
    assert any("/runs/10/attempts/1/jobs" in path for path in api.reads)


@pytest.mark.parametrize(
    "failure", ["missing", "duplicate", "expired", "other-run", "other-attempt", "digest", "oversized"]
)
def test_metadata_must_be_the_unique_available_sibling_before_download(platform_candidate, failure) -> None:
    api, output, receipt, args = platform_candidate
    metadata = api.platform_artifacts[0]
    if failure == "missing":
        api.platform_artifacts.pop(0)
    elif failure == "duplicate":
        api.platform_artifacts.append(metadata.copy())
    elif failure == "expired":
        metadata["expired"] = True
    elif failure == "other-run":
        metadata["workflow_run"]["id"] = 11
    elif failure == "other-attempt":
        metadata["name"] = "docwen-publication-metadata-10-2"
    elif failure == "digest":
        metadata["digest"] = "invalid"
    else:
        metadata["size_in_bytes"] = 3 * 1024**2
    assert publication.main(["fetch-platform", *args, "--platform", "windows"]) == 1
    assert api.downloads == [] and not receipt.exists() and not output.exists()


@pytest.mark.parametrize(
    "failure",
    ["failed-job", "wrong-commit", "wrong-attempt", "bad-provenance", "transport-corruption", "content-corruption"],
)
def test_no_verified_receipt_for_wrong_origin_or_bytes(platform_candidate, failure) -> None:
    api, _, receipt, args = platform_candidate
    if failure == "failed-job":
        api.jobs["jobs"][-1]["conclusion"] = "failure"
    elif failure == "wrong-commit":
        api.run["head_sha"] = "f" * 40
    elif failure == "wrong-attempt":
        api.run["run_attempt"] = 2
    elif failure == "bad-provenance":
        api.fail_provenance = True
    elif failure == "transport-corruption":
        api.archives[22] += b"corrupt"
    else:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("DocWen-windows-x64.zip", b"wrong package from same transport")
        api.set_archive(22, buffer.getvalue(), "windows")
    assert publication.main(["fetch-platform", *args, "--platform", "windows"]) == 1
    assert not receipt.exists() and not api.writes
    if failure in {"failed-job", "wrong-commit", "wrong-attempt", "bad-provenance"}:
        assert "22" not in api.downloads


@pytest.mark.parametrize(
    "failure", ["package", "checksum", "extra-file", "missing-file", "receipt", "revoked-job", "source-change"]
)
def test_reuse_revalidates_local_and_remote_identity_without_silent_redownload(platform_candidate, failure) -> None:
    api, output, receipt, args = platform_candidate
    assert publication.main(["fetch-platform", *args, "--platform", "windows"]) == 0
    before = receipt.read_bytes()
    if failure == "package":
        (output / "DocWen-windows-x64.zip").write_bytes(b"changed")
    elif failure == "checksum":
        (output / CHECKSUM_NAME).write_bytes(b"changed")
    elif failure == "extra-file":
        (output / "extra.txt").write_bytes(b"unexpected")
    elif failure == "missing-file":
        (output / "DocWen-windows-x64.zip").unlink()
    elif failure == "receipt":
        receipt.write_bytes(b"another candidate")
        before = receipt.read_bytes()
    elif failure == "revoked-job":
        api.jobs["jobs"][-1]["conclusion"] = "cancelled"
    else:
        api.run["head_sha"] = "e" * 40
    assert publication.main(["fetch-platform", *args, "--platform", "windows"]) == 1
    assert api.downloads == ["21", "22"] and receipt.read_bytes() == before


@pytest.mark.parametrize("name", ["../candidate.json", "DocWen-windows-x64.zip/", "unexpected.exe"])
def test_platform_archive_inventory_prevents_unexpected_extraction(platform_candidate, name) -> None:
    api, output, receipt, args = platform_candidate
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(name, b"unexpected")
    api.set_archive(22, buffer.getvalue(), "windows")
    assert publication.main(["fetch-platform", *args, "--platform", "windows"]) == 1
    assert not receipt.exists() and not (output.parent / "candidate.json").exists()
