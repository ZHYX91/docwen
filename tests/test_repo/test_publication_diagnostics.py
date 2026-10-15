from __future__ import annotations

import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest
from scripts.release import publication, publication_download, publication_session
from scripts.release.publication_contract import PublicationError, command_diagnostic
from scripts.release.publication_http import ApiError, GitHub, PendingRead, read_with_retry
from tests.support.publication import COMMIT, DIGEST, REPOSITORY, VERSION, FakeGitHub

pytestmark = pytest.mark.unit

_STDERR = (
    "certificate chain unavailable: https://storage.invalid/package?sig=private-signature "
    "Authorization: Bearer ghp_privatevalue token=plain-secret\n"
)


def _assert_reason(error: BaseException, subject: str, code: str = "17") -> None:
    text = str(error)
    assert subject in text and f"exit_code={code}" in text and "certificate chain unavailable" in text
    assert not any(
        secret in text for secret in ("storage.invalid", "private-signature", "ghp_privatevalue", "plain-secret")
    )


def test_command_diagnostic_redacts_before_truncation_and_neutralizes_terminal_controls(monkeypatch) -> None:
    monkeypatch.setenv("GH_TOKEN", "arbitrary-environment-secret")
    result = command_diagnostic("subject", 17, (_STDERR + "arbitrary-environment-secret\x1b[31m" + "x" * 3000).encode())
    assert len(result) < 1250 and "truncated" in result
    assert "arbitrary-environment-secret" not in result and "\x1b" not in result and "\n" not in result
    _assert_reason(PublicationError(result), "subject")


@pytest.mark.parametrize("key", ["GH_ENTERPRISE_TOKEN", "GITHUB_ENTERPRISE_TOKEN"])
def test_diagnostic_redacts_enterprise_credentials_and_quoted_fields(monkeypatch, key: str) -> None:
    monkeypatch.setenv(key, "enterprise-opaque-value")
    result = command_diagnostic(
        "subject",
        1,
        '''enterprise-opaque-value {"token":"opaque-secret"} {'access_token': 'other-secret'} password="two words"''',
    )
    assert not any(
        secret in result for secret in ("enterprise-opaque-value", "opaque-secret", "other-secret", "two words")
    )


@pytest.mark.parametrize("timeout", [False, True])
def test_provenance_failure_preserves_subject_and_safe_stderr(tmp_path: Path, monkeypatch, timeout: bool) -> None:
    def execute(command, **kwargs):
        if timeout:
            raise subprocess.TimeoutExpired(command, 60, stderr=_STDERR.encode())
        return subprocess.CompletedProcess(command, 17, stderr=_STDERR)

    monkeypatch.setattr(publication_session.subprocess, "run", execute)
    manifest = {"repository": REPOSITORY, "sourceCommit": COMMIT, "origin": {"sourceRef": "refs/heads/candidate"}}
    with pytest.raises(PublicationError) as caught:
        publication_session.verify_provenance(tmp_path, manifest, ("candidate.json",))
    _assert_reason(caught.value, "candidate.json", "timeout" if timeout else "17")


@pytest.mark.parametrize("status", [None, 503, 404])
def test_public_asset_download_keeps_error_type_and_processed_diagnostics(monkeypatch, status: int | None) -> None:
    monkeypatch.setenv("GH_TOKEN", "test-only-token")
    api = GitHub(REPOSITORY)
    stderr = _STDERR.encode() + (f"HTTP {status}".encode() if status else b"")
    monkeypatch.setattr(subprocess, "run", lambda *args, **kwargs: SimpleNamespace(returncode=17, stderr=stderr))
    with pytest.raises(ApiError if status else PendingRead) as caught:
        api.download_identity("/repos/example/docwen/releases/assets/42", timeout=3)
    _assert_reason(caught.value, "assets/42")
    if status:
        assert caught.value.status == status
        assert caught.value.transient is (status == 503)


@pytest.mark.parametrize("timeout", [False, True])
def test_candidate_download_failure_preserves_artifact_identity(tmp_path: Path, monkeypatch, timeout: bool) -> None:
    def execute(command, **kwargs):
        if timeout:
            raise subprocess.TimeoutExpired(command, 600, stderr=_STDERR.encode())
        return subprocess.CompletedProcess(command, 17, stderr=_STDERR.encode())

    monkeypatch.setattr(publication.subprocess, "run", execute)
    with pytest.raises(PublicationError) as caught:
        publication.fetch_candidate(
            FakeGitHub(),
            output=tmp_path / "candidate",
            artifact_id=20,
            digest=DIGEST,
            repository=REPOSITORY,
            version=VERSION,
            commit=COMMIT,
        )
    _assert_reason(caught.value, "20", "timeout" if timeout else "17")


def test_transfer_worker_failure_keeps_final_diagnostic_without_a_retry(tmp_path: Path, monkeypatch) -> None:
    calls = []

    def execute(command, **kwargs):
        calls.append(command)
        return SimpleNamespace(returncode=17, stdout=_STDERR)

    monkeypatch.setattr(publication_download.subprocess, "run", execute)
    with pytest.raises(PublicationError) as caught:
        publication_download.download_artifact(
            SimpleNamespace(repository=REPOSITORY, _token="unused"), {"id": 42}, tmp_path / "partial.zip"
        )
    _assert_reason(caught.value, "42")
    assert len(calls) == 1


def test_exhausted_read_retains_last_download_diagnostic_without_changing_retry_budget() -> None:
    calls = []

    def fail(timeout):
        calls.append(timeout)
        raise PendingRead(command_diagnostic("asset download failed: assets/42", 17, _STDERR))

    with pytest.raises(PublicationError, match="read recovery budget exhausted") as caught:
        read_with_retry(fail, budget=1, clock=lambda: 0)
    _assert_reason(caught.value, "assets/42")
    assert calls == [1]
