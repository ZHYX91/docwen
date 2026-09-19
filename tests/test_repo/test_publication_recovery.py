from __future__ import annotations

import copy
import functools
import http.client
import json
from pathlib import Path

import pytest
from scripts.release import publication_session
from scripts.release.publication_contract import PublicationError, read_object
from scripts.release.publication_http import ApiError, PendingRead, read_with_retry
from scripts.release.publication_session import ReleaseSession
from tests.support.publication import COMMIT, DIGEST, REPOSITORY, VERSION, Clock, FakeGitHub, candidate

pytestmark = pytest.mark.unit


@pytest.fixture
def session(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    directory, _ = candidate(tmp_path)
    clock = Clock()
    monkeypatch.setattr(
        publication_session, "read_with_retry", functools.partial(read_with_retry, clock=clock.time, sleep=clock.sleep)
    )
    monkeypatch.setattr(ReleaseSession, "verify_provenance", lambda self: self.save(provenance="verified"))
    api = FakeGitHub()
    receipt = tmp_path / "progress.json"

    def create():
        return ReleaseSession(
            api,
            directory,
            receipt,
            repository=REPOSITORY,
            version=VERSION,
            commit=COMMIT,
            artifact_id=20,
            artifact_digest=DIGEST,
        )

    return api, create, receipt, clock


@pytest.mark.parametrize("lost", ["create", "upload", "publish"])
def test_lost_write_response_is_reconciled_without_repeating_the_write(session, lost: str) -> None:
    api, create, receipt, clock = session
    api.lost.add(lost)
    result = create().publish(notes="Release notes")
    assert result["stage"] == "published-awaiting-readback"
    assert api.downloads == []
    assert len(api.writes) == 6  # one draft, four public assets, one publish
    assert len({path for _, path in api.writes}) == 6
    assert clock.delays[:2] == [2, 4]
    assert read_object(receipt)["releaseId"] == 30
    assert create().verify_published()["stage"] == "verified"
    assert len(api.downloads) == 4
    assert not any(path.endswith("/immutable-releases") for path in api.reads)


def test_published_readback_failure_can_resume_without_any_write(session) -> None:
    api, create, receipt, _ = session
    api.bad_download = True
    assert create().publish(notes="Release notes")["stage"] == "published-awaiting-readback"
    with pytest.raises(PublicationError, match="remote bytes mismatch"):
        create().verify_published()
    assert read_object(receipt)["stage"] == "published-awaiting-readback"
    writes = api.writes[:]
    api.bad_download = False
    assert create().publish(notes="Release notes")["stage"] == "published-awaiting-readback"
    assert create().verify_published()["stage"] == "verified"
    assert api.writes == writes


def test_draft_asset_identity_mismatch_prevents_publication(session, monkeypatch: pytest.MonkeyPatch) -> None:
    """The draft is pinned by the platform size and digest, not by a second byte download."""

    api, create, receipt, _ = session
    original = api.request

    def tampered(method: str, path: str, **kwargs):
        result = original(method, path, **kwargs)
        if method == "GET" and isinstance(result, dict) and result.get("assets"):
            result = copy.deepcopy(result)
            result["assets"][-1]["digest"] = "sha256:" + "0" * 64
        return result

    monkeypatch.setattr(api, "request", tampered)
    with pytest.raises(PublicationError, match="remote asset mismatch"):
        create().publish(notes="Release notes")
    assert api.release["draft"] is True
    assert not any(method == "PATCH" for method, _ in api.writes)
    assert read_object(receipt)["stage"] != "verified"


def test_process_interruption_after_upload_resumes_the_pending_write(session, monkeypatch: pytest.MonkeyPatch) -> None:
    api, create, receipt, _ = session
    original = api.request

    def interrupted(method, path, **kwargs):
        result = original(method, path, **kwargs)
        if method == "POST" and "/assets?" in path:
            raise KeyboardInterrupt
        return result

    monkeypatch.setattr(api, "request", interrupted)
    with pytest.raises(KeyboardInterrupt):
        create().publish(notes="Release notes")
    assert read_object(receipt)["pending"].startswith("upload:")
    monkeypatch.setattr(api, "request", original)
    assert create().publish(notes="Release notes")["stage"] == "published-awaiting-readback"
    assert len(api.writes) == 6


@pytest.mark.parametrize("failure", ["lint", "run", "digest", "expired"])
def test_missing_or_failed_preflight_prevents_every_write(session, failure: str) -> None:
    api, create, _, _ = session
    if failure == "lint":
        api.jobs["jobs"][0]["conclusion"] = "failure"
    elif failure == "run":
        api.run["head_branch"] = "main"
    elif failure == "expired":
        api.artifact["expired"] = True
    else:
        api.artifact["digest"] = "sha256:" + "f" * 64
    with pytest.raises(PublicationError):
        create().publish(notes="Release notes")
    assert api.writes == []


def test_exact_published_release_is_read_only_until_independent_verification(session) -> None:
    api, create, _, _ = session
    create().publish(notes="Release notes")
    writes = api.writes[:]
    assert create().publish(notes="Release notes")["stage"] == "published-awaiting-readback"
    assert api.writes == writes
    assert api.downloads == []
    assert create().verify_published()["stage"] == "verified"
    assert len(api.downloads) == 4
    assert api.writes == writes


def test_nonimmutable_publication_never_reports_verified_or_downloads_assets(session, monkeypatch) -> None:
    api, create, receipt, _ = session
    original = api.request

    def mutable_release(method, path, **kwargs):
        result = original(method, path, **kwargs)
        if method == "PATCH":
            api.release["immutable"] = False
            result["immutable"] = False
        return result

    monkeypatch.setattr(api, "request", mutable_release)
    with pytest.raises(PublicationError, match="budget exhausted"):
        create().publish(notes="Release notes")
    assert len(api.writes) == 6
    assert api.downloads == []
    assert read_object(receipt)["stage"] == "published-awaiting-readback"


def test_foreign_draft_is_never_modified(session) -> None:
    api, create, _, _ = session
    api.release = {"id": 31, "tag_name": VERSION, "prerelease": False, "draft": True, "body": "different candidate"}
    with pytest.raises(PublicationError, match="another candidate"):
        create().publish(notes="Release notes")
    assert api.writes == []


def test_successful_publish_response_with_wrong_identity_keeps_pending_write(session, monkeypatch) -> None:
    api, create, receipt, _ = session
    original = api.request

    def wrong_response(method, path, **kwargs):
        result = original(method, path, **kwargs)
        return {**result, "id": 999} if method == "PATCH" else result

    monkeypatch.setattr(api, "request", wrong_response)
    with pytest.raises(PublicationError, match="release ID changed"):
        create().publish(notes="Release notes")
    assert read_object(receipt)["pending"] == "publish"
    monkeypatch.setattr(api, "request", original)
    assert create().publish(notes="Release notes")["stage"] == "published-awaiting-readback"
    assert len(api.writes) == 6


@pytest.mark.parametrize(
    "error", [TimeoutError(), json.JSONDecodeError("truncated", "{", 1), http.client.IncompleteRead(b"{")]
)
def test_transient_reads_recover_within_one_run(error: Exception) -> None:
    clock = Clock()
    calls = []

    def operation(timeout: float):
        calls.append(timeout)
        if len(calls) == 1:
            raise error
        return "ok"

    assert read_with_retry(operation, clock=clock.time, sleep=clock.sleep) == "ok"
    assert calls == [60, 58]


def test_retry_after_is_respected_and_request_time_consumes_the_same_budget() -> None:
    clock = Clock()
    calls = []

    def operation(timeout: float):
        calls.append(timeout)
        clock.now += 7
        if len(calls) == 1:
            raise ApiError(429, retry_after=12)
        return "ok"

    assert read_with_retry(operation, clock=clock.time, sleep=clock.sleep) == "ok"
    assert clock.delays == [12]
    assert calls == [60, 41]


@pytest.mark.parametrize("error", [ApiError(401), ApiError(403), PublicationError("hash mismatch")])
def test_permission_and_identity_failures_are_not_retried(error: Exception) -> None:
    clock = Clock()

    def operation(_timeout: float):
        raise error

    with pytest.raises(PublicationError):
        read_with_retry(operation, clock=clock.time, sleep=clock.sleep)
    assert clock.delays == []


def test_request_consuming_the_budget_cannot_start_another_attempt() -> None:
    clock = Clock()

    def operation(timeout: float):
        clock.now += timeout
        raise PendingRead("not yet visible")

    with pytest.raises(PublicationError, match="budget exhausted"):
        read_with_retry(operation, clock=clock.time, sleep=clock.sleep)
    assert clock.now == 60
    assert clock.delays == []
