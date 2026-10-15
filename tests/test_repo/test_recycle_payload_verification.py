from __future__ import annotations

import json
import os
from pathlib import Path

import pytest
from tests.support.recycle_scenarios import directory_link, fake_bin
from tools import recycle, recycle_payload, workspace_cleanup

pytestmark = pytest.mark.integration


def _source(tmp_path: Path) -> tuple[Path, Path]:
    workspace = tmp_path / ".workspace"
    (workspace / "diagnostics").mkdir(parents=True)
    source = workspace / "temp" / "scratch"
    source.mkdir(parents=True)
    (source / "file.txt").write_bytes(b"AAAA")
    return workspace, source


def _saved(workspace: Path, source: Path) -> Path:
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=[source], reason="controlled recycle fixture", disposition="recycle"
    )
    return workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "recycle.json")


def test_payload_identity_is_stable_after_directory_rename_and_keeps_old_lease_bytes(tmp_path: Path) -> None:
    _, source = _source(tmp_path)
    marker = source / workspace_cleanup.LEASE_NAME
    original = json.dumps({"root": str(source), "owner": "original.fixture"}).encode()
    marker.write_bytes(original)
    before = recycle_payload.snapshot_payload(source)
    payload = tmp_path / "relocated"
    source.rename(payload)
    assert recycle_payload.snapshot_payload(payload) == before
    assert (payload / marker.name).read_bytes() == original


def test_absolute_internal_link_is_compared_as_raw_bytes_after_its_old_target_disappears(tmp_path: Path) -> None:
    _, source = _source(tmp_path)
    target = source / "inside"
    target.mkdir()
    (target / "file.txt").write_bytes(b"link payload")
    link = source / "link"
    directory_link(link, target)
    before = recycle_payload.snapshot_payload(source)
    payload = tmp_path / "relocated"
    source.rename(payload)
    assert not (payload / "link").exists()
    assert recycle_payload.snapshot_payload(payload) == before
    assert before["reparsePoints"]


@pytest.mark.parametrize("extension", [".txt", ".cmd", ".bat", ".exe"])
def test_file_reader_handles_windows_lstat_and_fstat_semantic_differences(tmp_path: Path, extension: str) -> None:
    file = tmp_path / ("ctime" + extension)
    file.write_bytes(b"ctime is not a cross-API identity")
    os.utime(file, ns=(1_600_000_000_000_000_000, 1_600_000_000_000_000_000))
    identity = recycle_payload.snapshot_payload(file)
    assert identity["bytes"] == len(file.read_bytes())
    assert identity["rootInode"] == file.lstat().st_ino


def test_payload_content_and_child_file_ids_both_bind_identical_length_replacements(tmp_path: Path) -> None:
    _, source = _source(tmp_path)
    before = recycle_payload.snapshot_payload(source)
    file = source / "file.txt"
    file.write_bytes(b"BBBB")
    changed = recycle_payload.snapshot_payload(source)
    assert changed["bytes"] == before["bytes"] and changed["contentSha256"] != before["contentSha256"]
    file.rename(source / "displaced")
    file.write_bytes(b"AAAA")
    (source / "displaced").rename(tmp_path / "old-file")
    replaced = recycle_payload.snapshot_payload(source)
    assert replaced["contentSha256"] == before["contentSha256"]
    assert replaced["entriesSha256"] != before["entriesSha256"]


def test_payload_snapshot_refuses_root_reparse_without_reading_its_target(tmp_path: Path) -> None:
    _, source = _source(tmp_path)
    link = tmp_path / "root-link"
    try:
        link.symlink_to(source, target_is_directory=True)
    except OSError as error:
        pytest.skip(f"symlink creation unavailable: {error}")
    with pytest.raises(OSError, match="plain_file_or_directory_required"):
        recycle_payload.snapshot_payload(link)


def test_recovery_metadata_and_current_payload_are_verified_in_full(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _, source = _source(tmp_path)
    _, calls = fake_bin(tmp_path, monkeypatch)
    expected = recycle_payload.snapshot_payload(source)
    recycle._perform_recycle(source)
    verified = recycle.verify_recovery_entry(source, calls[0], expected, require_recorded_proof=False)
    assert verified["payloadIdentity"] == expected
    assert verified["metadataIdentity"]["rootInode"] == Path(verified["metadata"]).stat().st_ino
    assert recycle.verify_recovery_entry(source, verified, expected) == verified
    Path(verified["payload"]).joinpath("file.txt").write_bytes(b"BBBB")
    with pytest.raises(OSError, match="identity_or_content_changed"):
        recycle.verify_recovery_entry(source, verified, expected)


@pytest.mark.skipif(os.name != "nt", reason="Windows saved recycle disposition")
def test_post_recycle_mismatch_stays_pending_and_is_never_retried(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workspace, source = _source(tmp_path)
    _, calls = fake_bin(tmp_path, monkeypatch, after_move=lambda payload: (payload / "file.txt").write_bytes(b"BBBB"))
    saved = _saved(workspace, source)
    with pytest.raises(recycle.RecycleUnconfirmedError, match="verification_unconfirmed"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    progress = json.loads(saved.with_suffix(".progress.json").read_text())
    assert not progress["completed"] and progress["pending"]["state"] == "outcome-unconfirmed"
    assert progress["pending"]["recycle"]["sourceAbsent"] is True
    assert progress["pending"]["recycle"]["candidates"] == calls
    assert progress["pending"]["recycle"]["verification"] == "unconfirmed"
    with pytest.raises(ValueError, match="previous_apply_outcome_unconfirmed"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert len(calls) == 1


@pytest.mark.parametrize(
    "change", ["missing", "same-bytes-root-replacement", "same-bytes-child-replacement", "content", "recreated-link"]
)
@pytest.mark.skipif(os.name != "nt", reason="Windows saved recycle disposition")
def test_completed_replay_refuses_changed_recovery_without_changing_original_success(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, change: str
) -> None:
    workspace, source = _source(tmp_path)
    fake_bin(tmp_path, monkeypatch)
    saved = _saved(workspace, source)
    first = workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace) == first
    journal = saved.with_suffix(".progress.json")
    original_journal = journal.read_bytes()
    payload = Path(first["removedEntries"][0]["recovery"][0]["payload"])
    if change == "recreated-link":
        target = tmp_path / "link-target"
        target.mkdir()
        directory_link(source, target)
        target.rename(tmp_path / "removed-link-target")
    elif change in {"missing", "same-bytes-root-replacement"}:
        payload.rename(tmp_path / "displaced-payload")
        if change == "same-bytes-root-replacement":
            payload.mkdir()
            (payload / "file.txt").write_bytes(b"AAAA")
    elif change == "same-bytes-child-replacement":
        (payload / "file.txt").rename(tmp_path / "displaced-file")
        (payload / "file.txt").write_bytes(b"AAAA")
    else:
        (payload / "file.txt").write_bytes(b"BBBB")
    with pytest.raises(
        ValueError, match="completed_target_changed" if change == "recreated-link" else "completed_recovery_unverified"
    ):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert journal.read_bytes() == original_journal
    assert (
        json.loads(original_journal)["completed"][0]["recovery"][0]["verification"] == "identity-and-content-confirmed"
    )


@pytest.mark.skipif(os.name != "nt", reason="Windows saved recycle disposition")
def test_old_recycle_plan_requires_new_payload_proof_before_any_mutation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workspace, source = _source(tmp_path)
    _, calls = fake_bin(tmp_path, monkeypatch)
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=[source], reason="old proof fixture", disposition="recycle"
    )
    del plan["entries"][0]["identity"]["recyclePayload"]
    plan["planFingerprint"] = workspace_cleanup._plan_fingerprint(plan)
    saved = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "old.json")
    with pytest.raises(ValueError, match="missing_payload_proof_regenerate"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert source.exists() and not calls and not saved.with_suffix(".progress.json").exists()


@pytest.mark.skipif(os.name != "nt", reason="Windows recycle observation")
def test_shell_error_survives_failed_recovery_and_unknown_source_observations(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _, source = _source(tmp_path)
    original_exists = recycle._exists
    started = False

    def perform(path: Path) -> None:
        nonlocal started
        started = True
        raise OSError("original shell error")

    def recovered(path: Path):
        if started:
            raise PermissionError("candidate observation denied")
        return []

    def exists(path: Path):
        if started:
            raise PermissionError("source observation denied")
        return original_exists(path)

    monkeypatch.setattr(recycle, "_perform_recycle", perform)
    monkeypatch.setattr(recycle, "_recovery_entries", recovered)
    monkeypatch.setattr(recycle, "_exists", exists)
    with pytest.raises(recycle.RecycleUnconfirmedError, match="original shell error") as caught:
        recycle.recycle_path(source)
    assert caught.value.source_absent is None
    assert len(caught.value.observation_errors) == 2
    assert isinstance(caught.value.__cause__, OSError) and str(caught.value.__cause__) == "original shell error"


@pytest.mark.skipif(os.name != "nt", reason="Windows recycle observation")
def test_unknown_source_after_success_is_unconfirmed_with_exact_candidates(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _, source = _source(tmp_path)
    _, calls = fake_bin(tmp_path, monkeypatch)
    original_exists = recycle._exists

    def exists(path: Path):
        if calls and path == source:
            raise PermissionError("source observation denied")
        return original_exists(path)

    monkeypatch.setattr(recycle, "_exists", exists)
    with pytest.raises(recycle.RecycleUnconfirmedError) as caught:
        recycle.recycle_path(source)
    assert caught.value.source_absent is None and caught.value.recovery == calls
    assert caught.value.observation_errors


@pytest.mark.skipif(os.name != "nt", reason="Windows recycle entrypoint")
def test_direct_recycle_refuses_ancestor_junction_before_resolve_or_shell(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _, source = _source(tmp_path)
    ancestor = tmp_path / "ancestor-link"
    directory_link(ancestor, source.parent)
    _, calls = fake_bin(tmp_path, monkeypatch)
    with pytest.raises(OSError, match="linked_boundary"):
        recycle.recycle_path(ancestor / source.name)
    assert not calls and (source / "file.txt").read_bytes() == b"AAAA"


@pytest.mark.skipif(os.name != "nt", reason="Windows native recycle")
def test_real_recycle_preserves_absolute_internal_junction_and_original_lease_bytes(tmp_path: Path) -> None:
    target = tmp_path / "absolute-link-proof"
    inner = target / "inside"
    inner.mkdir(parents=True)
    (inner / "sample.txt").write_bytes(b"controlled link payload")
    directory_link(target / "link", inner)
    lease = target / ".docwen-temp-lease.json"
    raw = json.dumps(
        {
            "schemaVersion": 1,
            "owner": "docwen.fixture",
            "pid": 999999999,
            "state": "completed-success",
            "root": str(target),
        }
    ).encode()
    lease.write_bytes(raw)
    entries = recycle.recycle_path(target)
    payload = Path(entries[0]["payload"])
    assert not target.exists() and not (payload / "link").exists()
    assert entries[0]["payloadIdentity"]["reparsePoints"]
    assert (payload / lease.name).read_bytes() == raw
    assert recycle.verify_recovery_entry(target, entries[0], entries[0]["payloadIdentity"]) == entries[0]
