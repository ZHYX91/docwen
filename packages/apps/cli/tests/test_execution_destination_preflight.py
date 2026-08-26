from __future__ import annotations

import argparse
import os
import tempfile
from pathlib import Path

import pytest

from docwen_cli.commands import execution_v3

pytestmark = pytest.mark.integration


def _args(
    *,
    output_path: Path | None = None,
    output_dir: Path | None = None,
    overwrite: bool = True,
) -> argparse.Namespace:
    return argparse.Namespace(
        output_path=str(output_path) if output_path is not None else None,
        output_dir=str(output_dir) if output_dir is not None else None,
        files=[],
        overwrite=overwrite,
        command_path="number markdown",
    )


def test_existing_output_directory_uses_real_write_probe_when_access_hint_is_wrong(
    tmp_path: Path,
    monkeypatch,
) -> None:
    output_dir = tmp_path / "acl-denied"
    output_dir.mkdir()
    observed: list[Path] = []

    monkeypatch.setattr(os, "access", lambda *_args, **_kwargs: True)

    def deny_probe(path, *_args, **_kwargs):
        observed.append(Path(path))
        raise PermissionError(13, "Access is denied", str(output_dir))

    monkeypatch.setattr(os, "open", deny_probe)

    result = execution_v3._preflight_destination(_args(output_dir=output_dir))

    assert len(observed) == 1
    assert observed[0].parent == execution_v3.filesystem_path(output_dir, force_extended=True)
    assert observed[0].name.startswith(".__docwen-write-probe-")
    assert result == (
        "invalid_input",
        f"Output directory is not writable: {output_dir.resolve()}",
        {"path": str(output_dir.resolve())},
    )
    assert list(output_dir.iterdir()) == []


def test_destination_write_probe_leaves_no_artifact_on_success(tmp_path: Path) -> None:
    output_dir = tmp_path / "writable"
    output_dir.mkdir()

    assert execution_v3._preflight_destination(_args(output_dir=output_dir)) is None
    assert list(output_dir.iterdir()) == []


def test_missing_output_directory_probes_existing_parent_without_creating_destination(
    tmp_path: Path,
    monkeypatch,
) -> None:
    output_dir = tmp_path / "new-output"
    observed: list[Path] = []
    original = os.open

    def observe_probe(path, *args, **kwargs):
        observed.append(Path(path))
        return original(path, *args, **kwargs)

    monkeypatch.setattr(os, "open", observe_probe)

    assert execution_v3._preflight_destination(_args(output_dir=output_dir)) is None
    assert len(observed) == 1
    assert observed[0].parent == execution_v3.filesystem_path(tmp_path, force_extended=True)
    assert not output_dir.exists()
    assert list(tmp_path.iterdir()) == []


def test_output_path_uses_real_parent_probe_and_fails_without_artifacts(
    tmp_path: Path,
    monkeypatch,
) -> None:
    output_path = tmp_path / "result.docx"

    def deny_probe(*_args, **_kwargs):
        raise PermissionError(13, "Access is denied", str(tmp_path))

    monkeypatch.setattr(os, "open", deny_probe)

    assert execution_v3._preflight_destination(_args(output_path=output_path)) == (
        "invalid_input",
        f"Output parent directory is not writable: {tmp_path.resolve()}",
        {"path": str(tmp_path.resolve())},
    )
    assert list(tmp_path.iterdir()) == []


def test_existing_output_conflict_wins_before_write_probe(tmp_path: Path, monkeypatch) -> None:
    output_path = tmp_path / "result.docx"
    output_path.write_bytes(b"existing")

    def unexpected_probe(*_args, **_kwargs):
        raise AssertionError("write probe must not run after an existing-output conflict")

    monkeypatch.setattr(os, "open", unexpected_probe)

    assert execution_v3._preflight_destination(_args(output_path=output_path, overwrite=False)) == (
        "output_exists",
        f"Output target already exists: {output_path.resolve()}",
        {"path": str(output_path.resolve())},
    )
    assert output_path.read_bytes() == b"existing"


def test_probe_close_failure_fails_closed(tmp_path: Path, monkeypatch) -> None:
    output_dir = tmp_path / "writable"
    output_dir.mkdir()

    monkeypatch.setattr(os, "open", lambda *_args, **_kwargs: 12345)
    monkeypatch.setattr(os, "write", lambda *_args, **_kwargs: 1)
    monkeypatch.setattr(
        os,
        "close",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            PermissionError(13, "Access is denied while closing", str(output_dir))
        ),
    )
    monkeypatch.setattr(os, "unlink", lambda *_args, **_kwargs: None)

    assert execution_v3._preflight_destination(_args(output_dir=output_dir)) == (
        "invalid_input",
        f"Output directory is not writable: {output_dir.resolve()}",
        {"path": str(output_dir.resolve())},
    )
    assert list(output_dir.iterdir()) == []


def test_probe_cleanup_failure_fails_closed(tmp_path: Path, monkeypatch) -> None:
    output_dir = tmp_path / "writable"
    output_dir.mkdir()

    monkeypatch.setattr(os, "open", lambda *_args, **_kwargs: 12345)
    monkeypatch.setattr(os, "write", lambda *_args, **_kwargs: 1)
    monkeypatch.setattr(os, "close", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        os,
        "unlink",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            PermissionError(13, "Access is denied while deleting", str(output_dir))
        ),
    )

    assert execution_v3._preflight_destination(_args(output_dir=output_dir)) == (
        "invalid_input",
        f"Output directory is not writable: {output_dir.resolve()}",
        {"path": str(output_dir.resolve())},
    )
    assert list(output_dir.iterdir()) == []


def test_write_probe_avoids_temporaryfile_on_windows_acl_boundaries(
    tmp_path: Path,
    monkeypatch,
) -> None:
    """TemporaryFile can block indefinitely under an explicit Windows deny ACL."""

    output_dir = tmp_path / "writable"
    output_dir.mkdir()

    def unexpected_temporary_file(*_args, **_kwargs):
        raise AssertionError("destination probes must not use tempfile.TemporaryFile")

    monkeypatch.setattr(tempfile, "TemporaryFile", unexpected_temporary_file)

    assert execution_v3._preflight_destination(_args(output_dir=output_dir)) is None
    assert list(output_dir.iterdir()) == []
