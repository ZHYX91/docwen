from __future__ import annotations

import os
import subprocess
from datetime import UTC, datetime
from pathlib import Path
from types import SimpleNamespace
from typing import cast

import pytest
from tests.test_repo.test_workspace_housekeeping_plan import _lease, _workspace
from tools import workspace_cleanup

pytestmark = [pytest.mark.unit, pytest.mark.windows_only, pytest.mark.skipif(os.name != "nt", reason="Windows links")]


def _link(link: Path, target: Path, kind: str = "junction") -> None:
    if kind != "junction":
        try:
            os.symlink(target, link, target_is_directory=kind == "directory-symlink")
        except OSError as error:
            pytest.skip(f"Real Windows symlink privilege unavailable: {error}")
        return
    completed = subprocess.run(
        ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target)],
        capture_output=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr or completed.stdout


def _owned(tmp_path: Path) -> tuple[Path, Path]:
    _, workspace = _workspace(tmp_path)
    root = workspace / "temp" / "owned"
    _lease(root, created_at=datetime.now(UTC), state="retained-failure")
    return workspace, root


@pytest.mark.parametrize("kind", ["junction", "directory-symlink", "file-symlink"])
def test_native_delete_removes_only_owned_dangling_link_not_moved_target(tmp_path: Path, kind: str) -> None:
    workspace, root = _owned(tmp_path)
    target = root / "target"
    if kind == "file-symlink":
        target.write_bytes(b"preserve target")
    else:
        target.mkdir()
        (target / "content.txt").write_bytes(b"preserve target")
    link = root / "link"
    _link(link, target, kind)
    retained = tmp_path / "moved-target"
    target.rename(retained)
    before = link.lstat()

    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=(root,), reason="superseded owned failure", disposition="delete"
    )
    point = plan["entries"][0]["identity"]["reparsePoints"][0]
    assert point["dangling"] is True
    assert point["tag"] == before.st_reparse_tag
    assert point["rawTarget"] == os.readlink(link)
    saved = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "dangling-plan.json")
    result = workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)

    assert result["removed"] == [str(root)]
    assert not root.exists()
    preserved = retained if kind == "file-symlink" else retained / "content.txt"
    assert preserved.read_bytes() == b"preserve target"


def test_dangling_link_target_change_invalidates_saved_plan(tmp_path: Path) -> None:
    workspace, root = _owned(tmp_path)
    first = root / "first"
    first.mkdir()
    link = root / "link"
    _link(link, first)
    first.rmdir()
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=(root,), reason="superseded failure", disposition="delete"
    )
    saved = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "changed.json")
    root_time = root.stat().st_mtime_ns
    link.rmdir()
    second = root / "second"
    second.mkdir()
    _link(link, second)
    second.rmdir()
    os.utime(root, ns=(root_time, root_time))

    with pytest.raises(workspace_cleanup.HousekeepingError, match="target_identity_changed"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert root.exists() and link.lstat().st_reparse_tag == 0xA0000003


def test_reappearing_dangling_target_invalidates_plan_and_preserves_new_content(tmp_path: Path) -> None:
    workspace, root = _owned(tmp_path)
    target = root / "target"
    target.mkdir()
    link = root / "link"
    _link(link, target)
    target.rmdir()
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=(root,), reason="superseded failure", disposition="delete"
    )
    saved = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "reappeared.json")
    target.mkdir()
    (target / "new.txt").write_text("new content", encoding="utf-8")

    with pytest.raises(workspace_cleanup.HousekeepingError, match="target_identity_changed"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert (target / "new.txt").read_text(encoding="utf-8") == "new content"


def test_dangling_external_junction_is_still_rejected(tmp_path: Path) -> None:
    workspace, root = _owned(tmp_path)
    target = tmp_path / "external"
    target.mkdir()
    link = root / "escape"
    _link(link, target)
    retained = tmp_path / "retained"
    target.rename(retained)
    (retained / "keep.txt").write_text("keep", encoding="utf-8")

    with pytest.raises(workspace_cleanup.HousekeepingError, match="unresolved_reparse_target"):
        workspace_cleanup.create_plan(workspace_root=workspace, explicit_targets=(root,), reason="external refused")
    assert (retained / "keep.txt").read_text(encoding="utf-8") == "keep"


def test_dangling_target_through_another_junction_is_still_rejected(tmp_path: Path) -> None:
    workspace, root = _owned(tmp_path)
    outside = tmp_path / "outside"
    outside.mkdir()
    _link(root / "alias", outside)
    # The internal lexical name must not hide an external or linked ancestor.
    _link(root / "link", root / "alias" / "missing")
    with pytest.raises(workspace_cleanup.HousekeepingError):
        workspace_cleanup.create_plan(workspace_root=workspace, explicit_targets=(root,), reason="chain refused")
    assert outside.is_dir()


@pytest.mark.parametrize("ownership", ["unowned", "active", "live-terminal"])
def test_dangling_admission_requires_dead_terminal_owner(tmp_path: Path, ownership: str) -> None:
    _, workspace = _workspace(tmp_path)
    root = workspace / "temp" / "held"
    if ownership == "unowned":
        root.mkdir()
    else:
        _lease(
            root,
            created_at=datetime.now(UTC),
            state="active" if ownership == "active" else "retained-failure",
            pid=999_999_999 if ownership == "active" else os.getpid(),
        )
    target = root / "missing"
    target.mkdir()
    _link(root / "link", target)
    target.rmdir()
    with pytest.raises(workspace_cleanup.HousekeepingError):
        workspace_cleanup.create_plan(workspace_root=workspace, explicit_targets=(root,), reason="held refused")
    assert root.exists()


def test_unknown_reparse_tag_is_still_rejected(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workspace, root = _owned(tmp_path)
    target = root / "missing"
    target.mkdir()
    link = root / "link"
    _link(link, target)
    target.rmdir()
    original = os.lstat

    def unknown(path: str | os.PathLike[str], *, dir_fd: int | None = None) -> os.stat_result:
        result = original(path, dir_fd=dir_fd)
        if workspace_cleanup._logical_path(path) == link:
            fields = {name: getattr(result, name) for name in dir(result) if name.startswith("st_")}
            fields["st_reparse_tag"] = 0x80000000
            return cast(os.stat_result, SimpleNamespace(**fields))
        return result

    monkeypatch.setattr(os, "lstat", unknown)
    monkeypatch.setattr(Path, "lstat", lambda self: unknown(self))
    with pytest.raises(workspace_cleanup.HousekeepingError, match="unsupported_reparse_tag"):
        workspace_cleanup.create_plan(workspace_root=workspace, explicit_targets=(root,), reason="unknown refused")
