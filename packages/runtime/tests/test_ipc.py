"""Tests for the runtime single-instance primitive."""

from __future__ import annotations

import os
import stat
import subprocess
import sys
from collections.abc import Generator
from pathlib import Path
from typing import cast

import pytest

from docwen_runtime.ipc import SingleInstance, disable_ipc, enable_ipc, is_ipc_disabled
from docwen_runtime.ipc.single_instance import SingleInstanceError, create_single_instance

pytestmark = pytest.mark.integration


def _unix_uid() -> int:
    getuid = vars(os).get("getuid")
    assert callable(getuid)
    return cast(int, getuid())


@pytest.fixture(autouse=True)
def _reset_ipc_toggle(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> Generator[None, None, None]:
    if sys.platform != "win32":
        runtime_root = tmp_path / "xdg-runtime"
        runtime_root.mkdir(mode=0o700)
        runtime_root.chmod(0o700)
        monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
    enable_ipc()
    yield
    enable_ipc()


class TestIpcToggle:
    def test_default_enabled(self) -> None:
        assert not is_ipc_disabled()

    def test_disable_and_enable(self) -> None:
        disable_ipc()
        assert is_ipc_disabled()
        enable_ipc()
        assert not is_ipc_disabled()

    def test_disable_makes_acquire_succeed(self) -> None:
        disable_ipc()
        instance = SingleInstance("should_not_matter")
        assert instance.acquire() is True
        assert instance.is_acquired is True
        instance.release()
        assert instance.is_acquired is False


class TestSingleInstance:
    def test_construct_with_empty_name_raises(self) -> None:
        with pytest.raises(ValueError, match="app_name"):
            SingleInstance("")

    def test_construct_with_whitespace_name_raises(self) -> None:
        with pytest.raises(ValueError, match="app_name"):
            SingleInstance("   ")

    def test_acquire_succeeds_when_no_contention(self) -> None:
        instance = create_single_instance("test_no_contention")
        try:
            assert instance.acquire() is True
            assert instance.is_acquired is True
        finally:
            instance.release()

    def test_second_acquire_fails(self) -> None:
        first = SingleInstance("test_double_lock")
        second = SingleInstance("test_double_lock")
        try:
            assert first.acquire() is True
            assert second.acquire() is False
            assert second.is_acquired is False
        finally:
            first.release()
            second.close()

    def test_release_allows_reacquire(self) -> None:
        first = SingleInstance("test_reacquire")
        assert first.acquire() is True
        first.release()
        second = SingleInstance("test_reacquire")
        try:
            assert second.acquire() is True
        finally:
            second.release()

    def test_context_manager_acquire(self) -> None:
        instance = SingleInstance("test_context")
        with instance as acquired:
            assert acquired is True
            assert instance.is_acquired is True
        assert instance.is_acquired is False

    def test_context_manager_contention(self) -> None:
        first = SingleInstance("test_ctx_contend")
        with first as acquired_first:
            assert acquired_first is True
            second = SingleInstance("test_ctx_contend")
            with second as acquired_second:
                assert acquired_second is False
            second.close()

    def test_ipc_dir_property(self) -> None:
        instance = SingleInstance("test_ipc_dir")
        try:
            if sys.platform == "win32":
                assert "test_ipc_dir" in instance.ipc_dir
            else:
                assert Path(instance.ipc_dir).name.startswith("docwen-instance-")
            assert os.path.isdir(instance.ipc_dir)
        finally:
            instance.close()

    def test_lock_path_property(self) -> None:
        instance = SingleInstance("test_lock_path")
        try:
            assert instance.lock_path.endswith("instance.lock")
            assert instance.ipc_dir in instance.lock_path
        finally:
            instance.close()

    def test_release_idempotent(self) -> None:
        instance = SingleInstance("test_idempotent")
        instance.acquire()
        instance.release()
        instance.release()

    def test_release_without_acquire_is_safe(self) -> None:
        SingleInstance("test_no_acquire").release()

    def test_create_single_instance_factory(self) -> None:
        assert isinstance(create_single_instance("test_factory"), SingleInstance)


@pytest.mark.skipif(sys.platform == "win32", reason="Unix descriptor and permission contract")
class TestUnixSingleInstanceSecurity:
    @staticmethod
    def _use_private_xdg(monkeypatch: pytest.MonkeyPatch, root: Path) -> None:
        root.mkdir(mode=0o700, exist_ok=True)
        root.chmod(0o700)
        monkeypatch.setenv("XDG_RUNTIME_DIR", str(root))

    def test_private_xdg_namespace_and_lock_modes(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        runtime_root = tmp_path / "runtime"
        self._use_private_xdg(monkeypatch, runtime_root)
        instance = SingleInstance("../../unsafe/name")
        lock_path = Path(instance.lock_path)
        try:
            assert Path(instance.ipc_dir).parent == runtime_root
            assert ".." not in Path(instance.ipc_dir).name
            assert stat.S_IMODE(Path(instance.ipc_dir).stat().st_mode) == 0o700
            assert instance.acquire() is True
            lock_metadata = lock_path.stat(follow_symlinks=False)
            assert stat.S_ISREG(lock_metadata.st_mode)
            assert stat.S_IMODE(lock_metadata.st_mode) == 0o600
            assert lock_metadata.st_uid == _unix_uid()
            assert lock_path.read_text(encoding="utf-8") == str(os.getpid())
        finally:
            instance.release()
        assert lock_path.is_file()
        assert lock_path.read_text(encoding="utf-8") == ""

    def test_relative_xdg_runtime_dir_fails_closed(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("XDG_RUNTIME_DIR", "relative/runtime")
        with pytest.raises(SingleInstanceError, match="unix_ipc_root_not_absolute"):
            SingleInstance("relative-xdg")

    def test_xdg_runtime_symlink_fails_closed(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        target = tmp_path / "target"
        target.mkdir(mode=0o700)
        target.chmod(0o700)
        link = tmp_path / "runtime-link"
        link.symlink_to(target, target_is_directory=True)
        monkeypatch.setenv("XDG_RUNTIME_DIR", str(link))
        with pytest.raises(SingleInstanceError, match="unix_ipc_root_link_forbidden"):
            SingleInstance("linked-xdg")

    def test_xdg_runtime_permissions_fail_closed(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        runtime_root = tmp_path / "runtime"
        runtime_root.mkdir(mode=0o755)
        runtime_root.chmod(0o755)
        monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
        with pytest.raises(SingleInstanceError, match="unix_ipc_root_permissions_unsafe"):
            SingleInstance("permissive-xdg")

    def test_existing_namespace_permissions_fail_closed(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        runtime_root = tmp_path / "runtime"
        self._use_private_xdg(monkeypatch, runtime_root)
        first = SingleInstance("unsafe-namespace")
        namespace = Path(first.ipc_dir)
        namespace.chmod(0o755)
        with pytest.raises(SingleInstanceError, match="unix_ipc_directory_permissions_unsafe"):
            SingleInstance("unsafe-namespace")

    def test_lock_symlink_is_rejected_without_touching_target(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        runtime_root = tmp_path / "runtime"
        self._use_private_xdg(monkeypatch, runtime_root)
        instance = SingleInstance("symlink-lock")
        target = tmp_path / "target.txt"
        target.write_text("preserve", encoding="utf-8")
        Path(instance.lock_path).symlink_to(target)
        with pytest.raises(SingleInstanceError, match="unix_lock_open_failed"):
            instance.acquire()
        assert target.read_text(encoding="utf-8") == "preserve"

    def test_lock_permissions_are_not_repaired(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        runtime_root = tmp_path / "runtime"
        self._use_private_xdg(monkeypatch, runtime_root)
        instance = SingleInstance("permissive-lock")
        lock_path = Path(instance.lock_path)
        lock_path.write_text("untrusted", encoding="utf-8")
        lock_path.chmod(0o644)
        with pytest.raises(SingleInstanceError, match="unix_lock_permissions_unsafe"):
            instance.acquire()
        assert lock_path.read_text(encoding="utf-8") == "untrusted"
        assert stat.S_IMODE(lock_path.stat().st_mode) == 0o644

    def test_hardlinked_lock_is_rejected(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        runtime_root = tmp_path / "runtime"
        self._use_private_xdg(monkeypatch, runtime_root)
        instance = SingleInstance("hardlinked-lock")
        source = Path(instance.ipc_dir) / "source"
        source.write_text("untrusted", encoding="utf-8")
        source.chmod(0o600)
        os.link(source, instance.lock_path)
        with pytest.raises(SingleInstanceError, match="unix_lock_link_count_unsafe"):
            instance.acquire()

    def test_fallback_namespace_is_per_user(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        fallback_root = tmp_path / "fallback"
        fallback_root.mkdir(mode=0o700)
        fallback_root.chmod(0o700)
        monkeypatch.delenv("XDG_RUNTIME_DIR", raising=False)
        monkeypatch.setattr("docwen_runtime.ipc.single_instance.tempfile.gettempdir", lambda: str(fallback_root))
        instance = SingleInstance("fallback")
        assert Path(instance.ipc_dir).parent == fallback_root
        assert f"-{_unix_uid()}-" in Path(instance.ipc_dir).name

    def test_lock_contends_across_processes_and_reacquires_after_release(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        runtime_root = tmp_path / "runtime"
        self._use_private_xdg(monkeypatch, runtime_root)
        instance = SingleInstance("cross-process")
        try:
            assert instance.acquire() is True
            child_code = (
                "from docwen_runtime.ipc import SingleInstance; "
                "value=SingleInstance('cross-process'); "
                "raise SystemExit(0 if value.acquire() is False else 9)"
            )
            completed = subprocess.run(
                [sys.executable, "-c", child_code],
                check=False,
                capture_output=True,
                text=True,
                timeout=10,
            )
            assert completed.returncode == 0, completed.stderr
        finally:
            instance.release()

        child_code = (
            "from docwen_runtime.ipc import SingleInstance; "
            "value=SingleInstance('cross-process'); "
            "acquired=value.acquire(); value.release(); "
            "raise SystemExit(0 if acquired else 9)"
        )
        completed = subprocess.run(
            [sys.executable, "-c", child_code],
            check=False,
            capture_output=True,
            text=True,
            timeout=10,
        )
        assert completed.returncode == 0, completed.stderr
