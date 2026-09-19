"""One startup profile owns config, template data, logging and IPC selection."""

from __future__ import annotations

import os
from concurrent.futures import ThreadPoolExecutor
from dataclasses import FrozenInstanceError
from pathlib import Path

import pytest

from docwen_runtime import profile_paths

pytestmark = pytest.mark.contract


@pytest.fixture
def platform_roots(tmp_path, monkeypatch):
    roots = {name: tmp_path / name for name in ("data", "config", "log")}
    for name, root in roots.items():
        monkeypatch.setattr(
            profile_paths.platformdirs, f"user_{name}_dir", lambda *args, root=root, **kwargs: str(root)
        )
    monkeypatch.setattr(profile_paths.sys, "frozen", False, raising=False)
    return roots


def test_data_selects_the_whole_profile_without_changing_environment_or_disk(tmp_path, platform_roots):
    root = tmp_path / "selected"
    environment = {"DOCWEN_DATA_DIR": str(root)}
    before = dict(environment)
    profile = profile_paths.resolve_profile_paths(environment)
    assert (profile.root, profile.data_dir, profile.config_dir, profile.log_dir) == (
        root,
        root,
        root / "configs",
        root / "logs",
    )
    assert profile.source == "explicit" and profile.log_override is None
    assert environment == before and not root.exists()
    with pytest.raises(FrozenInstanceError):
        profile.data_dir = tmp_path / "other"  # pyright: ignore[reportAttributeAccessIssue]


def test_component_overrides_do_not_relocate_other_components(tmp_path, platform_roots):
    config = tmp_path / "only-config"
    profile = profile_paths.resolve_profile_paths({"DOCWEN_CONFIG_DIR": str(config)})
    assert profile.config_dir == config
    assert profile.data_dir == platform_roots["data"]
    assert profile.log_dir == platform_roots["log"]
    assert profile.config_override and profile.source == "platform"
    root, log_root = tmp_path / "whole", tmp_path / "log-override"
    profile = profile_paths.resolve_profile_paths(
        {
            "DOCWEN_DATA_DIR": str(root),
            "DOCWEN_CONFIG_DIR": str(config),
            "DOCWEN_LOG_DIR": str(log_root),
            "DOCWEN_LOG_TO_TEMP": "1",
        }
    )
    assert profile.config_dir == config and profile.data_dir == root
    assert profile.log_dir == log_root / "logs" and profile.log_override == "DOCWEN_LOG_DIR"


def test_source_platform_defaults_remain_explicit_and_separate(platform_roots):
    profile = profile_paths.resolve_profile_paths({})
    assert profile.source == "platform" and not profile.portable
    assert profile.data_dir == platform_roots["data"]
    assert profile.config_dir == platform_roots["config"] / "configs"
    assert profile.log_dir == platform_roots["log"]


@pytest.mark.parametrize("packaged", [False, True])
def test_archive_copy_and_store_upgrade_resolve_the_correct_profile(tmp_path, monkeypatch, platform_roots, packaged):
    monkeypatch.setattr(profile_paths.sys, "frozen", True, raising=False)
    monkeypatch.setattr(profile_paths, "_windows_has_package_identity", lambda: packaged)
    observed = []
    for version in ("version-one", "version-two"):
        executable = tmp_path / version / "DocWenCLI.exe"
        monkeypatch.setattr(profile_paths.sys, "executable", str(executable))
        profile = profile_paths.resolve_profile_paths({})
        expected = platform_roots["data"] if packaged else executable.parent / "data"
        assert profile.data_dir == expected and profile.config_dir == expected / "configs"
        assert profile.log_dir == expected / "logs"
        assert profile.source == ("package" if packaged else "portable")
        assert not executable.parent.exists()
        observed.append(profile.data_dir)
    assert (observed[0] == observed[1]) is packaged


def test_unknown_windows_package_probe_cannot_select_a_portable_root(monkeypatch):
    monkeypatch.setattr(profile_paths.sys, "platform", "win32")
    monkeypatch.delattr(profile_paths.ctypes, "windll", raising=False)
    assert profile_paths._windows_has_package_identity() is True


def test_bound_profile_is_shared_by_threads_and_survives_environment_and_cwd_changes(tmp_path, monkeypatch):
    from docwen_runtime.config.loader import _default_user_config_dir
    from docwen_runtime.logging import log_directory_override_source, resolve_log_file_path
    from docwen_runtime.templates.state import template_data_root

    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("DOCWEN_DATA_DIR", "selected")
    monkeypatch.delenv("DOCWEN_CONFIG_DIR", raising=False)
    monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
    monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)
    startup_env = dict(os.environ)
    root = tmp_path / "selected"
    with profile_paths.bind_process_profile() as profile:
        name = profile_paths.profile_instance_name()
        assert dict(os.environ) == startup_env
        monkeypatch.setenv("DOCWEN_DATA_DIR", str(tmp_path / "other-data"))
        monkeypatch.setenv("DOCWEN_CONFIG_DIR", str(tmp_path / "other-config"))
        monkeypatch.setenv("DOCWEN_LOG_DIR", str(tmp_path / "other-log"))
        monkeypatch.chdir(tmp_path.parent)
        with ThreadPoolExecutor(max_workers=1) as executor:
            assert executor.submit(profile_paths.current_profile_paths).result() is profile
        with profile_paths.bind_process_profile() as nested:
            assert nested is profile
        assert _default_user_config_dir() == root / "configs"
        assert template_data_root() == root
        assert Path(resolve_log_file_path({})).parent == root / "logs"
        assert log_directory_override_source() is None
        assert profile_paths.profile_instance_name() == name
        child_env = profile_paths.profile_process_environment()
        assert child_env["DOCWEN_DATA_DIR"] == str(root)
        assert "DOCWEN_CONFIG_DIR" not in child_env and "DOCWEN_LOG_DIR" not in child_env
        child_env["DOCWEN_DATA_DIR"] = "cannot-mutate-binding"
        assert profile_paths.profile_process_environment()["DOCWEN_DATA_DIR"] == str(root)
    assert profile_paths.current_profile_paths().data_dir == tmp_path / "other-data"


def test_ipc_identity_separates_configuration_overrides_but_not_log_preferences(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCWEN_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("DOCWEN_CONFIG_DIR", str(tmp_path / "config-a"))
    first = profile_paths.profile_instance_name()
    monkeypatch.setenv("DOCWEN_LOG_DIR", str(tmp_path / "log"))
    assert profile_paths.profile_instance_name() == first
    monkeypatch.setenv("DOCWEN_CONFIG_DIR", str(tmp_path / "config-b"))
    assert profile_paths.profile_instance_name() != first
    monkeypatch.setenv("DOCWEN_CONFIG_DIR", str(tmp_path / "other" / ".." / "config-a"))
    assert profile_paths.profile_instance_name() == first


def test_invalid_selected_configuration_directory_does_not_fall_back(tmp_path, monkeypatch, platform_roots):
    from docwen_runtime.config import ConfigLoader

    root = tmp_path / "selected"
    root.mkdir()
    (root / "configs").write_bytes(b"not a directory")
    monkeypatch.setenv("DOCWEN_DATA_DIR", str(root))
    monkeypatch.delenv("DOCWEN_CONFIG_DIR", raising=False)
    with pytest.raises(OSError), profile_paths.bind_process_profile():
        ConfigLoader()
    assert (root / "configs").read_bytes() == b"not a directory"
    assert not platform_roots["config"].exists()


def test_unwritable_selected_profile_does_not_write_an_alternative(tmp_path, monkeypatch, platform_roots):
    from docwen_runtime.config import ConfigLoader
    from docwen_runtime.templates import TemplateManager

    root = tmp_path / "selected"
    root.mkdir()
    monkeypatch.setenv("DOCWEN_DATA_DIR", str(root))
    monkeypatch.delenv("DOCWEN_CONFIG_DIR", raising=False)
    original_mkdir = Path.mkdir

    def denied(path, *args, **kwargs):
        if path == root or root in path.parents:
            raise PermissionError("selected profile is read-only")
        return original_mkdir(path, *args, **kwargs)

    with profile_paths.bind_process_profile():
        config = ConfigLoader(runtime_overrides={"logger": {"enable": False, "console_enable": False}})
        before = {path.name: path.read_bytes() for path in root.iterdir() if path.is_file()}
        monkeypatch.setattr(Path, "mkdir", denied)
        assert not config.set_value("gui.language.locale", "en_US")
        with pytest.raises(PermissionError, match="read-only"):
            TemplateManager.default().ensure_user_directory()
    assert {path.name: path.read_bytes() for path in root.iterdir() if path.is_file()} == before
    assert not (root / "configs").exists() and not (root / "templates").exists()
    assert all(not path.exists() for path in platform_roots.values())
