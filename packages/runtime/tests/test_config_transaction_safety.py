"""Failure-injection contracts for failure-safe configuration persistence."""

from __future__ import annotations

import logging
import os
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _install_fake_file_symlinks(
    monkeypatch: pytest.MonkeyPatch,
    links: dict[Path, Path],
) -> None:
    """Model file symlinks without requiring Windows symlink privileges."""
    managed_paths = set(links)
    original_exists = Path.exists
    original_is_symlink = Path.is_symlink
    original_readlink = Path.readlink
    original_resolve = Path.resolve
    original_symlink_to = Path.symlink_to
    original_unlink = Path.unlink

    def _resolved_target(path: Path, *, strict: bool = False) -> Path:
        raw_target = links[path]
        target = raw_target if raw_target.is_absolute() else path.parent / raw_target
        return original_resolve(target, strict=strict)

    def fake_exists(path: Path) -> bool:
        if path in links:
            return original_exists(_resolved_target(path))
        return original_exists(path)

    def fake_is_symlink(path: Path) -> bool:
        return path in links or original_is_symlink(path)

    def fake_readlink(path: Path) -> Path:
        if path in links:
            return links[path]
        return original_readlink(path)

    def fake_resolve(path: Path, strict: bool = False) -> Path:
        if path in links:
            return _resolved_target(path, strict=strict)
        return original_resolve(path, strict=strict)

    def fake_symlink_to(
        path: Path,
        target: str | os.PathLike[str],
        target_is_directory: bool = False,
    ) -> None:
        if path in managed_paths:
            assert target_is_directory is False
            links[path] = Path(target)
            return
        original_symlink_to(path, target, target_is_directory=target_is_directory)

    def fake_unlink(path: Path, *args, **kwargs) -> None:
        if path in links:
            del links[path]
            return
        original_unlink(path, *args, **kwargs)

    monkeypatch.setattr(Path, "exists", fake_exists)
    monkeypatch.setattr(Path, "is_symlink", fake_is_symlink)
    monkeypatch.setattr(Path, "readlink", fake_readlink)
    monkeypatch.setattr(Path, "resolve", fake_resolve)
    monkeypatch.setattr(Path, "symlink_to", fake_symlink_to)
    monkeypatch.setattr(Path, "unlink", fake_unlink)


def _write_minimal_base_config_tree(base_dir: Path) -> None:
    from docwen_runtime.config.registry import CONFIG_FILES

    for spec in CONFIG_FILES:
        path = base_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


def _loader_with_three_overrides(tmp_path: Path):
    from docwen_runtime.config.loader import ConfigLoader

    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    _write_minimal_base_config_tree(base_dir)

    base_contents = {
        "gui.toml": '[window]\ndefault_mode = "single"\n',
        "output.toml": '[directory]\nmode = "source"\n',
        "link.toml": '[format]\nimage_link_style = "wiki_embed"\n',
    }
    user_contents = {
        "gui.toml": '# gui preimage\n[window]\ndefault_mode = "batch"\n',
        "output.toml": '# output preimage\n[directory]\nmode = "custom"\n',
        "link.toml": '# link preimage\n[format]\nimage_link_style = "markdown_link"\n',
    }
    for rel_path, content in base_contents.items():
        (base_dir / rel_path).write_text(content, encoding="utf-8")
    for rel_path, content in user_contents.items():
        (user_dir / rel_path).write_text(content, encoding="utf-8")

    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    preimages = {rel_path: (user_dir / rel_path).read_bytes() for rel_path in user_contents}
    return loader, user_dir, preimages


def test_atomic_toml_replace_failure_preserves_original_and_cleans_temp(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.toml_io import write_toml_file

    target = tmp_path / "gui.toml"
    original = b'# original\n[window]\ndefault_mode = "single"\n'
    target.write_bytes(original)

    def fail_replace(_source: str | bytes | os.PathLike[str] | os.PathLike[bytes], _target) -> None:
        raise OSError("simulated atomic replace failure")

    monkeypatch.setattr(os, "replace", fail_replace)

    with pytest.raises(OSError, match="simulated atomic replace failure"):
        write_toml_file(target, {"window": {"default_mode": "batch"}})

    assert target.read_bytes() == original
    assert list(tmp_path.glob(f".{target.name}.*.tmp")) == []


def test_atomic_toml_fsync_failure_preserves_original_and_cleans_temp(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io

    target = tmp_path / "gui.toml"
    original = b'[window]\ndefault_mode = "single"\n'
    target.write_bytes(original)

    def fail_fsync(_descriptor: int) -> None:
        raise OSError("simulated fsync failure")

    monkeypatch.setattr(toml_io.os, "fsync", fail_fsync)

    with pytest.raises(OSError, match="simulated fsync failure"):
        toml_io.write_toml_file(target, {"window": {"default_mode": "batch"}})

    assert target.read_bytes() == original
    assert list(tmp_path.glob(f".{target.name}.*.tmp")) == []


def test_atomic_toml_replace_failure_keeps_absent_target_absent(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io

    target = tmp_path / "new.toml"

    def fail_replace(_source, _target) -> None:
        raise OSError("simulated atomic replace failure")

    monkeypatch.setattr(toml_io.os, "replace", fail_replace)

    with pytest.raises(OSError, match="simulated atomic replace failure"):
        toml_io.write_toml_file(target, {"value": "new"})

    assert not target.exists()
    assert list(tmp_path.glob(f".{target.name}.*.tmp")) == []


def test_atomic_toml_stages_in_destination_directory_before_replace(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io

    target = tmp_path / "nested" / "gui.toml"
    replace_calls: list[tuple[Path, Path]] = []
    original_replace = toml_io.os.replace

    def observe_replace(source, destination) -> None:
        source_path = Path(source)
        destination_path = Path(destination)
        replace_calls.append((source_path, destination_path))
        assert source_path.parent == destination_path.parent
        original_replace(source, destination)

    monkeypatch.setattr(toml_io.os, "replace", observe_replace)

    toml_io.write_toml_file(target, {"window": {"default_mode": "batch"}})

    assert len(replace_calls) == 1
    assert replace_calls[0][1] == target
    assert not replace_calls[0][0].exists()
    assert toml_io.read_toml_file(target)["window"]["default_mode"] == "batch"


def test_atomic_toml_write_preserves_symlink_path_and_replaces_resolved_target(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io

    link_path = tmp_path / "linked.toml"
    target_path = tmp_path / "target.toml"
    link_sentinel = b"logical symlink placeholder"
    link_path.write_bytes(link_sentinel)
    target_path.write_text('value = "old"\n', encoding="utf-8")
    original_is_symlink = Path.is_symlink
    original_resolve = Path.resolve

    def fake_is_symlink(path: Path) -> bool:
        return path == link_path or original_is_symlink(path)

    def fake_resolve(path: Path, *args, **kwargs) -> Path:
        if path == link_path:
            return target_path
        return original_resolve(path, *args, **kwargs)

    monkeypatch.setattr(Path, "is_symlink", fake_is_symlink)
    monkeypatch.setattr(Path, "resolve", fake_resolve)

    toml_io.atomic_write_text(link_path, 'value = "new"\n')

    assert link_path.read_bytes() == link_sentinel
    assert target_path.read_text(encoding="utf-8") == 'value = "new"\n'


def test_broken_symlink_write_reload_failure_restores_link_and_missing_target(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    loader, user_dir, _preimages = _loader_with_three_overrides(tmp_path)
    link_path = user_dir / "gui.toml"
    link_path.unlink()
    target_path = tmp_path / "linked" / "gui.toml"
    raw_target = Path("..") / "linked" / "gui.toml"
    links = {link_path: raw_target}
    _install_fake_file_symlinks(monkeypatch, links)

    original_reload = loader.reload
    reload_count = 0

    def fail_once_then_reload() -> None:
        nonlocal reload_count
        reload_count += 1
        if reload_count == 1:
            raise OSError("simulated post-write reload failure")
        original_reload()

    monkeypatch.setattr(loader, "reload", fail_once_then_reload)

    assert loader.set_value("gui.window.default_mode", "batch") is False
    assert reload_count == 2
    assert links == {link_path: raw_target}
    assert not target_path.exists()
    assert loader.config_state_trusted is True
    assert loader.config.gui.window.default_mode == "single"


def test_valid_symlink_reset_reload_failure_restores_link_and_target(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.config import loader as loader_module

    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    link_path = user_dir / "gui.toml"
    target_path = tmp_path / "linked" / "gui.toml"
    target_path.parent.mkdir()
    target_path.write_bytes(preimages["gui.toml"])
    link_path.unlink()
    raw_target = Path("..") / "linked" / "gui.toml"
    links = {link_path: raw_target}
    _install_fake_file_symlinks(monkeypatch, links)

    original_read_toml_file = loader_module.read_toml_file

    def read_fake_link(path: str | Path):
        file_path = Path(path)
        if file_path in links:
            return original_read_toml_file(target_path)
        return original_read_toml_file(file_path)

    monkeypatch.setattr(loader_module, "read_toml_file", read_fake_link)
    original_reload = loader.reload
    reload_count = 0

    def fail_once_then_reload() -> None:
        nonlocal reload_count
        reload_count += 1
        if reload_count == 1:
            raise OSError("simulated post-reset reload failure")
        original_reload()

    monkeypatch.setattr(loader, "reload", fail_once_then_reload)

    assert loader.reset_file("gui.toml") is False
    assert reload_count == 2
    assert links == {link_path: raw_target}
    assert target_path.read_bytes() == preimages["gui.toml"]
    assert loader.config_state_trusted is True
    assert loader.config.gui.window.default_mode == "batch"


def test_reset_file_removes_broken_symlink_instead_of_leaving_latent_override(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    loader, user_dir, _preimages = _loader_with_three_overrides(tmp_path)
    link_path = user_dir / "gui.toml"
    link_path.unlink()
    target_path = tmp_path / "linked" / "gui.toml"
    links = {link_path: Path("..") / "linked" / "gui.toml"}
    _install_fake_file_symlinks(monkeypatch, links)

    assert loader.reset_file("gui.toml") is True
    assert links == {}
    assert not target_path.exists()
    assert loader.config_state_trusted is True
    assert loader.config.gui.window.default_mode == "single"


def test_set_values_second_file_failure_restores_all_preimages(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.config import loader as loader_module

    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    original_write = loader_module.write_toml_file

    def fail_output_write(path: Path, data) -> None:
        if path.name == "output.toml":
            raise OSError("simulated later-file failure")
        original_write(path, data)

    monkeypatch.setattr(loader_module, "write_toml_file", fail_output_write)

    assert (
        loader.set_values(
            {
                "gui.window.default_mode": "single",
                "output.directory.mode": "source",
            }
        )
        is False
    )
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]
    assert loader.config.gui.window.default_mode == "batch"
    assert loader.config.output.directory.mode == "custom"


def test_reset_all_middle_unlink_failure_restores_all_preimages(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    original_unlink = Path.unlink
    unlink_count = 0

    def fail_second_unlink(path: Path, *args, **kwargs) -> None:
        nonlocal unlink_count
        unlink_count += 1
        if unlink_count == 2:
            raise OSError("simulated middle-file unlink failure")
        original_unlink(path, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", fail_second_unlink)

    assert loader.reset_all() is False
    assert unlink_count >= 2
    for rel_path, preimage in preimages.items():
        assert (user_dir / rel_path).read_bytes() == preimage
    assert loader.config.gui.window.default_mode == "batch"
    assert loader.config.output.directory.mode == "custom"
    assert loader.config.link.format.image_link_style == "markdown_link"


def test_reload_failure_after_document_write_restores_preimage(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import tomlkit

    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    original_reload = loader.reload
    reload_count = 0

    def fail_once_then_reload() -> None:
        nonlocal reload_count
        reload_count += 1
        if reload_count == 1:
            raise OSError("simulated post-write reload failure")
        original_reload()

    def mutate(doc) -> None:
        window = tomlkit.table()
        window["default_mode"] = "single"
        doc["window"] = window

    monkeypatch.setattr(loader, "reload", fail_once_then_reload)

    assert loader.update_file_document("gui.toml", mutate) is False
    assert reload_count == 2
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert loader.config.gui.window.default_mode == "batch"


def test_late_runtime_wiring_failure_restores_disk_and_effective_cache(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    original_config = loader.config.as_dict()
    original_wire_logging = loader._wire_logging
    wire_count = 0

    def fail_once_then_wire() -> None:
        nonlocal wire_count
        wire_count += 1
        if wire_count == 1:
            raise OSError("simulated late runtime wiring failure")
        original_wire_logging()

    monkeypatch.setattr(loader, "_wire_logging", fail_once_then_wire)

    assert (
        loader.set_values(
            {
                "gui.window.default_mode": "single",
                "output.directory.mode": "source",
            }
        )
        is False
    )
    assert wire_count == 2
    assert loader.config.as_dict() == original_config
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]


def test_double_reconciliation_failure_marks_effective_state_untrusted(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    original_wire_logging = loader._wire_logging

    def fail_logging_wiring() -> None:
        raise OSError("simulated persistent runtime wiring failure")

    monkeypatch.setattr(loader, "_wire_logging", fail_logging_wiring)

    with caplog.at_level(logging.ERROR):
        assert loader.set_value("gui.window.default_mode", "single") is False

    assert "transaction reconciliation failed" in caplog.text.lower()
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert loader.config_state_trusted is False

    monkeypatch.setattr(loader, "_wire_logging", original_wire_logging)
    loader.reload()
    assert loader.config_state_trusted is True
    assert loader.config.gui.window.default_mode == "batch"


def test_nested_persistence_from_document_mutator_fails_outer_operation_closed(
    tmp_path: Path,
) -> None:
    import tomlkit

    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    nested_results: list[bool] = []

    def mutate(doc) -> None:
        nested_results.append(loader.set_value("output.directory.mode", "source"))
        window = tomlkit.table()
        window["default_mode"] = "single"
        doc["window"] = window

    assert loader.update_file_document("gui.toml", mutate) is False
    assert nested_results == [False]
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]
    assert loader.config.gui.window.default_mode == "batch"
    assert loader.config.output.directory.mode == "custom"


def test_set_values_planning_failure_returns_false_without_mutation(tmp_path: Path) -> None:
    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)

    class Uncopyable:
        def __deepcopy__(self, _memo):
            raise RuntimeError("simulated deepcopy failure")

    assert loader.set_values({"gui.window.default_mode": Uncopyable()}) is False
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert loader.config.gui.window.default_mode == "batch"


def test_save_file_text_replace_failure_returns_false_without_mutation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime import toml_io

    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)

    def fail_replace(_source, _target) -> None:
        raise OSError("simulated editor replace failure")

    monkeypatch.setattr(toml_io.os, "replace", fail_replace)

    assert loader.save_file_text("gui.toml", '[window]\ndefault_mode = "single"\n') is False
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert loader.config.gui.window.default_mode == "batch"
    assert list(user_dir.glob(".gui.toml.*.tmp")) == []


def test_rollback_failure_reports_and_reconciles_actual_disk_state(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    from docwen_runtime.config import loader as loader_module

    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    original_write = loader_module.write_toml_file

    def fail_output_write(path: Path, data) -> None:
        if path.name == "output.toml":
            raise OSError("simulated later-file failure")
        original_write(path, data)

    def fail_gui_rollback(path: Path, _content: bytes) -> None:
        if path.name == "gui.toml":
            raise OSError("simulated rollback failure")
        raise AssertionError(f"unexpected rollback target: {path}")

    monkeypatch.setattr(loader_module, "write_toml_file", fail_output_write)
    monkeypatch.setattr(loader_module, "atomic_write_bytes", fail_gui_rollback, raising=False)

    with caplog.at_level(logging.ERROR):
        assert (
            loader.set_values(
                {
                    "gui.window.default_mode": "single",
                    "output.directory.mode": "source",
                }
            )
            is False
        )

    assert "rollback failed" in caplog.text.lower()
    assert (user_dir / "gui.toml").read_bytes() != preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]
    assert loader.config.gui.window.default_mode == "single"
    assert loader.config.output.directory.mode == "custom"


def test_rollback_comparison_read_failure_forces_preimage_restore(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.config import loader as loader_module

    loader, user_dir, preimages = _loader_with_three_overrides(tmp_path)
    original_write = loader_module.write_toml_file
    original_read_bytes = Path.read_bytes
    gui_read_count = 0

    def fail_output_write(path: Path, data) -> None:
        if path.name == "output.toml":
            raise OSError("simulated later-file failure")
        original_write(path, data)

    def fail_gui_comparison_once(path: Path) -> bytes:
        nonlocal gui_read_count
        if path == user_dir / "gui.toml":
            gui_read_count += 1
            if gui_read_count == 2:
                raise OSError("simulated rollback comparison failure")
        return original_read_bytes(path)

    monkeypatch.setattr(loader_module, "write_toml_file", fail_output_write)
    monkeypatch.setattr(Path, "read_bytes", fail_gui_comparison_once)

    assert (
        loader.set_values(
            {
                "gui.window.default_mode": "single",
                "output.directory.mode": "source",
            }
        )
        is False
    )
    assert gui_read_count >= 2
    assert (user_dir / "gui.toml").read_bytes() == preimages["gui.toml"]
    assert (user_dir / "output.toml").read_bytes() == preimages["output.toml"]
    assert loader.config.gui.window.default_mode == "batch"
