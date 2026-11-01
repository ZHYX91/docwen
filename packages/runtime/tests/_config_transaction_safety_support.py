"""Configuration preimages and privilege-independent symlink fault fixtures."""

from __future__ import annotations

import os
from pathlib import Path

import pytest


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
