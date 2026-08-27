"""Repository checks for the release build script and PyInstaller hooks."""

from __future__ import annotations

import ast
import os
import textwrap
from pathlib import Path
from unittest.mock import Mock

import pytest

pytestmark = pytest.mark.unit


def test_build_version_reads_current_project_metadata(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.build import build

    (tmp_path / "pyproject.toml").write_text(
        textwrap.dedent(
            """
            [project]
            name = "docwen"
            version = "9.8.7"
            """
        ).strip(),
        encoding="utf-8",
    )
    monkeypatch.setattr(build, "PROJECT_ROOT", tmp_path)

    assert build.read_version() == "9.8.7"


def test_build_version_falls_back_to_bundle_package_metadata(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.build import build

    (tmp_path / "pyproject.toml").write_text(
        textwrap.dedent(
            """
            [project]
            name = "docwen"
            """
        ).strip(),
        encoding="utf-8",
    )
    bundle_dir = tmp_path / "packages" / "bundle" / "src" / "docwen_bundle"
    bundle_dir.mkdir(parents=True)
    (bundle_dir / "__init__.py").write_text('__version__ = "7.6.5"\n', encoding="utf-8")

    monkeypatch.setattr(build, "PROJECT_ROOT", tmp_path)
    monkeypatch.setattr(build, "_PACKAGES_DIR", tmp_path / "packages")

    assert build.read_version() == "7.6.5"


def test_build_copies_only_the_canonical_readme(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.build import build

    project_root = tmp_path / "checkout"
    deploy_dir = tmp_path / "deploy"
    project_root.mkdir()
    deploy_dir.mkdir()
    (project_root / "README.md").write_text("canonical\n", encoding="utf-8")
    monkeypatch.setattr(build, "PROJECT_ROOT", project_root)
    monkeypatch.setattr(build, "logger", Mock())

    assert build.copy_readme_files(deploy_dir) == 1
    assert (deploy_dir / "README.md").read_text(encoding="utf-8") == "canonical\n"
    assert sorted(path.name for path in deploy_dir.iterdir()) == ["README.md"]


def test_build_rejects_a_missing_canonical_readme(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.build import build

    deploy_dir = tmp_path / "deploy"
    deploy_dir.mkdir()
    monkeypatch.setattr(build, "PROJECT_ROOT", tmp_path / "checkout")

    with pytest.raises(FileNotFoundError, match="canonical README is missing"):
        build.copy_readme_files(deploy_dir)


def test_build_logger_calls_use_build_logger_signature() -> None:
    tree = ast.parse(Path("scripts/build/build.py").read_text(encoding="utf-8"))
    offenders: list[str] = []
    checked_names = {"debug", "error", "info", "warning"}
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        if not isinstance(func, ast.Attribute) or func.attr not in checked_names:
            continue
        if not isinstance(func.value, ast.Name) or func.value.id != "logger":
            continue
        if len(node.args) > 1:
            offenders.append(f"logger.{func.attr}:{node.lineno}")

    assert offenders == []


def test_build_script_uses_only_workspace_packages_and_bundle_entries() -> None:
    source = Path("scripts/build/build.py").read_text(encoding="utf-8")

    assert 'cli_deploy_dir_name = f"DocWenCLI_v{version}_{PLATFORM_TAG}"' in source
    assert "verify_build(cli_deploy_dir, with_cli=True, with_gui=False)" in source
    assert '_PACKAGES_DIR.rglob("src")' in source
    assert 'PROJECT_ROOT / "src"' not in source
    assert "gui_run.py" not in source
    assert "cli_run.py" not in source
    assert "clean_source_files_from_dist" not in source
    assert 'docwen_bundle" / "pyi_gui_entry.py' in source
    assert 'docwen_bundle" / "pyi_cli_entry.py' in source


def test_pyinstaller_specs_follow_isolated_build_root(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.build import build

    project_root = tmp_path / "checkout"
    project_root.mkdir()
    work_root = tmp_path / "external-work"
    monkeypatch.setattr(build, "PROJECT_ROOT", project_root)
    monkeypatch.setattr(build, "DIST_DIR", project_root / "dist")
    monkeypatch.setattr(build, "BUILD_DIR", project_root / "build")
    monkeypatch.setattr(build, "LOGS_DIR", project_root / "logs")

    build.configure_build_work_root(work_root)
    output_args = build._pyinstaller_output_args()

    assert output_args == [
        f"--distpath={work_root / 'dist'}",
        f"--workpath={work_root / 'build'}",
        f"--specpath={work_root / 'build' / 'spec'}",
    ]
    assert (work_root / "build" / "spec").is_dir()
    assert not list(project_root.glob("DocWen*.spec"))

    source = Path("scripts/build/build.py").read_text(encoding="utf-8")
    assert source.count("*_pyinstaller_output_args(),") == 2


def test_windows_pyinstaller_path_excludes_ambient_native_toolchains(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.build import build

    system_root = tmp_path / "Windows"
    (system_root / "System32").mkdir(parents=True)
    foreign_native = tmp_path / "foreign-poppler" / "bin"
    foreign_native.mkdir(parents=True)
    original = os.pathsep.join((str(foreign_native), str(system_root / "System32")))
    monkeypatch.setattr(build, "IS_WINDOWS", True)
    monkeypatch.setenv("SYSTEMROOT", str(system_root))
    monkeypatch.setenv("PATH", original)

    with build._isolated_pyinstaller_path():  # pyright: ignore[reportPrivateUsage]
        isolated = os.environ["PATH"].split(os.pathsep)
        assert str(foreign_native) not in isolated
        assert str(system_root / "System32") in isolated

    assert os.environ["PATH"] == original


def test_build_fails_closed_for_cleanup_and_cython_errors() -> None:
    source = Path("scripts/build/build.py").read_text(encoding="utf-8")

    assert "if not force_remove_directory(DIST_DIR) or not force_remove_directory(BUILD_DIR):" in source
    assert "为避免混入陈旧产物，已中止构建" in source
    assert "如需纯 Python 构建，请显式使用 --skip-cython" in source
    assert "Cython 编译失败，将继续使用纯 Python 模块构建" not in source


def test_cython_core_module_manifest_is_exact_and_fails_closed(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.build import setup_cython

    missing_from_repository = [module for module in setup_cython.CORE_MODULES if not Path(module).is_file()]
    assert missing_from_repository == []
    assert len(setup_cython.CORE_MODULES) == len(set(setup_cython.CORE_MODULES))

    monkeypatch.setattr(setup_cython, "CORE_MODULES", ["missing/module.py"])
    with pytest.raises(FileNotFoundError, match="拒绝生成不完整构建"):
        setup_cython.CythonBuilder(
            tmp_path,
            output_dir=tmp_path / "output",
            work_dir=tmp_path / "work",
        )


def test_cython_staging_combines_package_roots_and_compiled_overlay(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.build import build

    first_root = tmp_path / "packages" / "core" / "src"
    second_root = tmp_path / "packages" / "apps" / "gui" / "src"
    (first_root / "docwen_core").mkdir(parents=True)
    (second_root / "docwen_gui").mkdir(parents=True)
    (first_root / "docwen_core" / "__init__.py").write_text("", encoding="utf-8")
    (second_root / "docwen_gui" / "__init__.py").write_text("", encoding="utf-8")
    compiled_root = tmp_path / "compiled"
    (compiled_root / "docwen_core").mkdir(parents=True)
    (compiled_root / "docwen_core" / "models.pyd").write_bytes(b"compiled")
    monkeypatch.setattr(build, "BUILD_DIR", tmp_path / "build")

    staged = build.prepare_staging_package_root(
        [first_root, second_root],
        cython_out_dir=compiled_root,
    )

    assert (staged / "docwen_core" / "__init__.py").is_file()
    assert (staged / "docwen_gui" / "__init__.py").is_file()
    assert (staged / "docwen_core" / "models.pyd").read_bytes() == b"compiled"


def test_build_registers_egress_guard_as_custom_runtime_hook() -> None:
    source = Path("scripts/build/build.py").read_text(encoding="utf-8")
    hook_path = Path("packages/bundle/src/docwen_bundle/pyi_runtime_egress_guard.py")

    assert hook_path.is_file()
    assert source.count('f"--runtime-hook={_PYINSTALLER_EGRESS_RUNTIME_HOOK}"') == 2


def test_build_uses_qtgui_hook_without_network_dependent_optional_plugins() -> None:
    import runpy

    source = Path("scripts/build/build.py").read_text(encoding="utf-8")
    hook_path = Path("scripts/build/pyinstaller_hooks/hook-PySide6.QtGui.py")

    assert hook_path.is_file()
    assert source.count('f"--additional-hooks-dir={_PYINSTALLER_HOOKS_DIR}"') == 2

    hook = runpy.run_path(str(hook_path))
    binaries = hook["binaries"]
    omitted_stems = hook["_OMITTED_PLUGIN_STEMS"]
    plugin_stem = hook["_plugin_stem"]
    collected_stems = {plugin_stem(source_path) for source_path, _destination in binaries}

    assert omitted_stems == {"qpdf", "qtuiotouchplugin", "qtvirtualkeyboardplugin", "qvnc"}
    assert omitted_stems.isdisjoint(collected_stems)
    assert {"qjpeg", "qsvg"} <= collected_stems


def test_build_collects_every_explicitly_lazy_imported_settings_page() -> None:
    import inspect

    from docwen_gui.widgets.settings.dialog import _TAB_SPECS

    hook_path = Path("scripts/build/pyinstaller_hooks/hook-docwen_gui.widgets.settings.py")
    assert not hook_path.exists()
    for spec in _TAB_SPECS.values():
        factory_source = inspect.getsource(spec.factory)
        assert f"from .{spec.module_name} import {spec.class_name}" in factory_source
        assert "importlib" not in factory_source


def test_build_rejects_runtime_hook_after_builtin_hooks(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.build import build

    build_dir = tmp_path / "build"
    # PyInstaller numbers TOCs process-wide. The second onedir build in the
    # same process therefore uses PKG-01.toc rather than PKG-00.toc.
    toc_path = build_dir / "DocWen" / "PKG-01.toc"
    toc_path.parent.mkdir(parents=True)
    monkeypatch.setattr(build, "BUILD_DIR", build_dir)

    toc_path.write_text(
        "('pyi_runtime_egress_guard', 'PYSOURCE')\n('pyi_rth_pyside6', 'PYSOURCE')\n('pyi_gui_entry', 'PYSOURCE')\n",
        encoding="utf-8",
    )
    build._verify_pyinstaller_runtime_hook_order("DocWen", entry_name="pyi_gui_entry")

    toc_path.write_text(
        "('pyi_rth_pyside6', 'PYSOURCE')\n('pyi_runtime_egress_guard', 'PYSOURCE')\n('pyi_gui_entry', 'PYSOURCE')\n",
        encoding="utf-8",
    )
    with pytest.raises(RuntimeError, match="pyinstaller_runtime_hook_order_invalid"):
        build._verify_pyinstaller_runtime_hook_order("DocWen", entry_name="pyi_gui_entry")


def test_build_fails_closed_when_runtime_hook_toc_is_missing_or_ambiguous(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.build import build

    build_dir = tmp_path / "build"
    target = build_dir / "DocWenCLI"
    target.mkdir(parents=True)
    monkeypatch.setattr(build, "BUILD_DIR", build_dir)

    with pytest.raises(RuntimeError, match="pyinstaller_runtime_hook_toc_unavailable"):
        build._verify_pyinstaller_runtime_hook_order("DocWenCLI", entry_name="pyi_cli_entry")

    toc_text = "('pyi_runtime_egress_guard', 'PYSOURCE')\n('pyi_cli_entry', 'PYSOURCE')\n"
    (target / "PKG-00.toc").write_text(toc_text, encoding="utf-8")
    (target / "PKG-01.toc").write_text(toc_text, encoding="utf-8")
    with pytest.raises(RuntimeError, match="pyinstaller_runtime_hook_toc_ambiguous"):
        build._verify_pyinstaller_runtime_hook_order("DocWenCLI", entry_name="pyi_cli_entry")
