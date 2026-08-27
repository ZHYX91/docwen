"""
构建脚本 - 文件夹模式（GUI版本）

创建完整的应用程序文件夹结构，包含所有依赖文件。
输出到 dist 目录并添加版本号。

构建流程：
    1. 读取构建版本号（不写回源码；可通过 --version 或 CI tag 覆盖）
    2. 清理构建目录
    3. Cython 编译核心模块
    4. PyInstaller 打包
    5. 部署文件整理（资源、README、许可证）
    6. 构建验证

使用方式：
    python scripts/build/build.py              # 完整构建（从项目根目录执行）
    python scripts/build/build.py --skip-cython  # 跳过 Cython 编译（从项目根目录执行）
    python scripts/build/build.py --cli-only     # 仅构建 CLI（适合 Linux/CI）
    python scripts/build/build.py --version 0.8.2  # 覆盖构建版本号（不写回源码）

依赖：
    - PyInstaller: pip install pyinstaller
    - Cython: pip install cython（可选，用于编译核心模块）
"""

import argparse
import contextlib
import datetime
import importlib.metadata
import importlib.util
import logging
import os
import platform
import re
import shutil
import stat
import subprocess
import sys
import time
import tomllib
from pathlib import Path, PurePosixPath
from typing import cast

try:
    from scripts.build.payload_normalization import normalize_base_library_zip, normalize_packaged_record_files
except ModuleNotFoundError:  # Direct ``python scripts/build/build.py`` execution.
    from payload_normalization import normalize_base_library_zip, normalize_packaged_record_files

try:
    import PyInstaller.__main__ as pyinstaller_main
except ModuleNotFoundError:
    pyinstaller_main = None

# ==================== 平台常量 ====================
_system = platform.system()
_machine = platform.machine()
IS_WINDOWS = _system == "Windows"
IS_LINUX = _system == "Linux"
IS_MACOS = _system == "Darwin"

ARCH_TAG = {
    "x86_64": "x64",
    "AMD64": "x64",
    "aarch64": "arm64",
    "arm64": "arm64",
}.get(_machine, _machine.lower())

PLATFORM_TAG = {
    "Windows": f"win-{ARCH_TAG}",
    "Linux": f"linux-{ARCH_TAG}",
    "Darwin": f"macos-{ARCH_TAG}",
}.get(_system, f"{_system.lower()}-{ARCH_TAG}")

EXE_NAME = "DocWen.exe" if IS_WINDOWS else "DocWen"
CLI_EXE_NAME = "DocWenCLI.exe" if IS_WINDOWS else "DocWenCLI"
ICON_REL_PATH = "assets/icon.ico" if IS_WINDOWS else "assets/icon.png"

# ==================== 路径配置 ====================
# 构建脚本在 scripts/build 目录，项目根目录向上两级
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
LOGS_DIR = PROJECT_ROOT / "logs"


def configure_build_work_root(work_root: Path) -> None:
    """Redirect mutable build, dist, spec, and log state to one isolated root."""

    global DIST_DIR, BUILD_DIR, LOGS_DIR
    if not work_root.is_absolute():
        raise ValueError("build_work_root_must_be_absolute")
    work_root.mkdir(parents=True, exist_ok=True)
    if work_root.is_symlink() or (hasattr(work_root, "is_junction") and work_root.is_junction()):
        raise ValueError("build_work_root_must_not_be_link")
    DIST_DIR = work_root / "dist"
    BUILD_DIR = work_root / "build"
    LOGS_DIR = work_root / "logs"


def _pyinstaller_output_args() -> list[str]:
    """Return output paths that keep every PyInstaller artifact under ``BUILD_DIR``."""

    spec_dir = BUILD_DIR / "spec"
    spec_dir.mkdir(parents=True, exist_ok=True)
    return [
        f"--distpath={DIST_DIR}",
        f"--workpath={BUILD_DIR}",
        f"--specpath={spec_dir}",
    ]


@contextlib.contextmanager
def _isolated_pyinstaller_path():
    """Keep unrelated native toolchains from contaminating Windows payloads.

    PyInstaller resolves transitive DLL names through ``PATH``. Developer
    shells commonly prepend Poppler, Qt, or media toolchains whose DLLs can
    share names with Windows system libraries. The frozen program would then
    bundle whichever unrelated copy happened to appear first.
    """

    if not IS_WINDOWS:
        yield
        return

    original_path = os.environ.get("PATH")
    system_root = Path(os.environ.get("SYSTEMROOT", r"C:\Windows"))
    candidates = (
        Path(sys.executable).resolve().parent,
        Path(sys.base_prefix).resolve(),
        Path(sys.base_prefix).resolve() / "DLLs",
        system_root / "System32",
        system_root,
    )
    safe_entries: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        normalized = os.path.normcase(str(candidate))
        if candidate.is_dir() and normalized not in seen:
            safe_entries.append(str(candidate))
            seen.add(normalized)

    os.environ["PATH"] = os.pathsep.join(safe_entries)
    try:
        yield
    finally:
        if original_path is None:
            os.environ.pop("PATH", None)
        else:
            os.environ["PATH"] = original_path


# Workspace package source roots used by PyInstaller.
_PACKAGES_DIR = PROJECT_ROOT / "packages"
_PACKAGE_SRC_DIRS: list[Path] = sorted(
    source_root
    for source_root in _PACKAGES_DIR.rglob("src")
    if source_root.is_dir() and (source_root.parent / "pyproject.toml").is_file()
)

_PYMUPDF_LAYOUT_IMPORT_PACKAGE = "pymupdf.layout"
_PYINSTALLER_COMMON_COLLECT_ALL_TARGETS = (
    "rapidocr_onnxruntime",
    "pymupdf4llm",
    _PYMUPDF_LAYOUT_IMPORT_PACKAGE,
    "easyofd",
)
# qfluentwidgets exposes the widgets DocWen imports through ordinary Python
# imports.  Collecting the whole distribution also pulls its unused multimedia
# surface and, transitively, QtNetwork/TLS into the frozen GUI.
_PYINSTALLER_GUI_COLLECT_ALL_TARGETS: tuple[str, ...] = ()
_PYINSTALLER_COMMON_COLLECT_DATA_TARGETS = (
    "onnxruntime",
    "latex2mathml",
    "docx",
)
_PYINSTALLER_COMMON_COLLECT_SUBMODULE_TARGETS = ("docwen_runtime.ipc",)
_PYINSTALLER_EGRESS_RUNTIME_HOOK = _PACKAGES_DIR / "bundle" / "src" / "docwen_bundle" / "pyi_runtime_egress_guard.py"
_PYINSTALLER_HOOKS_DIR = PROJECT_ROOT / "scripts" / "build" / "pyinstaller_hooks"


def _pyinstaller_collection_args(option: str, targets: tuple[str, ...]) -> list[str]:
    return [f"--{option}={target}" for target in targets]


def _validate_pyinstaller_package_targets(targets: tuple[str, ...]) -> None:
    """Fail before packaging when a PyInstaller collection target is invalid."""

    invalid_targets: list[str] = []
    for target in dict.fromkeys(targets):
        try:
            spec = importlib.util.find_spec(target)
        except (ImportError, ModuleNotFoundError, AttributeError, ValueError):
            spec = None
        if spec is None or spec.submodule_search_locations is None:
            invalid_targets.append(target)
    if invalid_targets:
        raise RuntimeError(f"pyinstaller_collection_targets_invalid: {invalid_targets}")


def _validate_pymupdf_layout_pyinstaller_data_collection() -> None:
    """Prove PyInstaller will collect every resource in the split distribution."""

    from PyInstaller.utils.hooks import collect_data_files

    from docwen_runtime.pymupdf_layout_resources import (
        PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT,
        pymupdf_layout_resource_paths,
        verify_installed_pymupdf_layout_distribution,
    )

    source_verification = verify_installed_pymupdf_layout_distribution()
    if not source_verification.available:
        raise RuntimeError(f"pyinstaller_pymupdf_layout_source_contract_failed:{source_verification.reason}")

    collected_resources: set[str] = set()
    for source, destination in collect_data_files(_PYMUPDF_LAYOUT_IMPORT_PACKAGE):
        destination_path = PurePosixPath(destination.replace("\\", "/"))
        try:
            relative_destination = destination_path.relative_to(PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT)
        except ValueError:
            continue
        collected_resources.add((relative_destination / Path(source).name).as_posix())

    required_resources = set(pymupdf_layout_resource_paths())
    missing_resources = sorted(required_resources - collected_resources)
    if missing_resources:
        raise RuntimeError(f"pyinstaller_pymupdf_layout_data_collection_incomplete: {missing_resources}")


def _verify_pyinstaller_runtime_hook_order(build_name: str, *, entry_name: str) -> None:
    """Require DocWen's guard hook to precede every built-in runtime hook."""

    toc_candidates = sorted((BUILD_DIR / build_name).glob("PKG-*.toc"))
    if not toc_candidates:
        raise RuntimeError(f"pyinstaller_runtime_hook_toc_unavailable:{build_name}")
    if len(toc_candidates) != 1:
        names = ",".join(path.name for path in toc_candidates)
        raise RuntimeError(f"pyinstaller_runtime_hook_toc_ambiguous:{build_name}:{names}")
    toc_path = toc_candidates[0]
    try:
        toc_text = toc_path.read_text(encoding="utf-8")
    except OSError as exc:
        raise RuntimeError(f"pyinstaller_runtime_hook_toc_unavailable:{build_name}") from exc

    guard_index = toc_text.find("pyi_runtime_egress_guard")
    entry_index = toc_text.find(entry_name)
    built_in_indices = [
        index for match in re.finditer(r"['\"]pyi_rth_[^'\"]+['\"]", toc_text) if (index := match.start()) >= 0
    ]
    if guard_index < 0 or entry_index < 0:
        raise RuntimeError(f"pyinstaller_runtime_hook_or_entry_missing:{build_name}")
    if guard_index >= entry_index or (built_in_indices and guard_index >= min(built_in_indices)):
        raise RuntimeError(f"pyinstaller_runtime_hook_order_invalid:{build_name}")


# ==================== 日志系统配置 ====================
class BuildLogger:
    """构建日志管理器 - 同时输出到控制台和文件"""

    def __init__(self):
        self.start_time = datetime.datetime.now()
        self.step_times = {}
        self.current_step = None

        # 创建 logs 目录
        LOGS_DIR.mkdir(exist_ok=True)

        # 日志文件名包含时间戳
        log_filename = f"build_{self.start_time.strftime('%Y%m%d_%H%M%S')}.log"
        self.log_file = LOGS_DIR / log_filename

        # 配置日志格式
        log_format = "%(asctime)s [%(levelname)s] %(message)s"
        date_format = "%Y-%m-%d %H:%M:%S"

        # 创建 logger
        self.logger = logging.getLogger("BuildLogger")
        self.logger.setLevel(logging.DEBUG)
        self.logger.propagate = False  # 禁止传播到根记录器

        # 清除已有的处理器
        self.logger.handlers.clear()

        # 文件处理器（记录所有级别）
        self.file_handler = logging.FileHandler(self.log_file, encoding="utf-8")
        self.file_handler.setLevel(logging.DEBUG)
        self.file_handler.setFormatter(logging.Formatter(log_format, date_format))

        # 控制台处理器（只显示 INFO 及以上）
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(logging.Formatter(log_format, date_format))

        self.logger.addHandler(self.file_handler)
        self.logger.addHandler(console_handler)

        # 配置根日志记录器，捕获 PyInstaller 等第三方库的日志到文件
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.DEBUG)
        root_logger.handlers.clear()
        root_file_handler = logging.FileHandler(self.log_file, encoding="utf-8")
        root_file_handler.setLevel(logging.DEBUG)
        root_file_handler.setFormatter(logging.Formatter(log_format, date_format))
        root_logger.addHandler(root_file_handler)

        self.info(f"日志文件: {self.log_file}")
        self._log_environment_info()

    def _log_environment_info(self):
        """记录环境信息"""
        self.info("=" * 70)
        self.info("构建环境信息")
        self.info("=" * 70)
        self.info(f"Python 版本: {sys.version}")
        self.info(f"Python 路径: {sys.executable}")
        self.info(f"操作系统: {platform.platform()}")
        self.info(f"架构: {platform.machine()}")
        self.info(f"项目根目录: {PROJECT_ROOT}")

        # 记录关键依赖版本
        try:
            self.info(f"PyInstaller 版本: {importlib.metadata.version('pyinstaller')}")
        except Exception:
            self.warning("无法获取 PyInstaller 版本")

        try:
            self.info(f"Cython 版本: {importlib.metadata.version('Cython')}")
        except Exception:
            self.warning("无法获取 Cython 版本")

        self.info("=" * 70 + "\n")

    def start_step(self, step_name: str):
        """开始一个构建步骤"""
        if self.current_step:
            self.end_step()

        self.current_step = step_name
        self.step_times[step_name] = {"start": time.time()}
        self.info(f"\n{'=' * 70}")
        self.info(f"开始步骤: {step_name}")
        self.info(f"{'=' * 70}")

    def end_step(self):
        """结束当前构建步骤"""
        if self.current_step:
            elapsed = time.time() - self.step_times[self.current_step]["start"]
            self.step_times[self.current_step]["elapsed"] = elapsed
            self.info(f"步骤 '{self.current_step}' 完成，耗时: {elapsed:.2f}秒")
            self.current_step = None

    def info(self, msg: str):
        self.logger.info(msg)

    def warning(self, msg: str):
        self.logger.warning(msg)

    def error(self, msg: str):
        self.logger.error(msg)

    def debug(self, msg: str):
        self.logger.debug(msg)

    def print_summary(self):
        """打印构建摘要"""
        if self.current_step:
            self.end_step()

        total_time = time.time() - self.start_time.timestamp()

        self.info("\n" + "=" * 70)
        self.info("构建摘要")
        self.info("=" * 70)
        self.info(f"总耗时: {total_time:.2f}秒 ({total_time / 60:.2f}分钟)")
        self.info("\n各步骤耗时:")

        for step_name, times in self.step_times.items():
            if "elapsed" in times:
                elapsed = times["elapsed"]
                percentage = (elapsed / total_time * 100) if total_time > 0 else 0
                self.info(f"  - {step_name}: {elapsed:.2f}秒 ({percentage:.1f}%)")

        self.info("=" * 70)
        self.info(f"日志文件已保存到: {self.log_file}")


# 全局日志对象
logger: BuildLogger = cast(BuildLogger, None)


def init_logger() -> BuildLogger:
    """初始化日志系统"""
    global logger
    logger = BuildLogger()
    return logger


def read_version() -> str:
    """
    从当前项目元数据读取版本号（不写回文件）。

    版本号应为纯语义化版本（SemVer），例如：0.8.1
    """
    project_file = PROJECT_ROOT / "pyproject.toml"
    if project_file.is_file():
        data = tomllib.loads(project_file.read_text(encoding="utf-8"))
        project = data.get("project")
        if isinstance(project, dict):
            version = project.get("version")
            if isinstance(version, str) and version.strip():
                return version.strip()

    version_file = _PACKAGES_DIR / "bundle" / "src" / "docwen_bundle" / "__init__.py"
    if not version_file.is_file():
        raise FileNotFoundError(f"未找到当前架构版本文件: {project_file} 或 {version_file}")
    content = version_file.read_text(encoding="utf-8")
    version_pattern = r'__version__\s*=\s*["\']([^"\']+)["\']'
    match = re.search(version_pattern, content)
    if not match:
        raise RuntimeError(
            f'版本文件 {version_file} 中未找到 __version__ 定义，请确保文件包含如 __version__ = "X.Y.Z" 的声明'
        )
    return match.group(1).strip()


def _strip_v_prefix(v: str) -> str:
    s = (v or "").strip()
    if s.lower().startswith("v") and re.fullmatch(r"v\d+\.\d+\.\d+.*", s):
        return s[1:]
    return s


def resolve_build_version(version_override: str | None) -> str:
    """
    解析构建版本号（只读），优先级：
    1) --version 显式传入
    2) GITHUB_REF_NAME（tag: vX.Y.Z）
    3) DOCWEN_BUILD_VERSION 环境变量
    4) 源码 __version__
    """
    override = (version_override or "").strip()
    if override:
        return _strip_v_prefix(override)

    ref_name = (os.environ.get("GITHUB_REF_NAME") or "").strip()
    if ref_name:
        return _strip_v_prefix(ref_name)

    env_v = (os.environ.get("DOCWEN_BUILD_VERSION") or "").strip()
    if env_v:
        return _strip_v_prefix(env_v)

    return read_version()


def _remove_readonly_onerror(func, path, excinfo):
    """移除只读属性并重试删除（Python < 3.12 的 onerror 回调）"""
    os.chmod(path, stat.S_IWRITE)
    func(path)


def _remove_readonly_onexc(func, path, exc):
    """移除只读属性并重试删除（Python >= 3.12 的 onexc 回调）"""
    os.chmod(path, stat.S_IWRITE)
    func(path)


def force_remove_directory(path: Path) -> bool:
    """
    强制删除目录，处理权限问题

    Args:
        path: 要删除的目录路径

    Returns:
        True 删除成功，False 删除失败
    """
    if not path.exists():
        return True

    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            # Python 3.12+ 使用 onexc 替代已弃用的 onerror
            shutil.rmtree(path, onexc=_remove_readonly_onexc)
            logger.info(f"成功删除目录: {path}")
            return True
        except PermissionError as e:
            if attempt < max_attempts - 1:
                logger.warning(f"删除目录失败，重试中 ({attempt + 1}/{max_attempts}): {e}")
                time.sleep(2)

                # 尝试使用系统命令
                try:
                    if os.name == "nt":
                        subprocess.run(["rmdir", "/s", "/q", str(path)], shell=True, check=False)
                    else:
                        subprocess.run(["rm", "-rf", str(path)], check=False)

                    if not path.exists():
                        logger.info(f"使用系统命令成功删除目录: {path}")
                        return True
                except Exception:
                    pass
            else:
                logger.error(f"无法删除目录 {path}: {e}")
                return False
        except Exception as e:
            logger.error(f"删除目录失败: {path}, 错误: {e}")
            return False

    return False


def copytree_robust(src: Path, dst: Path) -> None:
    try:
        shutil.copytree(src, dst, dirs_exist_ok=True)
        return
    except OSError as e:
        if not isinstance(e, PermissionError) and getattr(e, "winerror", None) != 5:
            raise
    dst.mkdir(parents=True, exist_ok=True)
    for root, dirs, files in os.walk(src):
        src_root = Path(root)
        rel = src_root.relative_to(src)
        dst_root = dst / rel
        dst_root.mkdir(parents=True, exist_ok=True)
        for d in dirs:
            (dst_root / d).mkdir(parents=True, exist_ok=True)
        for f in files:
            shutil.copy2(src_root / f, dst_root / f)


def compile_cython_modules() -> bool:
    """
    编译 Cython 模块

    调用 setup_cython.py 执行编译，输出记录到日志文件

    Returns:
        True 编译成功，False 编译失败
    """
    logger.info("开始编译 Cython 模块...")

    setup_cython_path = Path(__file__).parent / "setup_cython.py"
    output_dir = BUILD_DIR / "cython_out"
    work_dir = BUILD_DIR / "cython_work"

    try:
        # 强制子进程使用 UTF-8 编码，防止 Windows 下出现 GBK 编码错误
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        result = subprocess.run(
            [
                sys.executable,
                str(setup_cython_path),
                "compile",
                "--force",
                "--output-dir",
                str(output_dir),
                "--work-dir",
                str(work_dir),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=str(PROJECT_ROOT),
            env=env,
        )

        # 记录输出到日志
        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                if line.strip():
                    logger.debug(f"[Cython] {line}")
        if result.stderr:
            for line in result.stderr.strip().split("\n"):
                if line.strip():
                    logger.debug(f"[Cython] {line}")

        if result.returncode == 0:
            logger.info("Cython 编译成功完成!")
            with contextlib.suppress(Exception):
                force_remove_directory(work_dir)
            return True
        else:
            logger.error(f"Cython 编译失败，返回码: {result.returncode}")
            return False

    except Exception as e:
        logger.error(f"Cython 编译过程出错: {e}")
        return False


def prepare_staging_package_root(
    package_src_dirs: list[Path],
    *,
    cython_out_dir: Path,
) -> Path:
    """Create one deterministic import root with compiled modules overlaid."""

    staging_root = BUILD_DIR / "staging_packages"
    if staging_root.exists():
        force_remove_directory(staging_root)
    staging_root.mkdir(parents=True)

    for source_root in package_src_dirs:
        for entry in sorted(source_root.iterdir(), key=lambda item: item.name):
            if entry.name == "__pycache__" or entry.name.endswith(".egg-info"):
                continue
            destination = staging_root / entry.name
            if destination.exists():
                raise RuntimeError(f"duplicate top-level package in build roots: {entry.name}")
            if entry.is_dir():
                shutil.copytree(
                    entry,
                    destination,
                    ignore=shutil.ignore_patterns("*.pyd", "*.so", "*.pyx", "*.c", "__pycache__"),
                )
            elif entry.is_file():
                shutil.copy2(entry, destination)

    if cython_out_dir.is_dir():
        shutil.copytree(cython_out_dir, staging_root, dirs_exist_ok=True)
    return staging_root


def copy_readme_files(deploy_dir: Path) -> int:
    """Copy the single canonical README into the deployment root."""
    source = PROJECT_ROOT / "README.md"
    if not source.is_file():
        raise FileNotFoundError(f"canonical README is missing: {source}")
    shutil.copy2(source, deploy_dir / source.name)
    logger.info("README 文件已复制 (1 个)")
    return 1


def verify_build(deploy_dir: Path, *, with_cli: bool = True, with_gui: bool = True) -> bool:
    """
    验证构建产物完整性

    Args:
        deploy_dir: 部署目录路径

    Returns:
        True 验证通过，False 验证失败
    """
    logger.info("验证构建产物...")

    # 必需文件/目录
    required: list[str] = ["templates", "configs", "models", "_internal/docwen/i18n/locales", "README.md", "LICENSE"]
    if with_gui:
        required.insert(0, EXE_NAME)
        required.insert(1, ICON_REL_PATH)
    if with_cli:
        required.insert(1 if with_gui else 0, CLI_EXE_NAME)

    # 可选文件/目录
    optional = ["samples"]

    missing_required = []
    missing_optional = []

    for item in required:
        if not (deploy_dir / item).exists():
            missing_required.append(item)

    for item in optional:
        if not (deploy_dir / item).exists():
            missing_optional.append(item)

    # 报告结果
    if missing_required:
        logger.error(f"❌ 缺少必需文件: {missing_required}")
        return False

    if missing_optional:
        logger.warning(f"⚠️ 缺少可选文件: {missing_optional}")

    duplicated = []
    internal_root = deploy_dir / "_internal"
    if internal_root.exists():
        for name in ("models", "templates", "configs", "samples"):
            if (internal_root / name).exists():
                duplicated.append(f"_internal/{name}")
    if duplicated:
        logger.error(f"❌ 检测到资源重复（去重回归）: {duplicated}")
        return False

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    try:
        from scripts.release.packaged_resources import (
            REQUIRED_ASSET_FILES,
            missing_files,
            verify_common_resource_layout,
        )

        verify_common_resource_layout(deploy_dir, error_prefix="build")
        if with_gui:
            missing_assets = missing_files(deploy_dir / "assets", REQUIRED_ASSET_FILES)
            if missing_assets:
                raise RuntimeError(f"build_assets_missing: {missing_assets}")
            from scripts.release.verify_packaged_gui import main as verify_packaged_gui_main

            settings_gate_exit = verify_packaged_gui_main(
                [
                    "--binary-dir",
                    str(deploy_dir),
                    "--binary-name",
                    EXE_NAME,
                    "--settings-smoke",
                ]
            )
            if settings_gate_exit != 0:
                raise RuntimeError(f"build_settings_smoke_failed: exit={settings_gate_exit}")
    except RuntimeError as exc:
        logger.error(f"❌ 发布资源逐文件验证失败: {exc}")
        return False

    # 检查可执行文件大小
    exe_path = deploy_dir / EXE_NAME
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        logger.info(f"可执行文件大小: {size_mb:.2f} MB")

    logger.info("✅ 构建验证通过")
    return True


def build_app(
    skip_cython: bool = False,
    *,
    with_cli: bool = True,
    with_gui: bool = True,
    version_override: str | None = None,
) -> tuple[str, Path] | None:
    """
    构建应用程序（文件夹模式）

    Args:
        skip_cython: 是否跳过 Cython 编译

    Returns:
        版本号字符串，失败返回 None
    """
    global logger

    if pyinstaller_main is None:
        logger.error("缺少 PyInstaller 依赖，请先安装 pyinstaller 后再运行构建")
        return None

    collection_targets = (
        *_PYINSTALLER_COMMON_COLLECT_ALL_TARGETS,
        *_PYINSTALLER_COMMON_COLLECT_DATA_TARGETS,
        *_PYINSTALLER_COMMON_COLLECT_SUBMODULE_TARGETS,
        *(_PYINSTALLER_GUI_COLLECT_ALL_TARGETS if with_gui else ()),
    )
    try:
        _validate_pyinstaller_package_targets(collection_targets)
        _validate_pymupdf_layout_pyinstaller_data_collection()
    except RuntimeError as exc:
        logger.error(f"PyInstaller 收集目标预检失败: {exc}")
        return None

    # 1. 版本管理（只读，不写回源码）
    logger.start_step("版本管理")
    version = resolve_build_version(version_override)
    logger.info(f"构建版本: {version}")
    logger.end_step()

    # 2. 清理构建目录
    logger.start_step("清理构建目录")
    logger.info("清理 dist 和 build 目录...")
    if not force_remove_directory(DIST_DIR) or not force_remove_directory(BUILD_DIR):
        logger.error("无法清理既有构建目录；为避免混入陈旧产物，已中止构建")
        return None
    logger.end_step()

    # 3. Cython 编译
    cython_ok = False
    if not skip_cython:
        logger.start_step("Cython 编译")
        cython_ok = compile_cython_modules()
        if not cython_ok:
            logger.error("Cython 编译失败；如需纯 Python 构建，请显式使用 --skip-cython")
            return None
        logger.info("Cython 编译成功，将使用编译后的模块构建")
        logger.end_step()
    else:
        logger.info("跳过 Cython 编译")

    package_src_dirs = list(_PACKAGE_SRC_DIRS)
    staging_package_root: Path | None = None
    if cython_ok:
        try:
            staging_package_root = prepare_staging_package_root(
                package_src_dirs,
                cython_out_dir=BUILD_DIR / "cython_out",
            )
        except Exception as e:
            logger.error(f"创建 Cython staging package root 失败: {e}")
            return None
        package_src_dirs = [staging_package_root]

    # 创建构建目录
    DIST_DIR.mkdir(exist_ok=True)
    BUILD_DIR.mkdir(exist_ok=True)

    # 资源路径
    templates_src = PROJECT_ROOT / "templates"
    configs_src = PROJECT_ROOT / "configs"
    assets_src = PROJECT_ROOT / "assets"
    models_src = PROJECT_ROOT / "models"
    samples_src = PROJECT_ROOT / "samples"
    i18n_locales_src = PROJECT_ROOT / "i18n" / "locales"
    contract_schemas_src = PROJECT_ROOT / "contracts" / "schemas"
    icon_path = assets_src / "icon.ico" if IS_WINDOWS else assets_src / "icon.png"

    gui_entry = _PACKAGES_DIR / "bundle" / "src" / "docwen_bundle" / "pyi_gui_entry.py"
    cli_entry = _PACKAGES_DIR / "bundle" / "src" / "docwen_bundle" / "pyi_cli_entry.py"
    logger.info(f"使用 GUI PyInstaller 入口: {gui_entry}")
    logger.info(f"使用 CLI PyInstaller 入口: {cli_entry}")
    if with_gui and not gui_entry.exists():
        logger.error(f"GUI 入口脚本不存在: {gui_entry}")
        return None
    if with_cli and not cli_entry.exists():
        logger.error(f"CLI 入口脚本不存在: {cli_entry}")
        return None
    if not _PYINSTALLER_EGRESS_RUNTIME_HOOK.is_file():
        logger.error(f"出站保护 runtime hook 不存在: {_PYINSTALLER_EGRESS_RUNTIME_HOOK}")
        return None
    if not _PYINSTALLER_HOOKS_DIR.is_dir():
        logger.error(f"PyInstaller hooks 目录不存在: {_PYINSTALLER_HOOKS_DIR}")
        return None

    common_excludes = [
        "--exclude-module=poplib",
        "--exclude-module=imaplib",
        "--exclude-module=telnetlib",
        "--exclude-module=nntplib",
        "--exclude-module=xmlrpc",
        "--exclude-module=PySide6.QtMultimedia",
        "--exclude-module=PySide6.QtMultimediaWidgets",
        "--exclude-module=PySide6.QtNetwork",
        "--exclude-module=PySide6.QtNetworkAuth",
        "--exclude-module=PySide6.QtWebSockets",
    ]
    if IS_LINUX or IS_MACOS:
        common_excludes.extend(
            [
                "--exclude-module=win32com",
                "--exclude-module=win32api",
                "--exclude-module=pythoncom",
                "--exclude-module=pywintypes",
            ]
        )

    # Canonical package roots (PyInstaller --paths).
    _pkg_path_args: list[str] = []
    for _d in package_src_dirs:
        _pkg_path_args.append("--paths")
        _pkg_path_args.append(str(_d))

    gui_build_args: list[str] = []
    if with_gui:
        gui_build_args = [
            str(gui_entry),
            "--name=DocWen",
            "--onedir",
            "--clean",
            "--noconfirm",
            *_pyinstaller_output_args(),
            f"--additional-hooks-dir={_PYINSTALLER_HOOKS_DIR}",
            f"--runtime-hook={_PYINSTALLER_EGRESS_RUNTIME_HOOK}",
            *common_excludes,
            *_pyinstaller_collection_args("collect-all", _PYINSTALLER_COMMON_COLLECT_ALL_TARGETS),
            *_pyinstaller_collection_args("collect-all", _PYINSTALLER_GUI_COLLECT_ALL_TARGETS),
            *_pyinstaller_collection_args("collect-data", _PYINSTALLER_COMMON_COLLECT_DATA_TARGETS),
            "--copy-metadata=PyYAML",
            "--copy-metadata=tabulate",
            *_pyinstaller_collection_args("collect-submodules", _PYINSTALLER_COMMON_COLLECT_SUBMODULE_TARGETS),
            # 新架构 packages — 显式收集以支持 PyInstaller 静态分析
            "--hidden-import=docwen_core",
            "--hidden-import=docwen_application",
            "--hidden-import=docwen_runtime",
            "--hidden-import=docwen_bundle",
            "--hidden-import=docwen_gui",
            "--hidden-import=docwen_cli",
            "--hidden-import=docwen_plugin_document",
            "--hidden-import=docwen_plugin_presentation",
            "--hidden-import=docwen_plugin_spreadsheet",
            "--hidden-import=docwen_plugin_markup",
            "--hidden-import=docwen_plugin_image",
            "--hidden-import=docwen_plugin_print",
            "--hidden-import=docwen_plugin_layout",
            "--hidden-import=docwen_plugin_markdown",
            "--hidden-import=docwen_plugin_optimizer_gongwen",
            "--hidden-import=docwen_plugin_optimizer_invoice_cn",
            "--hidden-import=docwen_plugin_proofread",
            *_pkg_path_args,
        ]

        if icon_path.exists():
            gui_build_args.append(f"--icon={icon_path}")
            logger.info(f"添加应用程序图标: {icon_path}")
        else:
            logger.warning(f"图标文件不存在: {icon_path}")

        if os.name == "nt":
            gui_build_args.append("--noconsole")
            logger.info("Windows 平台: 隐藏控制台窗口")

    # 数据文件
    data_files = [
        f"{i18n_locales_src}{os.pathsep}docwen/i18n/locales",
        f"{contract_schemas_src}{os.pathsep}docwen_cli/contracts",
    ]

    for data_file in data_files:
        if with_gui:
            gui_build_args.append(f"--add-data={data_file}")

    # 4. PyInstaller 构建
    logger.start_step("PyInstaller 构建")
    try:
        if with_gui:
            logger.info("开始 PyInstaller 构建 (GUI)...")
            with _isolated_pyinstaller_path():
                pyinstaller_main.run(gui_build_args)
            _verify_pyinstaller_runtime_hook_order("DocWen", entry_name="pyi_gui_entry")
            logger.info("PyInstaller GUI 构建成功完成!")
        logger.end_step()

        if with_cli:
            logger.start_step("PyInstaller 构建 (CLI)")
            cli_build_args: list[str] = [
                str(cli_entry),
                "--name=DocWenCLI",
                "--onedir",
                "--clean",
                "--noconfirm",
                *_pyinstaller_output_args(),
                f"--additional-hooks-dir={_PYINSTALLER_HOOKS_DIR}",
                f"--runtime-hook={_PYINSTALLER_EGRESS_RUNTIME_HOOK}",
                *common_excludes,
                *_pyinstaller_collection_args("collect-all", _PYINSTALLER_COMMON_COLLECT_ALL_TARGETS),
                *_pyinstaller_collection_args("collect-data", _PYINSTALLER_COMMON_COLLECT_DATA_TARGETS),
                "--copy-metadata=PyYAML",
                "--copy-metadata=tabulate",
                *_pyinstaller_collection_args("collect-submodules", _PYINSTALLER_COMMON_COLLECT_SUBMODULE_TARGETS),
                # 新架构 packages — 显式收集以支持 PyInstaller 静态分析
                "--hidden-import=docwen_core",
                "--hidden-import=docwen_application",
                "--hidden-import=docwen_runtime",
                "--hidden-import=docwen_bundle",
                "--hidden-import=docwen_cli",
                "--hidden-import=docwen_plugin_document",
                "--hidden-import=docwen_plugin_presentation",
                "--hidden-import=docwen_plugin_spreadsheet",
                "--hidden-import=docwen_plugin_markup",
                "--hidden-import=docwen_plugin_image",
                "--hidden-import=docwen_plugin_print",
                "--hidden-import=docwen_plugin_layout",
                "--hidden-import=docwen_plugin_markdown",
                "--hidden-import=docwen_plugin_optimizer_gongwen",
                "--hidden-import=docwen_plugin_optimizer_invoice_cn",
                "--hidden-import=docwen_plugin_proofread",
                *_pkg_path_args,
            ]
            for data_file in data_files:
                cli_build_args.append(f"--add-data={data_file}")
            with _isolated_pyinstaller_path():
                pyinstaller_main.run(cli_build_args)
            _verify_pyinstaller_runtime_hook_order("DocWenCLI", entry_name="pyi_cli_entry")
            logger.info("PyInstaller CLI 构建成功完成!")
            logger.end_step()

        if staging_package_root is not None:
            with contextlib.suppress(Exception):
                force_remove_directory(staging_package_root)

        # 5. 部署文件整理
        logger.start_step("部署文件整理")

        deploy_base = "DocWen" if with_gui else "DocWenCLI"
        deploy_dir_name = f"{deploy_base}_v{version}_{PLATFORM_TAG}"
        deploy_dir = DIST_DIR / deploy_dir_name

        if with_gui:
            pyinstaller_output = DIST_DIR / "DocWen"
            if pyinstaller_output.exists():
                shutil.move(str(pyinstaller_output), str(deploy_dir))
                logger.info(f"重命名输出到: {deploy_dir}")
            else:
                logger.error(f"PyInstaller 输出目录不存在: {pyinstaller_output}")
                return None
        else:
            deploy_dir.mkdir(parents=True, exist_ok=True)
            logger.info(f"创建部署目录: {deploy_dir}")

        cli_output_dir = DIST_DIR / "DocWenCLI"
        if with_cli:
            if not cli_output_dir.exists():
                logger.error(f"PyInstaller CLI 输出目录不存在: {cli_output_dir}")
                return None

            cli_exe_src = cli_output_dir / CLI_EXE_NAME
            cli_exe_dst = deploy_dir / CLI_EXE_NAME
            shutil.copy2(cli_exe_src, cli_exe_dst)

            src_internal = cli_output_dir / "_internal"
            dst_internal = deploy_dir / "_internal"
            if src_internal.exists():
                copytree_robust(src_internal, dst_internal)

        # 补充复制资源文件
        resource_mappings = [
            (templates_src, deploy_dir / "templates"),
            (configs_src, deploy_dir / "configs"),
            (assets_src, deploy_dir / "assets"),
            (samples_src, deploy_dir / "samples"),
        ]

        for src, dst in resource_mappings:
            if not dst.exists() and src.exists():
                copytree_robust(src, dst)
                logger.info(f"复制: {src.name} -> {dst}")

        # OCR 模型
        models_dest = deploy_dir / "models"
        if not models_dest.exists() and models_src.exists() and list(models_src.iterdir()):
            copytree_robust(models_src, models_dest)
            logger.info(f"复制 OCR 模型到: {models_dest}")

        # CLI 独立目录资源补充
        # PyInstaller onedir 模式将 --add-data 数据放在 _internal/ 下，
        # 但运行时 get_project_root() 返回 exe 同级目录，需要顶层存在资源。
        # 在此将 templates/configs 补充复制到 CLI 原始输出目录顶层，
        # 确保 CLI zip 包解压后可直接运行且 CI 校验通过。
        if with_cli and cli_output_dir.exists():
            cli_resource_mappings = [
                (templates_src, cli_output_dir / "templates"),
                (models_src, cli_output_dir / "models"),
                (samples_src, cli_output_dir / "samples"),
                (configs_src, cli_output_dir / "configs"),
            ]
            for src, dst in cli_resource_mappings:
                if not dst.exists() and src.exists():
                    copytree_robust(src, dst)
                    logger.info(f"补充复制到 CLI 目录: {src.name} -> {dst}")

        # 复制 README 文件（所有语言版本，同级目录）
        copy_readme_files(deploy_dir)

        # 复制许可证文件
        license_files = ["LICENSE", "LICENSE_THIRD_PARTY.txt", "NOTICE.txt"]
        license_copied = 0
        for license_file in license_files:
            src = PROJECT_ROOT / license_file
            if src.exists():
                shutil.copy2(src, deploy_dir / license_file)
                logger.debug(f"复制许可证: {license_file}")
                license_copied += 1

        logger.info(f"许可证文件已复制 ({license_copied} 个)")

        if with_cli and cli_output_dir.exists():
            copy_readme_files(cli_output_dir)
            for license_file in license_files:
                src = PROJECT_ROOT / license_file
                if src.exists():
                    shutil.copy2(src, cli_output_dir / license_file)

        cli_deploy_dir: Path | None = None
        if with_gui and with_cli and cli_output_dir.exists():
            cli_deploy_dir_name = f"DocWenCLI_v{version}_{PLATFORM_TAG}"
            cli_deploy_dir = DIST_DIR / cli_deploy_dir_name
            if cli_deploy_dir.exists():
                force_remove_directory(cli_deploy_dir)
            shutil.move(str(cli_output_dir), str(cli_deploy_dir))

        normalization_targets = [deploy_dir]
        if cli_deploy_dir is not None:
            normalization_targets.append(cli_deploy_dir)
        for target in normalization_targets:
            record_result = normalize_packaged_record_files(target)
            base_library_result = normalize_base_library_zip(target)
            logger.info(f"规范化冻结载荷: {target.name}; RECORD={record_result}; base_library={base_library_result}")

        if cli_deploy_dir is not None and not verify_build(cli_deploy_dir, with_cli=True, with_gui=False):
            return None
        logger.end_step()

    except Exception as e:
        logger.error(f"构建失败: {e}")
        import traceback

        logger.debug(traceback.format_exc())
        return None

    return version, deploy_dir


def main():
    """主函数"""
    if IS_WINDOWS:
        try:
            stdout_reconfigure = getattr(sys.stdout, "reconfigure", None)
            if callable(stdout_reconfigure):
                stdout_reconfigure(encoding="utf-8", errors="replace")

            stderr_reconfigure = getattr(sys.stderr, "reconfigure", None)
            if callable(stderr_reconfigure):
                stderr_reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--skip-cython", action="store_true", help="跳过 Cython 编译")
    parser.add_argument("--gui-only", action="store_true", help="仅构建 GUI（不构建 CLI）")
    parser.add_argument("--cli-only", action="store_true", help="仅构建 CLI（不构建 GUI）")
    parser.add_argument("--version", type=str, default=None, help="构建版本号（覆盖源码 __version__，可带 v 前缀）")
    parser.add_argument(
        "--work-root",
        type=Path,
        default=None,
        help="将 build/spec/dist/logs 写入一个隔离的绝对目录",
    )
    args = parser.parse_args()

    if args.work_root is not None:
        configure_build_work_root(args.work_root.resolve())

    # 初始化日志
    init_logger()

    logger.info("=" * 70)
    logger.info("开始构建 DocWen")
    logger.info("=" * 70)

    skip_cython = bool(args.skip_cython)
    if args.gui_only and args.cli_only:
        logger.error("参数冲突：--gui-only 与 --cli-only 不能同时使用")
        sys.exit(2)

    with_gui = not bool(args.cli_only)
    with_cli = not bool(args.gui_only)

    try:
        result = build_app(
            skip_cython=skip_cython,
            with_cli=with_cli,
            with_gui=with_gui,
            version_override=args.version,
        )

        if result:
            _version, deploy_dir = result
            # 验证构建
            logger.start_step("构建验证")
            build_verified = verify_build(deploy_dir, with_cli=with_cli, with_gui=with_gui)
            logger.end_step()
            if not build_verified:
                logger.error("构建验证失败!")
                logger.print_summary()
                sys.exit(1)

            logger.info("\n" + "=" * 70)
            logger.info(f"✅ 构建成功完成! 软件已部署到 dist/{deploy_dir.name}")
            logger.info("=" * 70)
            logger.print_summary()
        else:
            logger.error("构建失败!")
            logger.print_summary()
            sys.exit(1)

    except Exception as e:
        logger.error(f"构建过程发生异常: {e}")
        import traceback

        logger.debug(traceback.format_exc())
        logger.print_summary()
        sys.exit(1)


if __name__ == "__main__":
    main()
