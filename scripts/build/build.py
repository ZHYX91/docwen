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
    6. 源文件清理（移除已编译模块的 .py）
    7. 构建验证

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
import logging
import os
import platform
import re
import shutil
import stat
import subprocess
import sys
import time
from pathlib import Path, PurePosixPath
from typing import cast

try:
    import PyInstaller.__main__ as pyinstaller_main
except ModuleNotFoundError:
    pyinstaller_main = None

# ==================== 平台常量 ====================
_system = platform.system()
IS_WINDOWS = _system == "Windows"
IS_LINUX = _system == "Linux"
PLATFORM_TAG = {
    "Windows": "win",
    "Linux": "linux",
    "Darwin": "macos",
}.get(_system, _system.lower())

EXE_NAME = "DocWen.exe" if IS_WINDOWS else "DocWen"
CLI_EXE_NAME = "DocWenCLI.exe" if IS_WINDOWS else "DocWenCLI"
ICON_REL_PATH = "assets/icon.ico" if IS_WINDOWS else "assets/icon.png"

# ==================== 路径配置 ====================
# 构建脚本在 scripts/build 目录，项目根目录向上两级
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
SRC_DIR = PROJECT_ROOT / "src"
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
LOGS_DIR = PROJECT_ROOT / "logs"


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


def get_version() -> str:
    """
    获取或更新版本号

    版本格式: X.Y.Z.YYYYMMDD.R
    - X.Y.Z: 主版本.次版本.修订版本
    - YYYYMMDD: 构建日期
    - R: 当日修订号（从1开始）

    Returns:
        新版本号字符串
    """
    raise RuntimeError("get_version() 已弃用，请使用 read_version()/resolve_build_version()")


def read_version() -> str:
    """
    从源码读取版本号（不写回文件）。

    版本号应为纯语义化版本（SemVer），例如：0.8.1
    """
    version_file = SRC_DIR / "docwen" / "__init__.py"
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


def prepare_staging_src(src_dir: Path, *, cython_out_dir: Path) -> Path:
    staging_src = BUILD_DIR / "staging_src"
    if staging_src.exists():
        force_remove_directory(staging_src)
    shutil.copytree(
        src_dir,
        staging_src,
        ignore=shutil.ignore_patterns("*.pyd", "*.so", "*.pyx", "*.c"),
    )

    compiled_docwen = cython_out_dir / "docwen"
    staging_docwen = staging_src / "docwen"
    if compiled_docwen.exists() and staging_docwen.exists():
        shutil.copytree(compiled_docwen, staging_docwen, dirs_exist_ok=True)
    return staging_src


def clean_cython_artifacts_in_src() -> int:
    try:
        sys.path.insert(0, str(Path(__file__).parent))
        from setup_cython import CORE_MODULES
    except Exception:
        return 0

    deleted = 0
    for module_path in CORE_MODULES:
        try:
            posix_rel = PurePosixPath(module_path).relative_to("src")
        except Exception:
            posix_rel = PurePosixPath(module_path)

        src_py = SRC_DIR / Path(*posix_rel.parts)
        parent = src_py.parent
        stem = src_py.stem
        for pattern in (f"{stem}*.pyd", f"{stem}*.so"):
            for compiled in parent.glob(pattern):
                try:
                    compiled.unlink()
                    deleted += 1
                except Exception:
                    pass
    return deleted


def copy_readme_files(deploy_dir: Path) -> int:
    """
    复制所有 README 文件到部署目录（同级放置）

    Args:
        deploy_dir: 部署目录路径

    Returns:
        复制的文件数量
    """
    readme_files = [
        "README.md",  # 英文主版本
        "README_zh-CN.md",  # 简体中文
        "README_zh-TW.md",  # 繁体中文
        "README_de-DE.md",  # 德语
        "README_es-ES.md",  # 西班牙语
        "README_fr-FR.md",  # 法语
        "README_ja-JP.md",  # 日语
        "README_ko-KR.md",  # 韩语
        "README_pt-BR.md",  # 葡萄牙语
        "README_ru-RU.md",  # 俄语
        "README_vi-VN.md",  # 越南语
    ]

    copied = 0
    for readme in readme_files:
        src = PROJECT_ROOT / readme
        if src.exists():
            dst = deploy_dir / readme
            shutil.copy2(src, dst)
            logger.debug(f"复制 README: {readme}")
            copied += 1

    logger.info(f"README 文件已复制 ({copied} 个)")
    return copied


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
    required: list[str] = ["templates", "configs", "README.md", "LICENSE"]
    if with_gui:
        required.insert(0, EXE_NAME)
        required.insert(1, ICON_REL_PATH)
    if with_cli:
        required.insert(1 if with_gui else 0, CLI_EXE_NAME)

    # 可选文件/目录
    optional = [
        "models",
        "samples",
        "README_zh-CN.md",
    ]

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

    # 1. 版本管理（只读，不写回源码）
    logger.start_step("版本管理")
    version = resolve_build_version(version_override)
    logger.info(f"构建版本: {version}")
    logger.end_step()

    # 2. 清理构建目录
    logger.start_step("清理构建目录")
    logger.info("清理 dist 和 build 目录...")
    force_remove_directory(DIST_DIR)
    force_remove_directory(BUILD_DIR)
    removed = clean_cython_artifacts_in_src()
    if removed:
        logger.info(f"已从源码目录清理旧的 Cython 编译产物: {removed} 个")
    logger.end_step()

    # 3. Cython 编译
    cython_ok = False
    if not skip_cython:
        logger.start_step("Cython 编译")
        cython_ok = compile_cython_modules()
        if not cython_ok:
            logger.warning("Cython 编译失败，将继续使用纯 Python 模块构建")
        else:
            logger.info("Cython 编译成功，将使用编译后的模块构建")
        logger.end_step()
    else:
        logger.info("跳过 Cython 编译")

    effective_src_dir = SRC_DIR
    if cython_ok:
        try:
            effective_src_dir = prepare_staging_src(SRC_DIR, cython_out_dir=BUILD_DIR / "cython_out")
        except Exception as e:
            logger.warning(f"创建 staging 源码目录失败，将回退使用纯源码构建: {e}")
            effective_src_dir = SRC_DIR

    # 创建构建目录
    DIST_DIR.mkdir(exist_ok=True)

    # 资源路径
    templates_src = PROJECT_ROOT / "templates"
    configs_src = PROJECT_ROOT / "configs"
    assets_src = PROJECT_ROOT / "assets"
    models_src = PROJECT_ROOT / "models"
    samples_src = PROJECT_ROOT / "samples"
    i18n_locales_src = effective_src_dir / "docwen" / "i18n" / "locales"
    icon_path = assets_src / "icon.ico" if IS_WINDOWS else assets_src / "icon.png"

    gui_entry = effective_src_dir / "docwen" / "gui_run.py"
    cli_entry = effective_src_dir / "docwen" / "cli_run.py"
    if with_gui and not gui_entry.exists():
        logger.error(f"GUI 入口脚本不存在: {gui_entry}")
        return None
    if with_cli and not cli_entry.exists():
        logger.error(f"CLI 入口脚本不存在: {cli_entry}")
        return None

    common_excludes = [
        "--exclude-module=poplib",
        "--exclude-module=imaplib",
        "--exclude-module=telnetlib",
        "--exclude-module=nntplib",
        "--exclude-module=xmlrpc",
    ]
    if IS_LINUX:
        common_excludes.extend(
            [
                "--exclude-module=win32com",
                "--exclude-module=win32api",
                "--exclude-module=pythoncom",
                "--exclude-module=pywintypes",
            ]
        )

    gui_build_args: list[str] = []
    if with_gui:
        gui_build_args = [
            str(gui_entry),
            "--name=DocWen",
            "--onedir",
            "--clean",
            "--noconfirm",
            *common_excludes,
            "--collect-all=rapidocr_onnxruntime",
            "--collect-data=onnxruntime",
            "--collect-data=latex2mathml",
            "--copy-metadata=PyYAML",
            "--collect-data=ttkbootstrap",
            "--collect-data=docx",
            "--collect-all=pymupdf4llm",
            "--collect-all=pymupdf_layout",
            "--collect-all=easyofd",
            "--collect-submodules=docwen.services.strategies",
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
    ]

    for data_file in data_files:
        if with_gui:
            gui_build_args.append(f"--add-data={data_file}")

    # 4. PyInstaller 构建
    logger.start_step("PyInstaller 构建")
    try:
        if with_gui:
            logger.info("开始 PyInstaller 构建 (GUI)...")
            pyinstaller_main.run(gui_build_args)
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
                *common_excludes,
                "--collect-all=rapidocr_onnxruntime",
                "--collect-data=onnxruntime",
                "--collect-data=latex2mathml",
                "--copy-metadata=PyYAML",
                "--collect-data=ttkbootstrap",
                "--collect-data=docx",
                "--collect-all=pymupdf4llm",
                "--collect-all=pymupdf_layout",
                "--collect-all=easyofd",
                # 策略模块通过 importlib.import_module() 动态加载，PyInstaller 无法静态分析，需显式收集
                "--collect-submodules=docwen.services.strategies",
            ]
            for data_file in data_files:
                cli_build_args.append(f"--add-data={data_file}")
            pyinstaller_main.run(cli_build_args)
            logger.info("PyInstaller CLI 构建成功完成!")
            logger.end_step()

        if effective_src_dir != SRC_DIR:
            with contextlib.suppress(Exception):
                force_remove_directory(BUILD_DIR / "staging_src")

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

        # 国际化文件
        i18n_dest = deploy_dir / "docwen" / "i18n" / "locales"
        if not i18n_dest.exists() and i18n_locales_src.exists():
            i18n_dest.parent.mkdir(parents=True, exist_ok=True)
            copytree_robust(i18n_locales_src, i18n_dest)
            logger.info(f"复制国际化文件到: {i18n_dest}")

        # OCR 模型
        models_dest = deploy_dir / "models"
        if not models_dest.exists() and models_src.exists() and list(models_src.iterdir()):
            copytree_robust(models_src, models_dest)
            logger.info(f"复制 OCR 模型到: {models_dest}")

        # CLI 独立目录资源补充
        # PyInstaller onedir 模式将 --add-data 数据放在 _internal/ 下，
        # 但运行时 get_project_root() 返回 exe 同级目录，需要顶层存在资源。
        # 在此将 templates/configs/i18n 补充复制到 CLI 原始输出目录顶层，
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

            cli_i18n_dest = cli_output_dir / "docwen" / "i18n" / "locales"
            if not cli_i18n_dest.exists() and i18n_locales_src.exists():
                cli_i18n_dest.parent.mkdir(parents=True, exist_ok=True)
                copytree_robust(i18n_locales_src, cli_i18n_dest)
                logger.info(f"补充复制国际化文件到 CLI 目录: {cli_i18n_dest}")

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

        if IS_WINDOWS and with_gui and with_cli and cli_output_dir.exists():
            cli_deploy_dir_name = f"DocWenCLI_v{version}_{PLATFORM_TAG}"
            cli_deploy_dir = DIST_DIR / cli_deploy_dir_name
            if cli_deploy_dir.exists():
                force_remove_directory(cli_deploy_dir)
            shutil.move(str(cli_output_dir), str(cli_deploy_dir))
        logger.end_step()

    except Exception as e:
        logger.error(f"构建失败: {e}")
        import traceback

        logger.debug(traceback.format_exc())
        return None

    return version, deploy_dir


def clean_source_files_from_dist(deploy_dir: Path):
    """
    从部署目录中删除已被 Cython 编译的 .py 源文件

    Args:
        deploy_dir: 部署目录路径
    """
    logger.info("开始清理部署目录中的 Python 源文件...")

    try:
        # 同目录导入 setup_cython 模块
        sys.path.insert(0, str(Path(__file__).parent))
        from setup_cython import CORE_MODULES

        deleted_count = 0
        for module_path in CORE_MODULES:
            try:
                posix_rel = PurePosixPath(module_path).relative_to("src")
            except Exception:
                posix_rel = PurePosixPath(module_path)

            file_to_delete = deploy_dir / Path(*posix_rel.parts)

            if file_to_delete.exists():
                try:
                    file_to_delete.unlink()
                    logger.debug(f"已删除: {posix_rel}")
                    deleted_count += 1
                except Exception as e:
                    logger.warning(f"无法删除 {file_to_delete}: {e}")

        # 删除 .c 文件
        for c_file in deploy_dir.rglob("*.c"):
            # 跳过 Cython 运行时目录
            if "_internal" in str(c_file) and "Cython" in str(c_file):
                continue
            try:
                c_file.unlink()
                logger.debug(f"已删除 C 文件: {c_file.relative_to(deploy_dir)}")
                deleted_count += 1
            except Exception as e:
                logger.warning(f"无法删除 C 文件 {c_file}: {e}")

        logger.info(f"源文件清理完成，共删除 {deleted_count} 个文件")

    except ImportError:
        logger.warning("无法导入 CORE_MODULES 列表，跳过源文件清理")
    except Exception as e:
        logger.error(f"源文件清理过程中发生错误: {e}")


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

    # 初始化日志
    init_logger()

    logger.info("=" * 70)
    logger.info("开始构建 DocWen")
    logger.info("=" * 70)

    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--skip-cython", action="store_true", help="跳过 Cython 编译")
    parser.add_argument("--gui-only", action="store_true", help="仅构建 GUI（不构建 CLI）")
    parser.add_argument("--cli-only", action="store_true", help="仅构建 CLI（不构建 GUI）")
    parser.add_argument("--version", type=str, default=None, help="构建版本号（覆盖源码 __version__，可带 v 前缀）")
    args = parser.parse_args()

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
            # 清理源文件
            logger.start_step("清理源文件")
            if deploy_dir.is_dir():
                clean_source_files_from_dist(deploy_dir)
            logger.end_step()

            # 验证构建
            logger.start_step("构建验证")
            verify_build(deploy_dir, with_cli=with_cli, with_gui=with_gui)
            logger.end_step()

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
