"""
Cython 构建脚本

将选定的 Python 模块编译为 C 扩展（.pyd），提高性能和安全性。
编译后的 .pyd 文件默认输出到 build/cython_out，避免污染源码目录。

编译流程：
    1. 验证模块存在性
    2. 检测增量编译（跳过未修改的模块）
    3. 复制 .py → build/cython_work 下的 .pyx
    4. Cython 转译 .pyx → .c
    5. C 编译器编译 .c → .pyd（输出到 build/cython_out）
    6. 自动清理中间文件 (.pyx, .c)

使用方式：
    python setup_cython.py           # 增量编译
    python setup_cython.py compile   # 增量编译
    python setup_cython.py --force   # 强制全量编译
    python setup_cython.py clean     # 清理所有生成文件（.pyd, .pyx, .c）

依赖：
    - Cython: pip install cython
    - setuptools: pip install setuptools
    - C 编译器: Windows 需要 Visual Studio Build Tools
"""

import os
import shutil
import sys
from pathlib import Path

from Cython.Build import cythonize
from Cython.Compiler import Options
from setuptools import Extension, setup

# ==================== 项目路径配置 ====================
# 脚本位于 scripts/build/，项目根目录向上两级
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
SRC_DIR = PROJECT_ROOT / "src"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "build" / "cython_out"
DEFAULT_WORK_DIR = PROJECT_ROOT / "build" / "cython_work"

# ==================== Cython 编译器全局选项 ====================
Options.docstrings = False  # 移除文档字符串（减小体积）
Options.embed_pos_in_docstring = False
Options.buffer_max_dims = 8

# ==================== 编译配置 ====================
COMPILE_OPTIONS = {
    "compiler_directives": {
        "language_level": 3,  # Python 3 语法
        "boundscheck": False,  # 禁用边界检查（提高性能）
        "wraparound": False,  # 禁用负索引包装
        "initializedcheck": False,  # 禁用初始化检查
        "nonecheck": False,  # 禁用 None 检查
        "cdivision": True,  # 使用 C 除法（不抛 ZeroDivision）
        "embedsignature": False,  # 不嵌入函数签名
        "binding": True,  # 保持 Python 绑定兼容性
        "optimize.use_switch": True,  # 优化 switch 语句
        "optimize.unpack_method_calls": True,
    },
    "nthreads": os.cpu_count() or 4,  # 并行编译线程数
}

# ==================== 核心模块列表（约 10 个）====================
# 选择原则：
#   🔴 必须编译：安全相关模块
#   🟡 推荐编译：核心转换逻辑
#   🟢 可选编译：高频调用的工具函数
#   ❌ 不编译：入口文件、GUI、配置、外部库封装（避免杀软误报）

CORE_MODULES = [
    # ==================== 🔴 安全模块（必须编译）====================
    "src/docwen/security/network_isolation.py",
    "src/docwen/security/protection_utils.py",
    # ⚠️ 入口文件（gui_run.py, cli_run.py）不可编译为 .pyd：
    #   - python -m 需要 code object，.pyd 是原生 DLL 无法提供
    #   - 会导致 "No code object available for docwen.gui_run" 错误
    #   - 入口模块非性能瓶颈，安全保护由 PyInstaller 打包实现
    # ==================== 🟡 核心转换器（推荐编译）====================
    "src/docwen/converter/smart_converter.py",
    "src/docwen/converter/md2docx/core.py",
    "src/docwen/converter/docx2md/core.py",
    "src/docwen/converter/md2xlsx/core.py",
    "src/docwen/converter/xlsx2md/core.py",
    # ==================== 🟢 高频工具函数（可选编译）====================
    "src/docwen/utils/text_utils.py",
    "src/docwen/utils/validation_utils.py",
    "src/docwen/utils/number_utils.py",
]


class CythonBuilder:
    """Cython 编译管理器"""

    def __init__(self, project_root: Path, *, output_dir: Path, work_dir: Path):
        """
        初始化编译器

        Args:
            project_root: 项目根目录路径
        """
        self.project_root = project_root
        self.output_dir = output_dir
        self.work_dir = work_dir
        self.valid_modules: list[Path] = []
        self.skipped_modules: list[str] = []
        self._validate_modules()

    def _validate_modules(self) -> None:
        """验证模块存在性，分类有效和无效模块"""
        for mod_str in CORE_MODULES:
            mod_path = self.project_root / mod_str
            if mod_path.exists():
                self.valid_modules.append(mod_path)
            else:
                self.skipped_modules.append(mod_str)

        if self.skipped_modules:
            print(f"⚠️ 以下 {len(self.skipped_modules)} 个模块不存在，已跳过:")
            for mod in self.skipped_modules:
                print(f"   - {mod}")

        print(f"✅ 有效模块: {len(self.valid_modules)} 个")

    def _compiled_files(self, py_path: Path) -> list[Path]:
        stem = py_path.stem
        rel_path = py_path.relative_to(self.project_root / "src")
        out_parent = self.output_dir / rel_path.parent
        return list(out_parent.glob(f"{stem}*.pyd")) + list(out_parent.glob(f"{stem}*.so"))

    def _needs_compile(self, py_path: Path) -> bool:
        """
        检查模块是否需要重新编译（增量编译检测）

        比较 .py 文件和对应 .pyd 文件的修改时间

        Args:
            py_path: Python 源文件路径

        Returns:
            True 如果需要编译，False 如果可以跳过
        """
        compiled_files = self._compiled_files(py_path)

        if not compiled_files:
            return True

        latest_mtime = max(f.stat().st_mtime for f in compiled_files)
        py_mtime = py_path.stat().st_mtime

        return py_mtime > latest_mtime

    def _create_extension(self, py_path: Path, pyx_path: Path) -> Extension:
        """
        为单个模块创建 Cython Extension 对象

        Args:
            py_path: 原始 .py 文件路径
            pyx_path: 复制后的 .pyx 文件路径

        Returns:
            配置好的 Extension 对象
        """
        # 计算模块名（相对于 src/ 目录）
        # 例如: src/docwen/utils/text_utils.py → docwen.utils.text_utils
        try:
            rel_path = py_path.relative_to(self.project_root / "src")
            module_name = str(rel_path.with_suffix("")).replace(os.sep, ".")
        except ValueError:
            # 如果不在 src/ 下（如 src/gui_run.py 直接在 src/ 根目录）
            module_name = py_path.stem

        # 平台特定的编译参数
        if os.name == "nt":  # Windows
            extra_compile_args = ["/O2"]  # MSVC 优化级别
            extra_link_args = []
        else:  # Linux/macOS
            extra_compile_args = ["-O3"]
            if os.environ.get("DOCWEN_MARCH_NATIVE", "").strip() == "1":
                extra_compile_args.append("-march=native")
            extra_link_args = []

        return Extension(
            name=module_name,
            sources=[str(pyx_path)],
            language="c",
            extra_compile_args=extra_compile_args,
            extra_link_args=extra_link_args,
        )

    def _copy_py_to_pyx(self, modules: list[Path]) -> list[tuple[Path, Path]]:
        """
        将 Python 文件复制为 .pyx 文件

        Args:
            modules: 需要编译的模块路径列表

        Returns:
            (py_path, pyx_path) 元组列表
        """
        pairs = []
        for py_path in modules:
            rel_path = py_path.relative_to(self.project_root / "src")
            pyx_path = (self.work_dir / rel_path).with_suffix(".pyx")
            pyx_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(py_path, pyx_path)
            pairs.append((py_path, pyx_path))
            print(f"📄 复制: {py_path.name} → {pyx_path.name}")
        return pairs

    def _cleanup_intermediate(self, pyx_paths: list[Path]) -> None:
        """
        清理中间文件（.pyx 和 .c）

        Args:
            pyx_paths: 编译的 .pyx 文件路径列表
        """
        cleaned = 0
        for pyx_path in pyx_paths:
            if pyx_path.exists():
                pyx_path.unlink()
                cleaned += 1
            c_path = pyx_path.with_suffix(".c")
            if c_path.exists():
                c_path.unlink()
                cleaned += 1

        if cleaned:
            print(f"🗑️ 已清理 {cleaned} 个中间文件")

    def compile(self, force: bool = False) -> bool:
        """
        执行 Cython 编译

        Args:
            force: 是否强制全量编译（忽略增量检测）

        Returns:
            True 编译成功，False 编译失败
        """
        print("\n" + "=" * 60)
        print("Cython 编译开始")
        print("=" * 60)

        if not self.valid_modules:
            print("❌ 没有有效的模块可编译")
            return False

        # 1. 筛选需要编译的模块
        if force:
            to_compile = self.valid_modules
            print(f"🔄 强制模式: 编译全部 {len(to_compile)} 个模块")
        else:
            to_compile = [m for m in self.valid_modules if self._needs_compile(m)]
            skipped = len(self.valid_modules) - len(to_compile)
            if skipped:
                print(f"⏭️ 增量编译: 跳过 {skipped} 个未修改模块")

        if not to_compile:
            print("✅ 所有模块已是最新，无需编译")
            return True

        print(f"📦 待编译模块: {len(to_compile)} 个\n")

        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            self.work_dir.mkdir(parents=True, exist_ok=True)

            # 2. 复制 .py → .pyx
            module_pairs = self._copy_py_to_pyx(to_compile)

            # 3. 创建 Extension 列表
            extensions = [self._create_extension(py_path, pyx_path) for py_path, pyx_path in module_pairs]

            print(f"\n🔨 开始编译（使用 {COMPILE_OPTIONS['nthreads']} 线程）...\n")

            # 4. Cython 转译和编译
            cythonized = cythonize(extensions, **COMPILE_OPTIONS)

            # 5. 构建扩展（输出到 output_dir）
            setup(
                name="docwen_cython",
                ext_modules=cythonized,
                script_args=[
                    "build_ext",
                    f"--build-lib={self.output_dir}",
                    f"--build-temp={self.work_dir / 'temp'}",
                ],
                zip_safe=False,
            )

            # 6. 清理中间文件
            print()
            self._cleanup_intermediate([pyx_path for _py_path, pyx_path in module_pairs])

            # 7. 验证编译结果
            self._verify_results(to_compile)

            print("\n✅ Cython 编译成功完成!")
            return True

        except Exception as e:
            print(f"\n❌ Cython 编译失败: {e}")
            # 尝试清理中间文件
            self._cleanup_intermediate(
                [
                    (self.work_dir / py_path.relative_to(self.project_root / "src")).with_suffix(".pyx")
                    for py_path in to_compile
                ]
            )
            return False

    def _verify_results(self, modules: list[Path]) -> None:
        """
        验证编译结果，打印生成的 .pyd 文件

        Args:
            modules: 编译的模块列表
        """
        print("\n" + "=" * 60)
        print("编译结果验证")
        print("=" * 60)

        success = 0
        failed = 0

        for py_path in modules:
            compiled_files = self._compiled_files(py_path)

            if compiled_files:
                for compiled in compiled_files:
                    rel_path = compiled.relative_to(self.project_root)
                    print(f"✓ {rel_path}")
                success += 1
            else:
                print(f"✗ 未生成: {py_path.stem}.*")
                failed += 1

        print(f"\n成功: {success}, 失败: {failed}")

    def clean(self) -> None:
        """清理所有编译生成的文件（.pyd, .so, .pyx, .c）"""
        print("\n" + "=" * 60)
        print("清理编译文件")
        print("=" * 60)

        cleaned = {"compiled": 0, "pyx": 0, "c": 0}

        for py_path in self.valid_modules:
            for compiled in self._compiled_files(py_path):
                compiled.unlink()
                print(f"🗑️ 删除: {compiled.relative_to(self.project_root)}")
                cleaned["compiled"] += 1

        if self.work_dir.exists():
            shutil.rmtree(self.work_dir, ignore_errors=True)
            print(f"🗑️ 删除目录: {self.work_dir.relative_to(self.project_root)}")

        if self.output_dir.exists():
            for root, _dirs, files in os.walk(self.output_dir):
                for f in files:
                    if f.endswith((".pyd", ".so")):
                        Path(root, f).unlink()
                        cleaned["compiled"] += 1

            try:
                shutil.rmtree(self.output_dir)
                print(f"🗑️ 删除目录: {self.output_dir.relative_to(self.project_root)}")
            except Exception:
                pass

        total = sum(cleaned.values())
        print(f"\n✅ 清理完成: 共删除 {total} 个文件")


def main():
    """主函数"""
    print("=" * 60)
    print("DocWen - Cython 编译工具")
    print("=" * 60)
    print(f"项目根目录: {PROJECT_ROOT}")
    print(f"待编译模块: {len(CORE_MODULES)} 个")

    output_dir = DEFAULT_OUTPUT_DIR
    work_dir = DEFAULT_WORK_DIR

    args = sys.argv[1:]

    def _extract_flag_value(flag: str) -> str | None:
        for i, a in enumerate(args):
            if a == flag and i + 1 < len(args):
                return args[i + 1]
        return None

    out_arg = _extract_flag_value("--output-dir")
    if out_arg:
        output_dir = Path(out_arg)
        if not output_dir.is_absolute():
            output_dir = (PROJECT_ROOT / output_dir).resolve()

    work_arg = _extract_flag_value("--work-dir")
    if work_arg:
        work_dir = Path(work_arg)
        if not work_dir.is_absolute():
            work_dir = (PROJECT_ROOT / work_dir).resolve()

    builder = CythonBuilder(PROJECT_ROOT, output_dir=output_dir, work_dir=work_dir)

    # 解析命令行参数

    if not args or args[0] == "compile":
        # 增量编译
        force = "--force" in args
        success = builder.compile(force=force)
        sys.exit(0 if success else 1)

    elif args[0] == "--force":
        # 强制全量编译
        success = builder.compile(force=True)
        sys.exit(0 if success else 1)

    elif args[0] == "clean":
        # 清理
        builder.clean()
        sys.exit(0)

    else:
        print("""
用法:
    python setup_cython.py              # 增量编译
    python setup_cython.py compile      # 增量编译
    python setup_cython.py --force      # 强制全量编译
    python setup_cython.py clean        # 清理所有生成文件
    python setup_cython.py compile --output-dir <dir> [--work-dir <dir>]
        """)
        sys.exit(1)


if __name__ == "__main__":
    main()
