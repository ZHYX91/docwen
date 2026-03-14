"""
清理运行产生文件

按顺序清理：__pycache__文件夹 -> logs文件夹 -> pytest产物 -> lint缓存 -> 散落临时文件 -> 生成的输出目录
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import glob

from clean_utils import safe_remove_directory, safe_remove_file

_SKIP_DIR_NAMES = {
    ".git",
    ".venv",
    ".workspace-links",
    "build",
    "dist",
    "env",
    "node_modules",
    "venv",
}


def clean_pycache_dirs(project_root):
    """清理所有__pycache__文件夹"""
    print("\n" + "=" * 50)
    print("步骤 1: 清理 __pycache__ 文件夹")
    print("=" * 50)

    pycache_dirs = []
    for root, dirs, _files in os.walk(project_root):
        dirs[:] = [d for d in dirs if d not in _SKIP_DIR_NAMES]
        if "__pycache__" in dirs:
            pycache_path = os.path.join(root, "__pycache__")
            pycache_dirs.append(pycache_path)

    if not pycache_dirs:
        print("未找到 __pycache__ 文件夹")
        return 0

    deleted_count = 0
    for pycache_dir in pycache_dirs:
        if safe_remove_directory(pycache_dir):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个 __pycache__ 文件夹")
    return deleted_count


def clean_logs_directory(project_root):
    """清理logs文件夹（包括项目根目录和src/docwen下的日志目录）"""
    print("\n" + "=" * 50)
    print("步骤 2: 清理 logs 文件夹")
    print("=" * 50)

    # 需要清理的日志目录列表
    logs_paths = [
        os.path.join(project_root, "logs"),
        os.path.join(project_root, "src", "docwen", "logs"),
    ]

    deleted_count = 0
    found_any = False

    for logs_path in logs_paths:
        if not os.path.exists(logs_path):
            continue

        found_any = True
        if safe_remove_directory(logs_path):
            print(f"[OK] 成功删除 logs 文件夹: {logs_path}")
            deleted_count += 1
        else:
            print(f"[FAIL] 删除 logs 文件夹失败: {logs_path}")

    if not found_any:
        print("未找到 logs 文件夹")

    return deleted_count


def clean_pytest_artifacts(project_root):
    """
    清理 pytest 产生的文件和目录

    清理项：
    - .pytest_cache/ 目录（递归查找所有）
    - .coverage 文件
    - coverage.xml 文件
    - htmlcov/ 目录
    - tests/temp/ 目录
    - tests/output/ 目录
    """
    print("\n" + "=" * 50)
    print("步骤 3: 清理 pytest 产物")
    print("=" * 50)

    deleted_count = 0

    # 清理 .pytest_cache 目录（可能存在于多个位置）
    pytest_cache_dirs = []
    for root, dirs, _files in os.walk(project_root):
        dirs[:] = [d for d in dirs if d not in _SKIP_DIR_NAMES]
        if ".pytest_cache" in dirs:
            cache_path = os.path.join(root, ".pytest_cache")
            pytest_cache_dirs.append(cache_path)

    for cache_dir in pytest_cache_dirs:
        if safe_remove_directory(cache_dir):
            print(f"[OK] 删除 .pytest_cache: {cache_dir}")
            deleted_count += 1

    if not pytest_cache_dirs:
        print("未找到 .pytest_cache 目录")

    # 清理根目录下的覆盖率文件
    coverage_files = [
        os.path.join(project_root, ".coverage"),
        os.path.join(project_root, "coverage.xml"),
    ]

    for coverage_file in coverage_files:
        if os.path.isfile(coverage_file) and safe_remove_file(coverage_file):
            deleted_count += 1

    # 清理 htmlcov 目录
    htmlcov_path = os.path.join(project_root, "htmlcov")
    if os.path.isdir(htmlcov_path):
        if safe_remove_directory(htmlcov_path):
            print(f"[OK] 删除 htmlcov 目录: {htmlcov_path}")
            deleted_count += 1
        else:
            print(f"[FAIL] 删除 htmlcov 目录失败: {htmlcov_path}")

    # 清理 tests/temp 和 tests/output 目录
    test_temp_dirs = [
        os.path.join(project_root, "tests", "temp"),
        os.path.join(project_root, "tests", "output"),
    ]

    for temp_dir in test_temp_dirs:
        if os.path.isdir(temp_dir):
            if safe_remove_directory(temp_dir):
                print(f"[OK] 删除测试临时目录: {temp_dir}")
                deleted_count += 1
            else:
                print(f"[FAIL] 删除测试临时目录失败: {temp_dir}")

    print(f"总计删除 {deleted_count} 个 pytest 产物")
    return deleted_count


def clean_lint_cache(project_root):
    """清理 lint 工具缓存目录（如 .ruff_cache）"""
    print("\n" + "=" * 50)
    print("步骤 4: 清理 lint 工具缓存")
    print("=" * 50)

    cache_dirs = [
        os.path.join(project_root, ".ruff_cache"),
    ]

    deleted_count = 0
    found_any = False

    for cache_dir in cache_dirs:
        if not os.path.isdir(cache_dir):
            continue

        found_any = True
        if safe_remove_directory(cache_dir):
            print(f"[OK] 删除缓存目录: {cache_dir}")
            deleted_count += 1
        else:
            print(f"[FAIL] 删除缓存目录失败: {cache_dir}")

    if not found_any:
        print("未找到 lint 缓存目录")

    return deleted_count


def clean_stray_temp_files(project_root):
    """
    清理项目根目录下散落的临时文件

    清理项：
    - *.log / *.log.* 日志文件
    - *.out 输出文件
    - *.tmp 临时文件
    """
    print("\n" + "=" * 50)
    print("步骤 5: 清理散落的临时文件")
    print("=" * 50)

    patterns = ["*.log", "*.log.*", "*.out", "*.tmp"]
    target_files = []
    for pattern in patterns:
        target_files.extend(glob.glob(os.path.join(project_root, pattern)))

    if not target_files:
        print("未找到散落的临时文件")
        return 0

    deleted_count = 0
    for f in target_files:
        if safe_remove_file(f):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个临时文件")
    return deleted_count


def clean_generated_output(project_root):
    """
    清理生成的输出目录和文件

    清理项：
    - samples/generated_templates/ 目录
    - samples/style_injection_output/ 目录
    - templates/空白模板.docx 文件（脚本生成的空白模板）
    """
    print("\n" + "=" * 50)
    print("步骤 6: 清理生成的输出目录和文件")
    print("=" * 50)

    deleted_count = 0

    # 清理生成的输出目录
    output_dirs = [
        os.path.join(project_root, "samples", "generated_templates"),
        os.path.join(project_root, "samples", "style_injection_output"),
    ]

    for output_dir in output_dirs:
        if os.path.isdir(output_dir):
            if safe_remove_directory(output_dir):
                print(f"[OK] 删除输出目录: {output_dir}")
                deleted_count += 1
            else:
                print(f"[FAIL] 删除输出目录失败: {output_dir}")

    # 清理生成的空白模板文件
    generated_files = [
        os.path.join(project_root, "templates", "空白模板.docx"),
    ]

    for gen_file in generated_files:
        if os.path.isfile(gen_file) and safe_remove_file(gen_file):
            deleted_count += 1

    if deleted_count == 0:
        print("未找到生成的输出目录或文件")

    return deleted_count


def main():
    """主清理函数"""
    print("=" * 60)
    print("DocWen - 运行产生文件清理工具")
    print("=" * 60)

    # 获取项目根目录（当前脚本在scripts/clean目录，需要向上两级）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(script_dir, "..", ".."))

    print(f"项目根目录: {project_root}")
    print("开始清理运行产生文件...")

    # 按顺序执行清理
    total_deleted = 0

    # 1. 清理 __pycache__ 文件夹
    pycache_count = clean_pycache_dirs(project_root)
    total_deleted += pycache_count

    # 2. 清理 logs 文件夹
    logs_count = clean_logs_directory(project_root)
    total_deleted += logs_count

    # 3. 清理 pytest 产物
    pytest_count = clean_pytest_artifacts(project_root)
    total_deleted += pytest_count

    # 4. 清理 lint 工具缓存
    lint_count = clean_lint_cache(project_root)
    total_deleted += lint_count

    # 5. 清理散落的临时文件
    temp_count = clean_stray_temp_files(project_root)
    total_deleted += temp_count

    # 6. 清理生成的输出目录和文件
    output_count = clean_generated_output(project_root)
    total_deleted += output_count

    print("\n" + "=" * 60)
    print("运行产生文件清理完成!")
    print(f"总计删除项目: {total_deleted}")
    print("=" * 60)


if __name__ == "__main__":
    main()
