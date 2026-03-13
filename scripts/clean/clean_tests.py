"""
清理 tests/ 目录下的测试运行产物

清理项：
- tests/**/__pycache__/ 目录
- tests/**/*.pyc 文件（包含残留的历史产物）
- tests/.pytest_cache/ 目录（如存在）
- tests/temp/ 与 tests/output/ 目录（如存在）
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from clean_utils import safe_remove_directory, safe_remove_file


def clean_tests_pycache(project_root):
    print("\n" + "=" * 50)
    print("步骤 1: 清理 tests/**/__pycache__")
    print("=" * 50)

    tests_root = os.path.join(project_root, "tests")
    if not os.path.isdir(tests_root):
        print("未找到 tests 目录")
        return 0

    deleted_count = 0
    for root, dirs, files in os.walk(tests_root):
        if "__pycache__" in dirs:
            pycache_path = os.path.join(root, "__pycache__")
            if safe_remove_directory(pycache_path):
                deleted_count += 1

    print(f"总计删除 {deleted_count} 个 tests/**/__pycache__ 目录")
    return deleted_count


def clean_tests_pyc_files(project_root):
    print("\n" + "=" * 50)
    print("步骤 2: 清理 tests/**/*.pyc")
    print("=" * 50)

    tests_root = os.path.join(project_root, "tests")
    if not os.path.isdir(tests_root):
        print("未找到 tests 目录")
        return 0

    deleted_count = 0
    for root, dirs, files in os.walk(tests_root):
        for file in files:
            if file.endswith(".pyc"):
                if safe_remove_file(os.path.join(root, file)):
                    deleted_count += 1

    print(f"总计删除 {deleted_count} 个 tests/**/*.pyc 文件")
    return deleted_count


def clean_tests_pytest_cache(project_root):
    print("\n" + "=" * 50)
    print("步骤 3: 清理 tests/.pytest_cache")
    print("=" * 50)

    tests_root = os.path.join(project_root, "tests")
    cache_path = os.path.join(tests_root, ".pytest_cache")
    if not os.path.isdir(cache_path):
        print("未找到 tests/.pytest_cache 目录")
        return 0

    if safe_remove_directory(cache_path):
        return 1
    return 0


def clean_tests_temp_dirs(project_root):
    print("\n" + "=" * 50)
    print("步骤 4: 清理 tests/temp 与 tests/output")
    print("=" * 50)

    deleted_count = 0
    for rel in ("tests/temp", "tests/output"):
        path = os.path.join(project_root, *rel.split("/"))
        if os.path.isdir(path):
            if safe_remove_directory(path):
                deleted_count += 1

    if deleted_count == 0:
        print("未找到 tests/temp 或 tests/output 目录")
    return deleted_count


def main():
    print("=" * 60)
    print("DocWen - tests 产物清理工具")
    print("=" * 60)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(script_dir, "..", ".."))

    print(f"项目根目录: {project_root}")
    print("开始清理 tests 运行产物...")

    total_deleted = 0
    total_deleted += clean_tests_pycache(project_root)
    total_deleted += clean_tests_pyc_files(project_root)
    total_deleted += clean_tests_pytest_cache(project_root)
    total_deleted += clean_tests_temp_dirs(project_root)

    print("\n" + "=" * 60)
    print("tests 产物清理完成!")
    print(f"总计删除项目: {total_deleted}")
    print("=" * 60)


if __name__ == "__main__":
    main()
