"""
清理 Cython 编译文件

按顺序清理：.c文件 -> .pyd文件 -> .so文件 -> .pyx文件
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from clean_utils import safe_remove_file


def clean_c_files(project_root):
    """清理所有.c文件"""
    print("\n" + "=" * 50)
    print("步骤 1: 清理 .c 文件")
    print("=" * 50)

    c_files = []
    for root, dirs, files in os.walk(project_root):
        for file in files:
            if file.endswith(".c"):
                c_files.append(os.path.join(root, file))

    if not c_files:
        print("未找到 .c 文件")
        return 0

    deleted_count = 0
    for c_file in c_files:
        if safe_remove_file(c_file):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个 .c 文件")
    return deleted_count


def clean_pyd_files(project_root):
    """清理所有.pyd文件"""
    print("\n" + "=" * 50)
    print("步骤 2: 清理 .pyd 文件")
    print("=" * 50)

    pyd_files = []
    for root, dirs, files in os.walk(project_root):
        for file in files:
            if file.endswith(".pyd"):
                pyd_files.append(os.path.join(root, file))

    if not pyd_files:
        print("未找到 .pyd 文件")
        return 0

    deleted_count = 0
    for pyd_file in pyd_files:
        if safe_remove_file(pyd_file):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个 .pyd 文件")
    return deleted_count


def clean_so_files(project_root):
    """清理所有.so文件（Linux Cython 编译产物）"""
    print("\n" + "=" * 50)
    print("步骤 3: 清理 .so 文件")
    print("=" * 50)

    so_files = []
    for root, dirs, files in os.walk(project_root):
        for file in files:
            if file.endswith(".so"):
                so_files.append(os.path.join(root, file))

    if not so_files:
        print("未找到 .so 文件")
        return 0

    deleted_count = 0
    for so_file in so_files:
        if safe_remove_file(so_file):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个 .so 文件")
    return deleted_count


def clean_pyx_files(project_root):
    """清理所有.pyx文件"""
    print("\n" + "=" * 50)
    print("步骤 4: 清理 .pyx 文件")
    print("=" * 50)

    pyx_files = []
    for root, dirs, files in os.walk(project_root):
        for file in files:
            if file.endswith(".pyx"):
                pyx_files.append(os.path.join(root, file))

    if not pyx_files:
        print("未找到 .pyx 文件")
        return 0

    deleted_count = 0
    for pyx_file in pyx_files:
        if safe_remove_file(pyx_file):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个 .pyx 文件")
    return deleted_count


def main():
    """主清理函数"""
    print("=" * 60)
    print("DocWen - Cython 编译文件清理工具")
    print("=" * 60)

    # 获取项目根目录（当前脚本在scripts/clean目录，需要向上两级）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(script_dir, "..", ".."))

    print(f"项目根目录: {project_root}")
    print("开始清理 Cython 编译文件...")

    # 按顺序执行清理
    total_deleted = 0

    # 1. 清理 .c 文件
    c_count = clean_c_files(project_root)
    total_deleted += c_count

    # 2. 清理 .pyd 文件
    pyd_count = clean_pyd_files(project_root)
    total_deleted += pyd_count

    # 3. 清理 .so 文件
    so_count = clean_so_files(project_root)
    total_deleted += so_count

    # 4. 清理 .pyx 文件
    pyx_count = clean_pyx_files(project_root)
    total_deleted += pyx_count

    print("\n" + "=" * 60)
    print("Cython 编译文件清理完成!")
    print(f"总计删除项目: {total_deleted}")
    print("=" * 60)


if __name__ == "__main__":
    main()
