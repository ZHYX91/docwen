"""
清理构建和打包过程产生的文件

按顺序清理：build/ 文件夹 -> dist/ 文件夹 -> .spec 文件 -> *.egg-info/ -> *.egg / *.whl
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import glob

from clean_utils import safe_remove_directory, safe_remove_file


def clean_build_directory(project_root):
    """清理 build/ 文件夹"""
    print("\n" + "=" * 50)
    print("步骤 1: 清理 build 文件夹")
    print("=" * 50)

    build_path = os.path.join(project_root, "build")

    if not os.path.exists(build_path):
        print("未找到 build 文件夹")
        return 0

    if safe_remove_directory(build_path):
        print(f"[OK] 成功删除 build 文件夹: {build_path}")
        return 1
    else:
        print(f"[FAIL] 删除 build 文件夹失败: {build_path}")
        return 0


def clean_dist_directory(project_root):
    """清理 dist/ 文件夹"""
    print("\n" + "=" * 50)
    print("步骤 2: 清理 dist 文件夹")
    print("=" * 50)

    dist_path = os.path.join(project_root, "dist")

    if not os.path.exists(dist_path):
        print("未找到 dist 文件夹")
        return 0

    if safe_remove_directory(dist_path):
        print(f"[OK] 成功删除 dist 文件夹: {dist_path}")
        return 1
    else:
        print(f"[FAIL] 删除 dist 文件夹失败: {dist_path}")
        return 0


def clean_spec_files(project_root):
    """清理 .spec 文件"""
    print("\n" + "=" * 50)
    print("步骤 3: 清理 .spec 文件")
    print("=" * 50)

    spec_files = glob.glob(os.path.join(project_root, "*.spec"))

    if not spec_files:
        print("未找到 .spec 文件")
        return 0

    deleted_count = 0
    for spec_file in spec_files:
        if safe_remove_file(spec_file):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个 .spec 文件")
    return deleted_count


def clean_egg_info_dirs(project_root):
    """清理 *.egg-info/ 目录（setuptools 生成的元数据）"""
    print("\n" + "=" * 50)
    print("步骤 4: 清理 *.egg-info 目录")
    print("=" * 50)

    egg_info_dirs = []
    for root, dirs, files in os.walk(project_root):
        for d in dirs:
            if d.endswith(".egg-info"):
                egg_info_dirs.append(os.path.join(root, d))

    if not egg_info_dirs:
        print("未找到 *.egg-info 目录")
        return 0

    deleted_count = 0
    for egg_dir in egg_info_dirs:
        if safe_remove_directory(egg_dir):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个 *.egg-info 目录")
    return deleted_count


def clean_package_files(project_root):
    """清理 *.egg 和 *.whl 打包产物文件"""
    print("\n" + "=" * 50)
    print("步骤 5: 清理 *.egg / *.whl 文件")
    print("=" * 50)

    patterns = ["*.egg", "*.whl"]
    target_files = []
    for pattern in patterns:
        target_files.extend(glob.glob(os.path.join(project_root, pattern)))
        target_files.extend(glob.glob(os.path.join(project_root, "dist", pattern)))

    if not target_files:
        print("未找到 *.egg / *.whl 文件")
        return 0

    deleted_count = 0
    for f in target_files:
        if safe_remove_file(f):
            deleted_count += 1

    print(f"总计删除 {deleted_count} 个打包产物文件")
    return deleted_count


def main():
    """主清理函数"""
    print("=" * 60)
    print("DocWen - 构建产物清理工具")
    print("=" * 60)

    # 获取项目根目录（当前脚本在scripts/clean目录，需要向上两级）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(script_dir, "..", ".."))

    print(f"项目根目录: {project_root}")
    print("开始清理构建产物...")

    # 按顺序执行清理
    total_deleted = 0

    # 1. 清理 build/ 文件夹
    build_count = clean_build_directory(project_root)
    total_deleted += build_count

    # 2. 清理 dist/ 文件夹
    dist_count = clean_dist_directory(project_root)
    total_deleted += dist_count

    # 3. 清理 .spec 文件
    spec_count = clean_spec_files(project_root)
    total_deleted += spec_count

    # 4. 清理 *.egg-info 目录
    egg_count = clean_egg_info_dirs(project_root)
    total_deleted += egg_count

    # 5. 清理 *.egg / *.whl 文件
    pkg_count = clean_package_files(project_root)
    total_deleted += pkg_count

    print("\n" + "=" * 60)
    print("构建产物清理完成!")
    print(f"总计删除项目: {total_deleted}")
    print("=" * 60)


if __name__ == "__main__":
    main()
