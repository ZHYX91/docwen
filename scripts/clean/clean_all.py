"""
执行所有清理任务

按顺序调用各清理子脚本：构建产物 → Cython 编译文件 → 运行时缓存 → 测试产物。

使用方式：
    python scripts/clean/clean_all.py
"""

import subprocess
import sys
import os


def run_script(script_name):
    """运行指定的清理脚本"""
    script_path = os.path.join(os.path.dirname(__file__), script_name)
    print("\n" + "=" * 60)
    print(f"正在运行脚本: {script_name}")
    print("=" * 60)

    try:
        result = subprocess.run([sys.executable, script_path], check=True, text=True, capture_output=True)
        print(result.stdout)
        if result.stderr:
            print("--- STDERR ---")
            print(result.stderr)
    except subprocess.CalledProcessError as e:
        print(f"运行脚本 {script_name} 失败:")
        print(e.stdout)
        print(e.stderr)


def main():
    """主函数"""
    print("=" * 70)
    print("DocWen - 全面清理工具")
    print("=" * 70)

    # 按顺序运行所有清理脚本
    run_script("clean_build.py")
    run_script("clean_cython.py")
    run_script("clean_runtime.py")
    run_script("clean_tests.py")

    print("\n" + "=" * 70)
    print("所有清理任务已完成!")
    print("=" * 70)


if __name__ == "__main__":
    main()
