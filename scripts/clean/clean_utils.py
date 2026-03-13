"""
清理脚本的共享工具函数

提供安全的文件/目录删除操作，自动处理只读权限和重试逻辑。

依赖：
    - 标准库：os, shutil, stat, sys, time, pathlib
"""

from __future__ import annotations

import os
import shutil
import stat
import sys
import time
from pathlib import Path


def get_repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _ensure_within_repo(target: Path) -> None:
    repo_root = get_repo_root()
    try:
        target.resolve().relative_to(repo_root)
    except Exception as e:
        raise ValueError(f"refuse_to_delete_outside_repo: {target}") from e


def remove_readonly(func, path, exc_info) -> None:
    try:
        os.chmod(path, stat.S_IWRITE)
    except Exception:
        pass
    func(path)


def _remove_readonly_onexc(func, path, exc_info) -> None:
    try:
        os.chmod(path, stat.S_IWRITE)
    except Exception:
        pass
    func(path)


def safe_remove_file(file_path: str | Path) -> bool:
    target = Path(file_path)
    if not target.exists():
        return True
    _ensure_within_repo(target)
    try:
        target.chmod(stat.S_IWRITE)
    except Exception:
        pass
    try:
        target.unlink()
        print(f"[OK] 删除文件: {target}")
        return True
    except Exception as e:
        print(f"[FAIL] 删除文件失败: {target} ({e})")
        return False


def safe_remove_directory(dir_path: str | Path) -> bool:
    target = Path(dir_path)
    if not target.exists():
        return True
    _ensure_within_repo(target)

    max_attempts = 3
    for attempt in range(1, max_attempts + 1):
        try:
            if sys.version_info >= (3, 12):
                shutil.rmtree(str(target), onexc=_remove_readonly_onexc)
            else:
                shutil.rmtree(str(target), onerror=remove_readonly)
            print(f"[OK] 删除目录: {target}")
            return True
        except PermissionError as e:
            if attempt >= max_attempts:
                print(f"[FAIL] 删除目录失败: {target} ({e})")
                return False
            time.sleep(0.2 * attempt)
        except Exception as e:
            print(f"[FAIL] 删除目录失败: {target} ({e})")
            return False


def safe_remove_glob(pattern: str, *, base_dir: str | Path) -> tuple[int, int]:
    base = Path(base_dir)
    _ensure_within_repo(base)

    ok = 0
    fail = 0
    for path in base.glob(pattern):
        if path.is_dir():
            if safe_remove_directory(path):
                ok += 1
            else:
                fail += 1
        else:
            if safe_remove_file(path):
                ok += 1
            else:
                fail += 1
    return ok, fail
