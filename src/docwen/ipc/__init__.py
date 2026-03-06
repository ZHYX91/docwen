"""
IPC (进程间通信) 模块
使用文件系统实现进程间通信，不依赖网络连接
"""

from .file_ipc import FileIPC
from .single_instance import SingleInstance

__all__ = ["FileIPC", "SingleInstance"]
