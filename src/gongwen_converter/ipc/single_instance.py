"""
单实例锁管理模块
使用文件锁确保程序只运行一个实例
"""

import os
import sys
import tempfile
import logging
from pathlib import Path

# 根据平台导入不同的文件锁模块
if sys.platform == "win32":
    import msvcrt
else:
    import fcntl

logger = logging.getLogger(__name__)


class SingleInstance:
    """基于文件锁的单实例管理器"""
    
    def __init__(self, app_name="gongwen_converter"):
        """
        初始化单实例管理器
        
        参数:
            app_name: 应用程序名称，用于生成唯一的锁文件和IPC目录
        """
        self.app_name = app_name
        self.lock_file = None
        self.ipc_dir = self._get_ipc_dir()
        self.lock_path = os.path.join(self.ipc_dir, "instance.lock")
        
        # 确保 IPC 目录存在
        os.makedirs(self.ipc_dir, exist_ok=True)
        logger.debug(f"IPC 目录: {self.ipc_dir}")
    
    def _get_ipc_dir(self):
        """
        获取 IPC 通信目录路径
        
        返回:
            str: IPC 目录的完整路径
        """
        temp_dir = tempfile.gettempdir()
        ipc_dir = os.path.join(temp_dir, self.app_name)
        return ipc_dir
    
    def acquire(self):
        """
        尝试获取单实例锁
        
        返回:
            bool: 成功获取锁返回 True，锁已被占用返回 False
        """
        try:
            # 以写模式打开锁文件
            self.lock_file = open(self.lock_path, 'w')
            
            if sys.platform == "win32":
                # Windows 文件锁：使用 msvcrt.locking
                # LK_NBLCK: 非阻塞锁，如果无法锁定则立即返回错误
                msvcrt.locking(self.lock_file.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                # Unix/Linux 文件锁：使用 fcntl.flock
                # LOCK_EX: 排他锁
                # LOCK_NB: 非阻塞模式
                fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            
            # 写入当前进程 ID
            self.lock_file.write(str(os.getpid()))
            self.lock_file.flush()
            
            logger.info(f"成功获取单实例锁: {self.lock_path}")
            return True
            
        except (IOError, OSError) as e:
            logger.info(f"程序已在运行（无法获取锁）: {e}")
            if self.lock_file:
                try:
                    self.lock_file.close()
                except:
                    pass
                self.lock_file = None
            return False
    
    def release(self):
        """释放单实例锁"""
        if self.lock_file:
            try:
                if sys.platform == "win32":
                    # Windows: 解锁文件
                    msvcrt.locking(self.lock_file.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    # Unix/Linux: 解锁文件
                    fcntl.flock(self.lock_file.fileno(), fcntl.LOCK_UN)
                
                self.lock_file.close()
                
                # 删除锁文件
                if os.path.exists(self.lock_path):
                    os.remove(self.lock_path)
                
                logger.info("已释放单实例锁")
            except Exception as e:
                logger.error(f"释放锁失败: {e}")
    
    def get_ipc_dir(self):
        """
        获取 IPC 目录路径
        
        返回:
            str: IPC 目录的完整路径
        """
        return self.ipc_dir
    
    def __enter__(self):
        """上下文管理器入口"""
        return self.acquire()
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.release()


if __name__ == "__main__":
    # 测试代码
    logging.basicConfig(level=logging.DEBUG)
    
    print("测试单实例锁...")
    lock = SingleInstance("test_app")
    
    if lock.acquire():
        print("✓ 成功获取锁")
        print(f"IPC 目录: {lock.get_ipc_dir()}")
        input("按回车键释放锁...")
        lock.release()
        print("✓ 已释放锁")
    else:
        print("✗ 无法获取锁（程序可能已在运行）")
