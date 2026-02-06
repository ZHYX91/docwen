"""
基于文件系统的 IPC 通信模块
使用文件监控实现进程间通信，不依赖网络连接
"""

import os
import json
import time
import logging
import threading
from pathlib import Path
from typing import Callable, Optional, Dict, Any

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except ImportError:
    raise ImportError("需要安装 watchdog 库: pip install watchdog")

logger = logging.getLogger(__name__)


class CommandHandler(FileSystemEventHandler):
    """命令文件监控处理器"""
    
    def __init__(self, callback: Callable[[Dict[str, Any]], None]):
        """
        初始化命令处理器
        
        参数:
            callback: 接收到命令时的回调函数
        """
        super().__init__()
        self.callback = callback
        self.processed_files = set()
    
    def on_created(self, event):
        """当新文件创建时触发"""
        if event.is_directory:
            return
        
        if not event.src_path.endswith('.json'):
            return
        
        # 避免重复处理
        if event.src_path in self.processed_files:
            return
        
        # 等待文件写入完成
        time.sleep(0.1)
        
        try:
            # 读取命令文件
            with open(event.src_path, 'r', encoding='utf-8') as f:
                command = json.load(f)
            
            logger.info(f"接收到命令: {command.get('action', 'unknown')}")
            
            # 执行回调
            self.callback(command)
            
            # 标记为已处理
            self.processed_files.add(event.src_path)
            
            # 删除命令文件
            try:
                os.remove(event.src_path)
                logger.debug(f"已删除命令文件: {os.path.basename(event.src_path)}")
            except Exception as e:
                logger.warning(f"删除命令文件失败: {e}")
                
        except Exception as e:
            logger.error(f"处理命令文件失败: {e}")


class FileIPC:
    """基于文件系统的 IPC 通信"""
    
    def __init__(self, ipc_dir: str, callback: Callable[[Dict[str, Any]], None]):
        """
        初始化文件 IPC
        
        参数:
            ipc_dir: IPC 通信目录路径
            callback: 接收到命令时的回调函数
        """
        self.ipc_dir = ipc_dir
        self.commands_dir = os.path.join(ipc_dir, "commands")
        self.status_file = os.path.join(ipc_dir, "status.json")
        self.callback = callback
        
        # 创建命令目录
        os.makedirs(self.commands_dir, exist_ok=True)
        
        # 文件监控
        self.observer = None
        self.handler = None
    
    def start(self):
        """
        启动 IPC 监听
        
        返回:
            bool: 启动成功返回 True，失败返回 False
        """
        try:
            # 清理旧的命令文件
            self._cleanup_old_commands()
            
            # 创建文件监控
            self.handler = CommandHandler(self.callback)
            self.observer = Observer()
            self.observer.schedule(self.handler, self.commands_dir, recursive=False)
            self.observer.start()
            
            # 写入状态文件
            self._update_status({'running': True, 'pid': os.getpid()})
            
            logger.info(f"文件 IPC 已启动，监控目录: {self.commands_dir}")
            return True
            
        except Exception as e:
            logger.error(f"启动文件 IPC 失败: {e}")
            return False
    
    def stop(self):
        """停止 IPC 监听"""
        if self.observer:
            try:
                self.observer.stop()
                self.observer.join(timeout=2)
                logger.info("文件 IPC 已停止")
            except Exception as e:
                logger.error(f"停止文件 IPC 失败: {e}")
        
        # 删除状态文件
        try:
            if os.path.exists(self.status_file):
                os.remove(self.status_file)
                logger.debug("已删除状态文件")
        except Exception as e:
            logger.warning(f"删除状态文件失败: {e}")
    
    def _cleanup_old_commands(self):
        """清理旧的命令文件"""
        try:
            if not os.path.exists(self.commands_dir):
                return
            
            for filename in os.listdir(self.commands_dir):
                if filename.endswith('.json'):
                    file_path = os.path.join(self.commands_dir, filename)
                    try:
                        os.remove(file_path)
                        logger.debug(f"清理旧命令文件: {filename}")
                    except Exception as e:
                        logger.warning(f"清理命令文件失败 {filename}: {e}")
        except Exception as e:
            logger.warning(f"清理旧命令文件失败: {e}")
    
    def _update_status(self, status: Dict[str, Any]):
        """
        更新状态文件
        
        参数:
            status: 状态信息字典
        """
        try:
            with open(self.status_file, 'w', encoding='utf-8') as f:
                json.dump(status, f, ensure_ascii=False, indent=2)
            logger.debug("状态文件已更新")
        except Exception as e:
            logger.error(f"更新状态文件失败: {e}")
    
    @staticmethod
    def send_command(ipc_dir: str, command: Dict[str, Any]) -> bool:
        """
        发送命令到运行中的实例
        
        参数:
            ipc_dir: IPC 目录路径
            command: 命令字典
            
        返回:
            bool: 发送成功返回 True，失败返回 False
        """
        try:
            commands_dir = os.path.join(ipc_dir, "commands")
            
            # 确保命令目录存在
            if not os.path.exists(commands_dir):
                logger.error("命令目录不存在，程序可能未运行")
                return False
            
            # 生成唯一的命令文件名
            timestamp = int(time.time() * 1000)
            cmd_file = os.path.join(commands_dir, f"cmd_{timestamp}.json")
            
            # 写入命令文件
            with open(cmd_file, 'w', encoding='utf-8') as f:
                json.dump(command, f, ensure_ascii=False, indent=2)
            
            logger.info(f"命令已发送: {command.get('action', 'unknown')}")
            return True
            
        except Exception as e:
            logger.error(f"发送命令失败: {e}")
            return False
    
    @staticmethod
    def check_instance_running(ipc_dir: str) -> bool:
        """
        检查实例是否正在运行
        
        参数:
            ipc_dir: IPC 目录路径
            
        返回:
            bool: 实例正在运行返回 True，否则返回 False
        """
        status_file = os.path.join(ipc_dir, "status.json")
        return os.path.exists(status_file)
