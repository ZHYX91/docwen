"""
CLI包 - 公文转换器命令行界面

提供完整的命令行界面功能，包括：
- 交互式菜单系统
- Headless模式（命令行参数）
- AI友好的JSON输出
- 批量文件处理

核心模块：
- main: 命令行入口和参数解析
- interactive: 交互式菜单系统
- executor: Strategy统一调用器
- utils: 辅助函数
"""

from .main import main

__version__ = "2.0.0"
__all__ = ['main']
