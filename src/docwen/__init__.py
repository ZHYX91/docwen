"""
公文转换器 - 主包
提供公文与表格格式转换、文本校对、模板管理等功能

主要模块:
- cli: 命令行界面
- config: 配置管理
- converter: 文档转换核心（Word/Markdown/Excel多向转换）
- docx_spell: 文档校对引擎（4种校对规则）
- gui: 图形用户界面（MVP架构）
- security: 安全模块（网络隔离）
- services: 服务策略层（策略模式）
- template: 模板管理（工厂模式 + 缓存）
- utils: 工具函数库
"""

# 版本号格式: 语义化版本（SemVer）
# 例如: 0.8.1
__version__ = "0.8.2"
