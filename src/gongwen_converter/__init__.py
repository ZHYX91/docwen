"""
公文转换器 - 主包
提供公文与表格格式转换、文本校对、模板管理等功能

主要模块:
- cli: 命令行界面
- config: 配置管理
- converter: 文档转换核心（Word/Markdown/Excel多向转换）
- docx_spell: 文档校对引擎（4种校对规则）
- gui: 图形用户界面（MVP架构）
- security: 安全模块（网络隔离 + 试用期检查）
- services: 服务策略层（策略模式）
- template: 模板管理（工厂模式 + 缓存）
- utils: 工具函数库
"""

# 版本号格式: 0.1.0.YYYYMMDD.REV
# 0.1.0: 基础开发版本
# YYYYMMDD: 构建日期
# REV: 当天修订号（从1开始）
__version__ = "0.5.1.20260101.1"
