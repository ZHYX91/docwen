"""
服务层包 (Service Layer)

该包负责编排核心业务逻辑，处理文件转换、校对等任务流程。
它作为表现层 (GUI/CLI) 和底层转换器 (Converter) 之间的桥梁。
"""

# 日志配置
import logging
logger = logging.getLogger(__name__)
logger.info("服务层包初始化")
