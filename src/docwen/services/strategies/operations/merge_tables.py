"""
表格汇总策略类

实现表格汇总的策略接口，集成到项目的策略系统中。
"""

import logging
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_UNKNOWN_ERROR,
)
from docwen.services.result import ConversionResult
from docwen.table_merger import TableMerger
from docwen.translation import t

from .. import register_action
from ..base_strategy import BaseStrategy

logger = logging.getLogger(__name__)


@register_action("merge_tables")
class MergeTablesStrategy(BaseStrategy):
    """
    表格汇总策略类

    负责：
    1. 从选项中提取基准表格和待汇总表格列表
    2. 调用TableMerger执行汇总
    3. 返回标准的Result对象
    """

    def execute(
        self, file_path: str, options: dict[str, Any], progress_callback: Callable[[str], None] | None = None
    ) -> ConversionResult:
        """
        执行表格汇总

        参数:
            file_path: 基准表格文件路径
            options: 选项字典，必须包含：
                - mode: 汇总模式（1=按行, 2=按列, 3=按单元格）
                - file_list: 所有表格文件路径列表
            progress_callback: 进度回调函数

        返回:
            ConversionResult: 标准结果对象
        """
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        logger.info("表格汇总策略开始执行")
        logger.info(f"  基准表格: {Path(file_path).name}")
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

        try:
            # 步骤1：参数验证
            if progress_callback:
                progress_callback(t("conversion.progress.validating_params"))

            # 提取汇总模式
            mode = options.get("mode", 1)
            if mode not in [1, 2, 3]:
                error_msg = f"无效的汇总模式: {mode}"
                logger.error(error_msg)
                return ConversionResult(
                    success=False, message=error_msg, error_code=ERROR_CODE_INVALID_INPUT, details=error_msg
                )

            # 提取文件列表
            file_list = options.get("file_list", [])
            if not file_list:
                error_msg = "文件列表为空"
                logger.error(error_msg)
                return ConversionResult(
                    success=False, message=error_msg, error_code=ERROR_CODE_INVALID_INPUT, details=error_msg
                )

            # 确认基准表格在文件列表中
            if file_path not in file_list:
                error_msg = "基准表格不在文件列表中"
                logger.error(error_msg)
                return ConversionResult(
                    success=False, message=error_msg, error_code=ERROR_CODE_INVALID_INPUT, details=error_msg
                )

            # 构建待汇总文件列表（排除基准表格）
            collect_files = [f for f in file_list if f != file_path]

            if len(collect_files) == 0:
                error_msg = "至少需要2个表格文件才能汇总（1个基准 + 1个以上待汇总）"
                logger.error(error_msg)
                return ConversionResult(
                    success=False, message=error_msg, error_code=ERROR_CODE_INVALID_INPUT, details=error_msg
                )

            logger.info(f"参数验证通过: 基准表格=1, 待汇总表格={len(collect_files)}")

            # 步骤2：执行汇总
            logger.info("开始执行表格汇总...")

            merger = TableMerger()
            offset_range = options.get("offset_range")
            success, message, output_path = merger.merge_tables(
                base_file=file_path,
                collect_files=collect_files,
                mode=mode,
                offset_range=offset_range,
                progress_callback=progress_callback,
            )

            # 步骤3：返回结果
            if success:
                logger.info("✓ 表格汇总成功")
                logger.info(f"  输出文件: {output_path}")
                return ConversionResult(success=True, output_path=output_path, message=message)
            else:
                logger.error(f"✗ 表格汇总失败: {message}")
                return ConversionResult(
                    success=False, message=message, error_code=ERROR_CODE_CONVERSION_FAILED, details=message
                )

        except Exception as e:
            error_msg = f"表格汇总策略执行异常: {e!s}"
            logger.error(error_msg, exc_info=True)
            return ConversionResult(
                success=False,
                message=error_msg,
                error=e,
                error_code=ERROR_CODE_UNKNOWN_ERROR,
                details=str(e) or None,
            )

        finally:
            logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            logger.info("表格汇总策略执行结束")
            logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
