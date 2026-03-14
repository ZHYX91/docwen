"""
MD文件序号处理策略

提供纯Markdown文件的小标题序号处理功能：
- 清除原有序号
- 添加新序号（支持多种方案）
- 规范化序号（先清除后添加）

适用于CLI的 process_md_numbering action 和 Obsidian插件的序号处理命令。
"""

import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.services.error_codes import (
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_UNKNOWN_ERROR,
    ERROR_CODE_UNSUPPORTED_FORMAT,
)
from docwen.services.result import ConversionResult

from .. import register_action
from ..base_strategy import BaseStrategy

logger = logging.getLogger(__name__)


@register_action("process_md_numbering")
class MdNumberingStrategy(BaseStrategy):
    """
    MD文件序号处理策略

    用于纯MD文件的序号清理和添加，不转换格式。

    支持的选项:
        - remove_numbering: bool - 是否清除原有序号
        - add_numbering: bool - 是否新增序号
        - numbering_scheme: str - 序号方案ID
            - 'gongwen_standard': 公文标准（一、（一）1.（1）①）
            - 'hierarchical_standard': 层级数字标准（1 1.1 1.1.1）
            - 'legal_standard': 法律条文标准（第一编 第一章 第一节 第一条）

    使用示例:
        strategy = MdNumberingStrategy()
        result = strategy.execute(
            file_path='document.md',
            options={
                'remove_numbering': True,
                'add_numbering': True,
                'numbering_scheme': 'gongwen_standard'
            }
        )
    """

    # 策略元信息
    name = "MD序号处理"
    description = "处理Markdown文件的小标题序号（清除/添加/规范化）"

    # 支持的源格式和目标格式
    supported_source_formats = ("md", "markdown")
    supported_target_formats = ("md", "markdown")

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行MD文件序号处理

        Args:
            file_path: MD文件路径
            options: 选项字典
                - remove_numbering: 是否清除原有序号
                - add_numbering: 是否新增序号
                - numbering_scheme: 序号方案ID
            progress_callback: 进度回调

        Returns:
            ConversionResult: 处理结果
        """
        options = options or {}

        # 获取选项
        remove_numbering = options.get("remove_numbering", False)
        add_numbering = options.get("add_numbering", False)
        numbering_scheme = options.get("numbering_scheme", "gongwen_standard")

        # 验证：至少需要执行一个操作
        if not remove_numbering and not add_numbering:
            msg = self._t(options, "conversion.messages.no_numbering_operation", default="未选择序号处理操作")
            return ConversionResult(success=False, message=msg, error_code=ERROR_CODE_INVALID_INPUT, details=msg)

        logger.info(f"开始处理MD序号: {file_path}")
        logger.debug(f"选项: remove={remove_numbering}, add={add_numbering}, scheme={numbering_scheme}")

        try:
            input_file = Path(file_path)
            # 验证文件存在
            if not input_file.exists():
                msg = f"文件不存在: {file_path}"
                return ConversionResult(success=False, message=msg, error_code=ERROR_CODE_INVALID_INPUT, details=msg)

            # 验证文件扩展名
            ext = input_file.suffix.lower()
            if ext not in [".md", ".markdown"]:
                msg = f"不支持的文件格式: {ext}，仅支持 .md/.markdown 文件"
                return ConversionResult(
                    success=False, message=msg, error_code=ERROR_CODE_UNSUPPORTED_FORMAT, details=msg
                )

            # 读取MD文件
            if progress_callback:
                progress_callback(self._t(options, "conversion.progress.reading_file", default="读取文件..."))

            with input_file.open(encoding="utf-8") as f:
                content = f.read()

            original_length = len(content)
            logger.debug(f"原始内容长度: {original_length} 字符")

            # 导入序号处理函数
            from docwen.utils.heading_numbering import process_md_numbering

            # 处理序号
            if progress_callback:
                progress_callback(self._t(options, "conversion.progress.processing_numbering", default="处理序号..."))

            processed_content = process_md_numbering(
                content=content, remove_existing=remove_numbering, add_new=add_numbering, scheme_id=numbering_scheme
            )

            processed_length = len(processed_content)
            logger.debug(f"处理后内容长度: {processed_length} 字符")

            # 写回文件
            if progress_callback:
                progress_callback(self._t(options, "conversion.progress.saving_file", default="保存文件..."))

            target_dir = input_file.parent
            suffix = input_file.suffix
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                delete=False,
                suffix=suffix,
                dir=None if str(target_dir) in {".", ""} else str(target_dir),
            ) as tf:
                tf.write(processed_content)
                temp_path = tf.name

            from docwen.utils.workspace_manager import replace_file_atomic

            replace_file_atomic(temp_path, str(input_file))

            # 构建成功消息
            operations = []
            if remove_numbering:
                operations.append("去除序号")
            if add_numbering:
                operations.append(f"添加序号({numbering_scheme})")

            message = f"序号处理完成: {', '.join(operations)}"

            logger.info(f"序号处理完成: {file_path}")

            return ConversionResult(
                success=True,
                message=message,
                output_path=file_path,
                metadata={
                    "original_length": original_length,
                    "processed_length": processed_length,
                    "operations": operations,
                    "scheme": numbering_scheme if add_numbering else None,
                },
            )

        except UnicodeDecodeError as e:
            logger.error(f"文件编码错误: {e}")
            return ConversionResult(
                success=False,
                message=f"文件编码错误，请确保文件为UTF-8编码: {e}",
                error=e,
                error_code=ERROR_CODE_INVALID_INPUT,
                details=str(e) or None,
            )

        except PermissionError as e:
            logger.error(f"文件权限错误: {e}")
            return ConversionResult(
                success=False,
                message=f"无法写入文件，请检查文件权限: {e}",
                error=e,
                error_code=ERROR_CODE_INVALID_INPUT,
                details=str(e) or None,
            )

        except Exception as e:
            logger.error(f"序号处理失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"序号处理失败: {e!s}",
                error=e,
                error_code=ERROR_CODE_UNKNOWN_ERROR,
                details=str(e) or None,
            )

    def validate_options(self, options: dict[str, Any]) -> bool:
        """
        验证选项是否有效

        Args:
            options: 选项字典

        Returns:
            bool: 选项是否有效
        """
        # 验证序号方案
        valid_schemes = ["gongwen_standard", "hierarchical_standard", "legal_standard"]
        scheme = options.get("numbering_scheme", "gongwen_standard")

        if scheme not in valid_schemes:
            logger.warning(f"无效的序号方案: {scheme}，有效方案: {valid_schemes}")
            return False

        return True
