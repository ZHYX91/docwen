"""
文档策略辅助函数模块

提供文档转换策略所需的辅助函数：
- 文档文件预处理（格式转换）

依赖：
- formats.document: 文档格式转换
- workspace_manager: 工作空间管理
"""

import logging
from dataclasses import dataclass, field
from pathlib import Path

from docwen.utils.path_utils import generate_output_path
from docwen.utils.workspace_manager import IntermediateItem

logger = logging.getLogger(__name__)


@dataclass
class DocumentPreprocessResult:
    processed_file: str
    actual_format: str
    intermediates: list[IntermediateItem] = field(default_factory=list)
    preprocess_chain: list[str] = field(default_factory=list)


def preprocess_document_file(
    file_path: str, temp_dir: str, cancel_event=None, actual_format: str | None = None
) -> DocumentPreprocessResult:
    """
    预处理文档文件：创建输入副本，并将非标准格式（DOC/WPS/RTF/ODT）转换为DOCX

    参数:
        file_path: 原始文件路径
        temp_dir: 临时目录路径，转换后的中间文件将输出到此目录
        cancel_event: 用于取消操作的事件对象
        actual_format: 实际文件格式（可选，如果不提供则自动检测）

    返回:
        str: 处理后的文件路径（临时目录中）
            - 所有格式都会先创建 input.{ext} 副本
            - 如果需要转换（DOC/WPS/RTF/ODT），返回转换后的DOCX路径
            - 如果已是标准格式（DOCX），返回副本路径

    说明:
        - 使用actual_format参数避免重复检测文件格式
        - 所有中间文件都输出到temp_dir，由调用者的上下文管理器统一清理
    """
    # 如果没有提供actual_format，则检测
    if actual_format is None:
        from docwen.utils.file_type_utils import detect_actual_file_format

        actual_format = detect_actual_file_format(file_path)
        logger.debug(f"自动检测文档文件格式: {actual_format}")
    else:
        logger.debug(f"使用传入的文件格式: {actual_format}")

    # 步骤1：无论什么格式，都先创建输入副本 input.{ext}
    from docwen.utils.workspace_manager import prepare_input_file

    temp_input = prepare_input_file(file_path, temp_dir, actual_format)
    logger.debug(f"已创建输入副本: {Path(temp_input).name}")

    preprocess_chain: list[str] = [actual_format]

    # 步骤2：如果是标准格式（DOCX），直接返回副本路径
    if actual_format == "docx":
        logger.debug("文件已是DOCX格式，返回副本路径")
        return DocumentPreprocessResult(
            processed_file=temp_input, actual_format=actual_format, preprocess_chain=preprocess_chain
        )

    # 步骤3：需要转换的格式，从副本转换为DOCX
    if actual_format in ["doc", "wps", "rtf", "odt"]:
        logger.info(f"检测到{actual_format.upper()}格式，从副本转换为DOCX: {Path(temp_input).name}")

        try:
            from docwen.converter.formats.document import odt_to_docx, office_to_docx, rtf_to_docx

            # 生成目标文件路径
            output_filename = generate_output_path(
                file_path,
                section="",
                add_timestamp=True,
                description=f"from{actual_format.capitalize()}",
                file_type="docx",
            )
            output_path = str(Path(temp_dir) / Path(output_filename).name)

            # 根据格式选择转换函数，使用副本作为输入
            if actual_format == "rtf":
                converted_path = rtf_to_docx(
                    temp_input,  # 使用副本
                    output_path,
                    cancel_event=cancel_event,
                )
            elif actual_format == "odt":
                converted_path = odt_to_docx(
                    temp_input,  # 使用副本
                    output_path,
                    cancel_event=cancel_event,
                )
            else:  # doc 或 wps
                converted_path = office_to_docx(
                    temp_input,  # 使用副本
                    output_path,
                    actual_format=actual_format,
                    cancel_event=cancel_event,
                )

            if converted_path:
                logger.info(f"{actual_format.upper()}转DOCX成功: {Path(converted_path).name}")
                return DocumentPreprocessResult(
                    processed_file=converted_path,
                    actual_format=actual_format,
                    intermediates=[IntermediateItem(kind="preprocessed_docx", path=converted_path)],
                    preprocess_chain=[*preprocess_chain, "docx"],
                )
            else:
                logger.error("格式转换失败，返回None")
                raise RuntimeError(f"{actual_format.upper()}转DOCX失败")

        except Exception as e:
            logger.error(f"{actual_format.upper()}转DOCX失败: {e}")
            raise RuntimeError(f"{actual_format.upper()}转DOCX失败: {e}") from e

    # 步骤4：其他不支持的格式，返回副本路径（尝试直接处理）
    logger.warning(f"不支持的文档格式: {actual_format}，返回副本路径尝试处理")
    return DocumentPreprocessResult(
        processed_file=temp_input, actual_format=actual_format, preprocess_chain=preprocess_chain
    )


_preprocess_document_file = preprocess_document_file
