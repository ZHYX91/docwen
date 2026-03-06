"""
表格转 PDF 导出

支持 XLSX/XLS/ET/ODS/CSV → PDF。

依赖：
- common.fallback: 三级容错机制
- spreadsheet: 表格预处理
- config: 软件优先级配置
"""

import logging
import tempfile
import threading
from pathlib import Path

from docwen.config import config_manager

from ..common.detection import OfficeSoftwareNotFoundError
from ..common.fallback import (
    build_com_converters,
    convert_with_fallback,
    try_com_conversion,
)

logger = logging.getLogger(__name__)


def _make_excel_to_pdf_converter(input_path: str, output_path: str, prog_id: str) -> str | None:
    """Excel/WPS Spreadsheets → PDF COM 转换器（FileFormat=57）"""
    return try_com_conversion(input_path, output_path, prog_id, 57, "excel")


def xlsx_to_pdf(
    input_path: str, output_path: str | None = None, cancel_event: threading.Event | None = None
) -> str | None:
    """
    表格转 PDF（支持 XLSX/XLS/ET/ODS/CSV 输入）

    转换流程：
    1. 如果输入不是 XLSX，先转换为 XLSX
    2. 将 XLSX 转换为 PDF
    3. 使用 office_to_pdf 配置的软件优先级

    参数:
        input_path: 输入文件路径（支持 XLSX/XLS/ET/ODS/CSV）
        output_path: 输出PDF文件路径（可选，如不提供则自动生成）
        cancel_event: 取消事件（可选）

    返回:
        成功时返回输出PDF文件路径，失败时返回 None

    异常:
        OfficeSoftwareNotFoundError: 未找到可用的 Office 软件
    """
    from docwen.utils.file_type_utils import detect_actual_file_format
    from docwen.utils.path_utils import generate_output_path

    from ..spreadsheet import (
        ods_to_xlsx,
        office_to_xlsx,
    )

    # 检测实际文件格式
    actual_format = detect_actual_file_format(input_path)
    logger.info(f"表格转 PDF: {Path(input_path).name} (格式: {actual_format})")

    # 如果未提供输出路径，自动生成
    if output_path is None:
        output_path = generate_output_path(
            input_path, section="", add_timestamp=True, description=f"from{actual_format.capitalize()}", file_type="pdf"
        )
        logger.debug(f"自动生成输出路径: {output_path}")

    try:
        # 使用临时目录处理中间文件
        with tempfile.TemporaryDirectory() as temp_dir:
            # 步骤1：如果不是 XLSX，先转换为 XLSX
            if actual_format not in ["xlsx", "csv"]:
                logger.info(f"需要先将 {actual_format.upper()} 转换为 XLSX")

                # 选择合适的转换函数
                if actual_format in ["xls", "et"]:
                    converter = office_to_xlsx
                elif actual_format == "ods":
                    converter = ods_to_xlsx
                else:
                    raise ValueError(f"不支持的表格格式: {actual_format}")

                # 生成临时 XLSX 文件路径
                temp_xlsx = str(Path(temp_dir) / "intermediate.xlsx")

                # 转换为 XLSX
                xlsx_result = converter(input_path, temp_xlsx, actual_format=actual_format, cancel_event=cancel_event)

                if not xlsx_result or not Path(xlsx_result).exists():
                    logger.error(f"{actual_format.upper()} 转 XLSX 失败")
                    raise OfficeSoftwareNotFoundError(f"无法将 {actual_format.upper()} 转换为 XLSX")

                # 使用转换后的 XLSX 作为输入
                pdf_input = xlsx_result
                logger.info(f"✓ {actual_format.upper()} → XLSX 转换成功")
            elif actual_format == "csv":
                # CSV 需要先转换为 XLSX
                logger.info("需要先将 CSV 转换为 XLSX")
                from docwen.converter.formats.spreadsheet import csv_to_xlsx

                temp_xlsx = str(Path(temp_dir) / "intermediate.xlsx")
                xlsx_result = csv_to_xlsx(input_path, output_path=temp_xlsx)

                if not xlsx_result or not Path(xlsx_result).exists():
                    logger.error("CSV 转 XLSX 失败")
                    raise OfficeSoftwareNotFoundError("无法将 CSV 转换为 XLSX")

                pdf_input = xlsx_result
                logger.info("✓ CSV → XLSX 转换成功")
            else:
                # 已是 XLSX，直接使用
                pdf_input = input_path
                logger.debug("输入已是 XLSX 格式，直接转 PDF")

            if cancel_event and cancel_event.is_set():
                logger.info("转换被取消")
                return None

            # 步骤2：XLSX → PDF
            logger.info("开始转换 XLSX → PDF")

            # 获取表格转PDF优先级配置
            software_priority = config_manager.get_spreadsheet_to_pdf_priority()
            logger.debug(f"表格转PDF软件优先级: {software_priority}")

            # 构建转换器列表
            converters = build_com_converters(software_priority, _make_excel_to_pdf_converter)

            temp_output_path = str(Path(temp_dir) / Path(output_path).name)

            # 执行转换（LibreOffice 也支持 PDF 导出）
            result, used_software = convert_with_fallback(
                input_path=pdf_input,
                output_path=temp_output_path,
                com_converters=converters,
                libreoffice_format="pdf",
                cancel_event=cancel_event,
            )

            if not result:
                logger.error("✗ XLSX → PDF 转换失败")
                raise OfficeSoftwareNotFoundError("无法转换为 PDF")

            from docwen.utils.workspace_manager import finalize_output

            saved_path, _ = finalize_output(
                result,
                output_path,
                original_input_file=input_path,
            )
            if not saved_path:
                raise OfficeSoftwareNotFoundError("保存输出PDF失败")

            logger.info(f"✓ PDF 转换成功: {Path(saved_path).name} (使用软件: {used_software})")
            return saved_path

    except OfficeSoftwareNotFoundError:
        raise
    except Exception as e:
        logger.error(f"表格转 PDF 失败: {e}", exc_info=True)
        raise OfficeSoftwareNotFoundError(f"转换失败: {e}") from e
