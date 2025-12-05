"""
表格文件到Markdown的转换策略及表格格式互转策略
"""
import os
import logging
from .base_strategy import BaseStrategy
from gongwen_converter.services.result import ConversionResult
from . import register_conversion, CATEGORY_SPREADSHEET
from gongwen_converter.converter.xlsx2md import convert_spreadsheet_to_md
from gongwen_converter.converter.table_converters import convert_csv_to_xlsx, convert_xlsx_to_csv
from gongwen_converter.utils.path_utils import generate_output_path
from typing import Dict, Any, Callable, Optional
import tempfile
import shutil

logger = logging.getLogger(__name__)


def _preprocess_table_file(file_path: str, temp_dir: str, cancel_event=None, actual_format: str = None) -> str:
    """
    预处理表格文件：创建输入副本，并将非标准格式（XLS/ET/ODS）转换为XLSX
    
    参数:
        file_path: 原始文件路径
        temp_dir: 临时目录路径，转换后的中间文件将输出到此目录
        cancel_event: 用于取消操作的事件对象
        actual_format: 实际文件格式（可选，如果不提供则自动检测）
        
    返回:
        str: 处理后的文件路径（临时目录中）
            - 所有格式都会先创建 input.{ext} 副本
            - 如果需要转换（XLS/ET/ODS），返回转换后的XLSX路径
            - 如果已是标准格式（XLSX/CSV），返回副本路径
            
    说明:
        - 使用actual_format参数避免重复检测文件格式
        - 所有中间文件都输出到temp_dir，由调用者的上下文管理器统一清理
    """
    # 如果没有提供actual_format，则检测
    if actual_format is None:
        from gongwen_converter.utils.file_type_utils import detect_actual_file_format
        actual_format = detect_actual_file_format(file_path)
        logger.debug(f"自动检测表格文件格式: {actual_format}")
    else:
        logger.debug(f"使用传入的文件格式: {actual_format}")
    
    # 步骤1：无论什么格式，都先创建输入副本 input.{ext}
    from gongwen_converter.utils.workspace_manager import prepare_input_file
    temp_input = prepare_input_file(file_path, temp_dir, actual_format)
    logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
    
    # 步骤2：如果是标准格式（XLSX/CSV），直接返回副本路径
    if actual_format in ['xlsx', 'csv']:
        logger.debug(f"文件已是标准表格格式({actual_format})，返回副本路径")
        return temp_input
    
    # 步骤3：需要转换的格式，从副本转换为XLSX
    if actual_format in ['xls', 'et', 'ods']:
        logger.info(f"检测到{actual_format.upper()}格式，从副本转换为XLSX: {os.path.basename(temp_input)}")
        
        try:
            from gongwen_converter.converter.formats.office import (
                office_to_xlsx, ods_to_xlsx
            )
            
            # 生成目标文件路径
            output_filename = generate_output_path(
                file_path,
                section="",
                add_timestamp=True,
                description=f'from{actual_format.capitalize()}',
                file_type='xlsx'
            )
            output_path = os.path.join(temp_dir, os.path.basename(output_filename))
            
            # 根据格式选择转换函数，使用副本作为输入
            if actual_format == 'ods':
                converted_path = ods_to_xlsx(
                    temp_input,  # 使用副本
                    output_path,
                    cancel_event=cancel_event
                )
            else:  # xls 或 et
                converted_path = office_to_xlsx(
                    temp_input,  # 使用副本
                    output_path,
                    actual_format=actual_format,
                    cancel_event=cancel_event
                )
            
            if converted_path:
                logger.info(f"{actual_format.upper()}转XLSX成功: {os.path.basename(converted_path)}")
                return converted_path
            else:
                logger.error("格式转换失败，返回None")
                raise RuntimeError(f"{actual_format.upper()}转XLSX失败")
                
        except Exception as e:
            logger.error(f"{actual_format.upper()}转XLSX失败: {e}")
            raise RuntimeError(f"{actual_format.upper()}转XLSX失败: {e}")
    
    # 步骤4：其他不支持的格式，返回副本路径（尝试直接处理）
    logger.warning(f"不支持的表格格式: {actual_format}，返回副本路径尝试处理")
    return temp_input

@register_conversion(CATEGORY_SPREADSHEET, 'md')
class SpreadsheetToMarkdownStrategy(BaseStrategy):
    """
    将表格文件转换为Markdown文件的策略。
    
    支持的输入格式：
    - XLSX (Excel 2007+)：直接处理
    - XLS (Excel 97-2003)：自动转换为XLSX后处理
    - ET (WPS表格格式)：自动转换为XLSX后处理
    - CSV (逗号分隔值)：直接处理
    
    输出格式：
    - Markdown表格语法
    - 根据输入文件类型自动命名（fromXlsx 或 fromCsv）
    """
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行表格到Markdown的转换（支持XLS/ET自动转换）。
        
        Args:
            file_path: 输入的表格文件路径（XLSX/XLS/ET/CSV）
            options: 转换选项字典，包含：
                - cancel_event: (可选) 用于取消操作的事件对象
                - extract_image: (可选) 是否提取图片
                - extract_ocr: (可选) 是否进行OCR识别
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            - success: 转换是否成功
            - output_path: 生成的Markdown文件路径
            - message: 成功或失败的描述信息
            - error: 失败时的错误信息
            
        Note:
            - 如果输入是XLS/ET，会先自动转换为XLSX
            - CSV文件输出将标记为 "fromCsv"
            - 其他格式输出将标记为 "fromXlsx"
            - 支持图片提取和OCR功能（与文档转MD一致）
        """
        try:
            if progress_callback:
                progress_callback("正在转换为Markdown...")
            
            options = options or {}
            cancel_event = options.get("cancel_event")
            actual_format = options.get("actual_format")  # 从options中提取actual_format
            
            # 从GUI获取导出选项
            extract_image = options.get('extract_image', False)
            extract_ocr = options.get('extract_ocr', False)
            
            logger.info(f"表格转MD - 导出选项: 提取图片={extract_image}, OCR={extract_ocr}")
            
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            import tempfile
            import shutil
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 将XLS/ET转换为XLSX（如需要）
                if progress_callback:
                    progress_callback("检测文件格式...")
                
                processed_file = _preprocess_table_file(file_path, temp_dir, cancel_event, actual_format)
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成统一basename和创建临时子文件夹（在转换前）
                # 根据actual_format生成description
                description = f"from{actual_format.capitalize()}" if actual_format else "fromXlsx"
                
                # 生成基础路径（用于提取basename）
                base_path = generate_output_path(
                    file_path,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="md"
                )
                
                # 提取basename（不含.md扩展名）
                basename = os.path.splitext(os.path.basename(base_path))[0]
                logger.debug(f"统一basename: {basename}")
                
                # 创建临时子文件夹
                temp_output_folder = os.path.join(temp_dir, basename)
                os.makedirs(temp_output_folder, exist_ok=True)
                logger.debug(f"创建临时子文件夹: {temp_output_folder}")
                
                # 步骤3：调用核心转换函数，传递图片选项和输出文件夹
                if progress_callback:
                    progress_callback("转换中...")
                
                markdown_content = convert_spreadsheet_to_md(
                    processed_file,
                    extract_image=extract_image,
                    extract_ocr=extract_ocr,
                    output_folder=temp_output_folder,  # 传递输出文件夹用于图片提取
                    original_file_path=file_path  # 传递原始文件路径用于图片命名
                )

                # 步骤4：写入Markdown文件到临时子文件夹
                if progress_callback:
                    progress_callback("正在写入...")
                
                # 构建文件名并写入MD
                md_filename = f"{basename}.md"
                temp_output = os.path.join(temp_output_folder, md_filename)
                
                logger.debug(f"准备将Markdown内容写入: {temp_output}")
                with open(temp_output, 'w', encoding='utf-8') as f:
                    f.write(markdown_content)
                logger.info(f"Markdown内容已写入: {temp_output}")
                
                # 步骤5：移动整个文件夹到输出目录
                final_folder = os.path.join(output_dir, basename)
                
                # 如果目标文件夹已存在，先删除
                if os.path.exists(final_folder):
                    shutil.rmtree(final_folder)
                    logger.debug(f"已删除现有文件夹: {final_folder}")
                
                shutil.move(temp_output_folder, final_folder)
                logger.info(f"已移动文件夹到: {final_folder}")
                
                # 步骤6：如果保留中间文件，移动temp_dir中的其他规范文件（如中间XLSX）
                should_keep = self._should_keep_intermediates()
                if should_keep:
                    logger.info("检查并移动临时目录中的其他中间文件")
                    for filename in os.listdir(temp_dir):
                        # 排除输入副本和已移动的文件夹
                        if filename.startswith('input.') or filename == basename:
                            continue
                        src = os.path.join(temp_dir, filename)
                        if os.path.isfile(src):
                            dst = os.path.join(output_dir, filename)
                            shutil.move(src, dst)
                            logger.info(f"保留中间文件: {filename}")
                
                # 准备返回路径（MD文件的完整路径）
                output_path = os.path.join(final_folder, md_filename)
                
                if progress_callback:
                    progress_callback("转换完成。")

                return ConversionResult(
                    success=True,
                    output_path=output_path,
                    message="转换为Markdown成功。"
                )

        except Exception as e:
            logger.error(f"表格转Markdown失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取清理配置失败: {e}，使用默认设置（清理中间文件）")
            return False


@register_conversion(CATEGORY_SPREADSHEET, 'txt')
class SpreadsheetToTxtStrategy(BaseStrategy):
    """
    将表格文件转换为TXT文件的策略。
    
    实现说明：
    - 复用 SpreadsheetToMarkdownStrategy 的转换逻辑
    - 先将表格转换为Markdown格式
    - 然后以TXT扩展名保存（内容仍为Markdown格式）
    
    支持的输入格式：
    - XLSX (Excel 2007+)
    - XLS (Excel 97-2003)
    - ET (WPS表格格式)
    - CSV (逗号分隔值)
    
    Note:
        输出的TXT文件实际包含Markdown格式的表格语法，
        便于在纯文本编辑器中查看和编辑。
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行表格到TXT的转换。
        
        Args:
            file_path: 输入的表格文件路径
            options: 转换选项字典（当前未使用）
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            - success: 转换是否成功
            - output_path: 生成的TXT文件路径
            - message: 成功或失败的描述信息
            - error: 失败时的错误信息
        """
        try:
            if progress_callback:
                progress_callback("正在转换为TXT...")

            # 从options中提取actual_format
            options = options or {}
            actual_format = options.get("actual_format")
            
            # 根据actual_format生成description
            description = f"from{actual_format.capitalize()}" if actual_format else "fromXlsx"

            # 先转换为Markdown格式（内部表示）
            markdown_content = convert_spreadsheet_to_md(file_path)
            
            # 生成TXT输出路径（而非MD路径）
            txt_output = generate_output_path(file_path, section="", add_timestamp=True, description=description, file_type="txt")
            
            # 将Markdown格式内容写入TXT文件
            logger.debug(f"准备将内容写入TXT文件: {txt_output}")
            try:
                with open(txt_output, 'w', encoding='utf-8') as f:
                    f.write(markdown_content)
                logger.info(f"成功将内容写入TXT文件: {txt_output}")
            except Exception as e:
                logger.error(f"写入TXT文件失败: {e}", exc_info=True)
                raise IOError(f"无法写入输出文件: {e}")
            
            if progress_callback:
                progress_callback("转换完成。")

            return ConversionResult(
                success=True,
                output_path=txt_output,
                message="转换为TXT成功。"
            )

        except Exception as e:
            logger.error(f"表格转TXT失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                error=f"转换失败: {e}"
            )


@register_conversion('csv', 'xlsx')
class CsvToXlsxStrategy(BaseStrategy):
    """
    将CSV文件转换为XLSX文件的策略
    
    转换说明：
    - CSV文件将被读取并保存为XLSX格式
    - 输出文件使用默认工作表名 "Sheet1"
    - 文件名标记为 "fromCsv"
    - 支持扩展名不匹配的文件
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行CSV到XLSX的转换
        
        Args:
            file_path: 输入的CSV文件路径
            options: 转换选项字典，包含：
                - actual_format: (可选) 文件的真实格式
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            # 从options中提取参数
            options = options or {}
            actual_format = options.get('actual_format', 'csv')
            cancel_event = options.get('cancel_event')
            output_dir = os.path.dirname(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.csv
                if progress_callback:
                    progress_callback("准备文件...")
                
                from gongwen_converter.utils.workspace_manager import prepare_input_file
                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：读取CSV并转换
                if progress_callback:
                    progress_callback("正在转换为XLSX...")
                
                import pandas as pd
                df = pd.read_csv(temp_input, header=None, keep_default_na=False)
                logger.debug(f"CSV 文件读取成功，数据形状: {df.shape}")
                
                # 步骤3：生成输出文件名
                output_filename = os.path.basename(
                    generate_output_path(
                        file_path,
                        section="",
                        add_timestamp=True,
                        description="fromCsv",
                        file_type="xlsx"
                    )
                )
                
                # 步骤4：保存到临时目录
                temp_output = os.path.join(temp_dir, output_filename)
                with pd.ExcelWriter(temp_output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
                logger.debug(f"XLSX文件已生成: {temp_output}")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤5：移动到目标位置
                final_output = os.path.join(output_dir, output_filename)
                shutil.move(temp_output, final_output)
                logger.info(f"CSV转XLSX成功: {final_output}")
            
            if progress_callback:
                progress_callback("转换完成。")
            
            return ConversionResult(
                success=True,
                output_path=final_output,
                message="转换为XLSX成功。"
            )
            
        except Exception as e:
            logger.error(f"CSV转XLSX失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )


@register_conversion('xlsx', 'csv')
class XlsxToCsvStrategy(BaseStrategy):
    """
    将XLSX文件转换为CSV文件的策略
    
    转换说明：
    - 每个工作表将生成一个独立的CSV文件
    - 工作表名中的空格将被替换为下划线
    - 文件名包含工作表名作为section，标记为 "fromXlsx"
    - 所有CSV文件使用相同的时间戳
    - 支持扩展名不匹配的文件
    - 所有CSV文件输出到同一个子文件夹
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行XLSX到CSV的转换
        
        Args:
            file_path: 输入的XLSX文件路径
            options: 转换选项字典，包含：
                - actual_format: (可选) 文件的真实格式
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            
        Note:
            output_path 返回子文件夹中第一个CSV文件的路径，
            以便在批量模式下正确定位到转换结果
        """
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            # 从options中提取参数
            options = options or {}
            actual_format = options.get('actual_format', 'xlsx')
            cancel_event = options.get('cancel_event')
            
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用临时目录管理输出
            import tempfile
            import shutil
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.xlsx
                if progress_callback:
                    progress_callback("准备文件...")
                
                from gongwen_converter.utils.workspace_manager import prepare_input_file
                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成统一basename和时间戳描述部分
                base_path = generate_output_path(
                    file_path,
                    section="",
                    add_timestamp=True,
                    description="fromXlsx",
                    file_type="csv"
                )
                basename = os.path.splitext(os.path.basename(base_path))[0]
                logger.debug(f"统一basename: {basename}")
                
                # 提取原始文件名（不含扩展名、时间戳、描述）
                import re
                original_file_basename = os.path.splitext(os.path.basename(file_path))[0]
                # 移除可能存在的旧时间戳和描述
                timestamp_pattern = r'(_\d{8}_\d{6})(?:.*)?$'
                match = re.search(timestamp_pattern, original_file_basename)
                if match:
                    original_file_basename = original_file_basename[:match.start()]
                logger.debug(f"原始文件basename: {original_file_basename}")
                
                # 从basename中提取时间戳和描述部分
                parts = basename.split('_')
                unified_timestamp_desc = ""
                if len(parts) >= 3:
                    # 找到时间戳的位置（格式: YYYYMMDD）
                    timestamp_idx = None
                    for i, part in enumerate(parts):
                        if len(part) == 8 and part.isdigit():
                            timestamp_idx = i
                            break
                    
                    if timestamp_idx is not None:
                        # 提取时间戳和description部分
                        unified_timestamp_desc = '_'.join(parts[timestamp_idx:])
                
                if not unified_timestamp_desc:
                    # 降级方案
                    unified_timestamp_desc = '_'.join(parts[1:]) if len(parts) > 1 else parts[0]
                
                logger.debug(f"统一时间戳描述: {unified_timestamp_desc}")
                
                # 步骤3：创建临时子文件夹
                temp_output_folder = os.path.join(temp_dir, basename)
                os.makedirs(temp_output_folder, exist_ok=True)
                logger.debug(f"创建临时子文件夹: {temp_output_folder}")
                
                # 步骤4：转换副本，输出到临时子文件夹
                # 传递原始文件名和统一的时间戳描述，确保CSV文件名与子文件夹名一致
                if progress_callback:
                    progress_callback("正在转换为CSV...")
                
                csv_files = convert_xlsx_to_csv(
                    temp_input,  # 使用副本
                    actual_format, 
                    output_dir=temp_output_folder,
                    original_basename=original_file_basename,  # 传递原始文件basename
                    unified_timestamp_desc=unified_timestamp_desc  # 传递统一的时间戳描述
                )
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                if not csv_files:
                    return ConversionResult(
                        success=False,
                        message="未生成任何CSV文件（工作表可能为空）"
                    )
                
                logger.info(f"已生成 {len(csv_files)} 个CSV文件到临时子文件夹")
                
                # 步骤5：移动整个文件夹到输出目录
                final_folder = os.path.join(output_dir, basename)
                
                # 如果目标文件夹已存在，先删除
                if os.path.exists(final_folder):
                    shutil.rmtree(final_folder)
                    logger.debug(f"已删除现有文件夹: {final_folder}")
                
                shutil.move(temp_output_folder, final_folder)
                logger.info(f"已移动文件夹到: {final_folder}")
                
                # 准备返回路径（第一个CSV文件的完整路径）
                first_csv_name = os.path.basename(csv_files[0])
                output_path = os.path.join(final_folder, first_csv_name)
            
            if progress_callback:
                progress_callback(f"转换完成，共生成 {len(csv_files)} 个CSV文件。")
            
            # 返回第一个CSV文件路径用于定位（现在在子文件夹内）
            return ConversionResult(
                success=True,
                output_path=output_path,
                message=f"转换为CSV成功，共生成 {len(csv_files)} 个文件。"
            )
            
        except Exception as e:
            logger.error(f"XLSX转CSV失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )


@register_conversion(CATEGORY_SPREADSHEET, 'pdf')
class SpreadsheetToPdfStrategy(BaseStrategy):
    """
    将表格文件转换为PDF文件的策略。
    
    功能特性：
    - 使用本地Office软件（WPS或Microsoft Office）进行转换
    - 支持XLSX、XLS、ET、ODS、CSV等表格格式
    - 转换质量高，能保持表格格式和样式
    - 生成不可编辑的PDF文档，适合最终版本归档
    - 使用 office_to_pdf 配置的软件优先级
    
    支持的输入格式：
    - XLSX (Excel 2007+)
    - XLS (Excel 97-2003)
    - ET (WPS表格格式)
    - ODS (OpenDocument表格)
    - CSV (逗号分隔值)
    
    Note:
        需要本地安装WPS Office、Microsoft Office或LibreOffice才能使用此功能。
        如果都未安装，将返回错误提示。
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行表格到PDF的转换。
        
        Args:
            file_path: 输入的表格文件路径（支持 XLSX/XLS/ET/ODS/CSV）
            options: 转换选项字典，包含：
                - cancel_event: (可选) 用于取消操作的事件对象
                - actual_format: (可选) 文件的真实格式
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            - success: 转换是否成功
            - output_path: 生成的PDF文件路径
            - message: 成功或失败的描述信息
            - error: 失败时的错误信息
        """
        # 在try块外导入异常类，避免UnboundLocalError
        from gongwen_converter.converter.formats.office import OfficeSoftwareNotFoundError
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            options = options or {}
            cancel_event = options.get('cancel_event')
            actual_format = options.get('actual_format', 'xlsx')
            
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.{ext}
                if progress_callback:
                    progress_callback("准备文件...")
                
                from gongwen_converter.utils.workspace_manager import prepare_input_file
                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成输出文件名
                output_filename = os.path.basename(
                    generate_output_path(
                        file_path,
                        section="",
                        add_timestamp=True,
                        description=f"from{actual_format.capitalize()}",
                        file_type="pdf"
                    )
                )
                
                # 步骤3：在临时目录进行转换
                temp_output = os.path.join(temp_dir, output_filename)
                
                if progress_callback:
                    progress_callback("正在转换为PDF...")
                
                # 导入并调用转换函数
                from gongwen_converter.converter.formats.office import xlsx_to_pdf
                
                result_path = xlsx_to_pdf(
                    temp_input,  # 使用副本
                    output_path=temp_output,  # 输出到临时目录
                    cancel_event=cancel_event
                )
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                if not result_path or not os.path.exists(result_path):
                    return ConversionResult(success=False, message="转换失败，请查看日志。")
                
                # 步骤4：移动到目标位置
                final_output = os.path.join(output_dir, output_filename)
                shutil.move(result_path, final_output)
                logger.info(f"PDF转换完成，文件已保存: {final_output}")
                
                # 步骤5：如果保留中间文件，移动temp_dir中的其他规范文件
                should_keep = self._should_keep_intermediates()
                if should_keep:
                    logger.info("检查并移动临时目录中的其他中间文件")
                    for filename in os.listdir(temp_dir):
                        # 排除输入副本和已移动的文件
                        if filename.startswith('input.') or filename == output_filename:
                            continue
                        src = os.path.join(temp_dir, filename)
                        if os.path.isfile(src):
                            dst = os.path.join(output_dir, filename)
                            shutil.move(src, dst)
                            logger.info(f"保留中间文件: {filename}")
                
                if progress_callback:
                    progress_callback("转换完成。")
                
                return ConversionResult(
                    success=True,
                    output_path=final_output,
                    message="转换为PDF成功。"
                )
            
        except OfficeSoftwareNotFoundError as e:
            logger.error(f"表格转PDF失败 - 未找到Office软件: {e}")
            return ConversionResult(
                success=False,
                message="未找到Office软件（WPS、Microsoft Office或LibreOffice），无法转换为PDF。",
                error=str(e)
            )
        except Exception as e:
            logger.error(f"表格转PDF失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False


# ==================== 智能转换链：策略工厂 ====================

def _create_spreadsheet_conversion_strategy(source_fmt: str, target_fmt: str):
    """
    策略工厂：动态创建表格格式转换策略
    
    使用智能转换链自动处理单步或多步转换
    
    参数:
        source_fmt: 源格式（如 'xls', 'ods'）
        target_fmt: 目标格式（如 'xlsx', 'ods'）
    
    返回:
        动态生成的策略类
    """
    
    @register_conversion(source_fmt, target_fmt)
    class DynamicSpreadsheetConversionStrategy(BaseStrategy):
        """动态生成的表格转换策略"""
        
        def execute(
            self,
            file_path: str,
            options: Optional[Dict[str, Any]] = None,
            progress_callback: Optional[Callable[[str], None]] = None
        ) -> ConversionResult:
            """
            执行表格格式转换
            
            说明:
                从options中提取actual_format参数并传递给SmartConverter，
                确保即使文件扩展名被修改也能正确转换
            """
            try:
                options = options or {}
                cancel_event = options.get('cancel_event')
                preferred_software = options.get('preferred_software')
                actual_format = options.get('actual_format')  # 提取真实格式
                
                # 如果没有提供actual_format，从文件扩展名推断（降级方案）
                if not actual_format:
                    _, ext = os.path.splitext(file_path)
                    actual_format = ext.lower().lstrip('.')
                    logger.warning(f"未提供actual_format，从文件名推断: {actual_format}")
                
                logger.debug(f"表格转换策略: {source_fmt}→{target_fmt}, 真实格式: {actual_format}")
                
                # 使用智能转换链
                from gongwen_converter.converter.smart_converter import SmartConverter, OfficeSoftwareNotFoundError
                from gongwen_converter.utils.workspace_manager import get_output_directory
                
                converter = SmartConverter()
                output_dir = get_output_directory(file_path)
                
                # 使用临时目录管理中间文件
                import tempfile
                import shutil
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 调用SmartConverter，输出到临时目录
                    result_path = converter.convert(
                        input_path=file_path,
                        target_format=target_fmt,
                        category='spreadsheet',
                        actual_format=actual_format,  # 传递真实格式
                        output_dir=temp_dir,  # 输出到临时目录
                        cancel_event=cancel_event,
                        progress_callback=progress_callback,
                        preferred_software=preferred_software
                    )
                    
                    if not result_path or not os.path.exists(result_path):
                        return ConversionResult(
                            success=False,
                            message=f"转换为{target_fmt.upper()}失败。"
                        )
                    
                    # 检查是否是CSV转换（result_path在子文件夹中）
                    if target_fmt == 'csv':
                        # CSV转换：result_path是子文件夹内的文件
                        csv_subfolder = os.path.dirname(result_path)
                        subfolder_name = os.path.basename(csv_subfolder)
                        
                        # 移动整个CSV子文件夹到输出目录
                        final_output_folder = os.path.join(output_dir, subfolder_name)
                        
                        # 如果目标文件夹已存在，先删除
                        if os.path.exists(final_output_folder):
                            shutil.rmtree(final_output_folder)
                            logger.debug(f"已删除现有文件夹: {final_output_folder}")
                        
                        # 移动整个文件夹
                        shutil.move(csv_subfolder, final_output_folder)
                        logger.info(f"CSV文件夹已移动: {subfolder_name}")
                        
                        # 更新返回路径
                        final_output_path = os.path.join(final_output_folder, os.path.basename(result_path))
                        
                        # 处理中间文件（如果需要保留）
                        should_keep = self._should_keep_intermediates()
                        if should_keep:
                            logger.info("检查是否有中间文件需要保留")
                            for filename in os.listdir(temp_dir):
                                # 排除输入副本和已移动的CSV文件夹
                                if filename.startswith('input.') or filename == subfolder_name:
                                    continue
                                src = os.path.join(temp_dir, filename)
                                if os.path.isfile(src):
                                    dst = os.path.join(output_dir, filename)
                                    shutil.move(src, dst)
                                    logger.info(f"保留中间文件: {filename}")
                    else:
                        # 非CSV转换：移动单个文件
                        final_output_path = os.path.join(output_dir, os.path.basename(result_path))
                        
                        # 根据配置移动文件
                        should_keep = self._should_keep_intermediates()
                        if should_keep:
                            # 保留中间文件（排除输入副本）
                            logger.info("保留中间文件，移动规范命名的文件到输出目录")
                            for filename in os.listdir(temp_dir):
                                # 排除输入副本文件
                                if filename.startswith('input.'):
                                    logger.debug(f"跳过输入副本: {filename}")
                                    continue
                                src = os.path.join(temp_dir, filename)
                                if os.path.isfile(src):
                                    dst = os.path.join(output_dir, filename)
                                    shutil.move(src, dst)
                                    logger.debug(f"保留中间文件: {filename}")
                        else:
                            # 只移动最终文件
                            logger.debug("清理中间文件，只移动最终文件")
                            shutil.move(result_path, final_output_path)
                            logger.debug(f"已移动最终文件: {os.path.basename(result_path)}")
                    
                    return ConversionResult(
                        success=True,
                        output_path=final_output_path,
                        message=f"转换为{target_fmt.upper()}成功。"
                    )
            
            except OfficeSoftwareNotFoundError as e:
                logger.error(f"未找到Office软件: {e}")
                return ConversionResult(
                    success=False,
                    message="未找到Office软件（WPS或Microsoft Office），无法完成转换。",
                    error=str(e)
                )
            except Exception as e:
                logger.error(f"{source_fmt.upper()}转{target_fmt.upper()}失败: {e}", exc_info=True)
                return ConversionResult(
                    success=False,
                    message=f"转换失败: {e}",
                    error=e
                )
        
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False
    
    return DynamicSpreadsheetConversionStrategy


# 批量注册表格格式转换策略（排除 CSV ↔ XLSX，它们已有独立实现）
SPREADSHEET_FORMATS = ['xlsx', 'xls', 'ods', 'et']
for source in SPREADSHEET_FORMATS:
    for target in SPREADSHEET_FORMATS:
        if source != target:
            # 注意：这里会覆盖上面手写的 XlsxToXlsStrategy 和 XlsxToOdsStrategy
            _create_spreadsheet_conversion_strategy(source, target)

logger.info("表格格式转换策略已通过智能转换链批量注册")

# 添加 CSV 到其他格式的转换支持（CSV → XLSX 已有独立实现）
# CSV → XLS: CSV → XLSX → XLS
# CSV → ODS: CSV → XLSX → ODS
for target in ['xls', 'ods']:
    _create_spreadsheet_conversion_strategy('csv', target)

logger.info("CSV 转换策略已注册（CSV → XLS, CSV → ODS）")

# 添加其他格式到 CSV 的转换支持（XLSX → CSV 已有独立实现）
# XLS → CSV: XLS → XLSX → CSV
# ODS → CSV: ODS → XLSX → CSV
# ET → CSV: ET → XLSX → CSV
for source in ['xls', 'ods', 'et']:
    _create_spreadsheet_conversion_strategy(source, 'csv')

logger.info("其他格式转CSV策略已注册（XLS/ODS/ET → CSV）")
