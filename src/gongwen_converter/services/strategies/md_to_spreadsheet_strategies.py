"""
处理与Excel和电子表格相关的转换策略。
"""
import os
import shutil
import logging
import tempfile
from typing import Dict, Any, Callable, Optional

from gongwen_converter.services.result import ConversionResult
from gongwen_converter.services.strategies.base_strategy import BaseStrategy
from . import register_conversion, CATEGORY_MARKDOWN
from gongwen_converter.utils.path_utils import generate_output_path
from gongwen_converter.config.config_manager import config_manager
from gongwen_converter.converter.md2xlsx.core import convert as convert_md_to_xlsx
from gongwen_converter.converter.formats.office import (
    convert_xlsx_to_xls,
    office_to_xlsx,
    OfficeSoftwareNotFoundError,
    check_office_availability
)

logger = logging.getLogger(__name__)

class BaseMdToSpreadsheetStrategy(BaseStrategy):
    """
    MD转电子表格系列策略的基类。
    
    提供了将Markdown文件转换为电子表格的通用流程：
    1. 将MD转换为XLSX（中间格式）
    2. 可选：将XLSX进一步转换为其他格式（XLS、ODS、CSV等）
    3. 根据配置决定是否保留中间文件
    
    使用临时目录处理中间文件，确保转换过程的原子性。
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行Markdown到电子表格的转换。
        
        Args:
            file_path: 输入的Markdown文件路径
            options: 转换选项字典，包含：
                - template_name: (必需) 使用的模板名称
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象，包括成功状态、输出路径等
            
        Raises:
            OfficeSoftwareNotFoundError: 当需要Office/WPS软件但未找到时
        """
        if options is None:
            options = {}
            
        cancel_event = options.get("cancel_event")
        update_progress = progress_callback or (lambda msg: None)

        try:
            template_name = options.get("template_name")
            if not template_name:
                return ConversionResult(success=False, message="错误：未提供模板名称。")

            from gongwen_converter.utils.workspace_manager import prepare_input_file, get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤0: 创建输入副本 input.md
                temp_md = prepare_input_file(file_path, temp_dir, 'md')
                logger.debug(f"已创建输入副本: {os.path.basename(temp_md)}")
                
                # 步骤 1: MD -> XLSX (在临时目录中，使用副本)
                update_progress("正在将Markdown转换为XLSX...")
                intermediate_xlsx_path = self._convert_md_to_xlsx(temp_md, template_name, progress_callback, cancel_event, temp_dir, file_path)
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                if not intermediate_xlsx_path:
                    return ConversionResult(success=False, message="Markdown到XLSX转换失败。")

                # 步骤 2: (可选) XLSX -> 最终格式 (在临时目录中)
                update_progress("正在生成最终电子表格格式...")
                final_temp_path = self._finalize_conversion(intermediate_xlsx_path, file_path, progress_callback, cancel_event, temp_dir)
                if cancel_event and cancel_event.is_set(): 
                    return ConversionResult(success=False, message="操作已取消")
                if not (final_temp_path and os.path.exists(final_temp_path)):
                    return ConversionResult(success=False, message="最终格式生成失败。")

                # 准备最终输出路径
                final_output_path = os.path.join(output_dir, os.path.basename(final_temp_path))
                
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
                    shutil.move(final_temp_path, final_output_path)

            return ConversionResult(success=True, output_path=final_output_path, message="转换为Excel成功。")
        
        except OfficeSoftwareNotFoundError as e:
            logger.error(f"Office/WPS软件未找到: {e}")
            # 根据具体策略类提供精确的错误提示
            if isinstance(self, MdToXlsStrategy):
                return ConversionResult(success=False, message="XLS格式转换需要安装WPS或Microsoft Office任一软件", error=e)
            else:
                return ConversionResult(success=False, message=str(e), error=e)
        except Exception as e:
            logger.error(f"执行电子表格策略时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"发生未知错误: {e}", error=e)

    def _convert_md_to_xlsx(self, md_path: str, template_name: str, progress_callback, cancel_event, temp_dir: str, original_md_path: str) -> Optional[str]:
        """
        将Markdown文件转换为XLSX格式（第一步转换）。
        
        Args:
            md_path: Markdown文件路径（input.md副本）
            template_name: 使用的模板名称
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径，用于存放中间文件
            original_md_path: 原始MD文件路径（用于生成文件名）
            
        Returns:
            Optional[str]: 成功时返回生成的XLSX文件路径，失败时返回None
        """
        # 生成规范的临时输出路径（基于原始MD文件名，不是input.md）
        temp_output_filename = os.path.basename(
            generate_output_path(
                original_md_path, temp_dir, "", True, "fromMd", "xlsx"
            )
        )
        temp_output = os.path.join(temp_dir, temp_output_filename)
        
        # 调用底层转换函数（传递原始源文件路径用于嵌入功能的路径解析）
        output_path = convert_md_to_xlsx(
            md_path,
            temp_output,
            template_name=template_name,
            progress_callback=progress_callback,
            cancel_event=cancel_event,
            original_source_path=original_md_path
        )
        return output_path

    def _finalize_conversion(self, xlsx_path: str, original_md_path: str, progress_callback, cancel_event, temp_dir: str) -> Optional[str]:
        """
        完成最终格式转换（模板方法）。
        
        子类需要实现此方法来完成从XLSX到目标格式的转换。
        
        Args:
            xlsx_path: 中间XLSX文件路径
            original_md_path: 原始Markdown文件路径
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        Returns:
            Optional[str]: 最终文件在临时目录中的完整路径，失败时返回None
            
        Note:
            实现时应在temp_dir中生成文件，返回temp_dir中的文件完整路径
        """
        raise NotImplementedError
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False


@register_conversion(CATEGORY_MARKDOWN, 'xlsx')
class MdToXlsxStrategy(BaseMdToSpreadsheetStrategy):
    """
    将Markdown直接转换为XLSX的策略。
    
    这是最简单的转换策略，因为MD→XLSX已经是基类完成的，
    只需重命名文件即可。
    """
    def _generate_final_path(self, file_path: str) -> str:
        """生成XLSX输出路径"""
        return generate_output_path(file_path, section="", add_timestamp=True, description="fromMd", file_type="xlsx")
    
    def _finalize_conversion(self, xlsx_path: str, original_md_path: str, progress_callback, cancel_event, temp_dir: str) -> Optional[str]:
        """
        重命名XLSX文件为最终输出文件名。
        
        Args:
            xlsx_path: 临时目录中的XLSX文件路径
            original_md_path: 原始MD文件路径
            progress_callback: 进度回调函数（未使用）
            cancel_event: 取消事件对象（未使用）
            temp_dir: 临时目录路径
            
        Returns:
            str: 临时目录中最终文件的完整路径
        """
        final_filename = os.path.basename(
            generate_output_path(original_md_path, section="", add_timestamp=True, description="fromMd", file_type="xlsx")
        )
        final_temp_path = os.path.join(temp_dir, final_filename)
        os.rename(xlsx_path, final_temp_path)
        return final_temp_path


@register_conversion(CATEGORY_MARKDOWN, 'xls')
class MdToXlsStrategy(BaseMdToSpreadsheetStrategy):
    """
    将Markdown通过XLSX转换为XLS的策略。
    
    转换流程：MD → XLSX → XLS
    需要安装WPS或Microsoft Office软件来完成XLSX到XLS的转换。
    """
    def _generate_final_path(self, file_path: str) -> str:
        """生成XLS输出路径"""
        return generate_output_path(file_path, section="", add_timestamp=True, description="fromMd", file_type="xls")
    
    def _finalize_conversion(self, xlsx_path: str, original_md_path: str, progress_callback, cancel_event, temp_dir: str) -> Optional[str]:
        """
        将XLSX文件转换为XLS格式。
        
        Args:
            xlsx_path: 临时目录中的XLSX文件路径
            original_md_path: 原始MD文件路径
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        Returns:
            Optional[str]: 成功时返回临时目录中XLS文件的完整路径，失败时返回None
            
        Raises:
            OfficeSoftwareNotFoundError: 未找到Office或WPS软件时抛出
        """
        if progress_callback: progress_callback("转换为XLS...")
        final_filename = os.path.basename(
            generate_output_path(original_md_path, section="", add_timestamp=True, description="fromMd", file_type="xls")
        )
        temp_output_path = os.path.join(temp_dir, final_filename)
        result_path = convert_xlsx_to_xls(
            xlsx_path, 
            temp_output_path, 
            cancel_event=cancel_event, 
            progress_callback=progress_callback
        )
        return temp_output_path if result_path and os.path.exists(temp_output_path) else None


@register_conversion(CATEGORY_MARKDOWN, 'ods')
class MdToOdsStrategy(BaseMdToSpreadsheetStrategy):
    """
    将Markdown转换为ODS文件的策略。
    
    转换流程：MD → XLSX → ODS
    ODS是OpenDocument电子表格格式，需要安装Microsoft Excel或LibreOffice软件。
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行Markdown到ODS的转换，开始前进行软件可用性预检查。
        """
        # 预检查：ODS格式需要Microsoft Excel或LibreOffice，WPS不支持
        available, error_msg = check_office_availability('ods')
        if not available:
            logger.error(f"ODS转换预检查失败: {error_msg}")
            raise OfficeSoftwareNotFoundError(error_msg)
        
        # 调用父类的execute方法
        return super().execute(file_path, options, progress_callback)
    
    def _generate_final_path(self, file_path: str) -> str:
        """生成ODS输出路径"""
        return generate_output_path(file_path, section="", add_timestamp=True, description="fromMd", file_type="ods")
    
    def _finalize_conversion(self, xlsx_path: str, original_md_path: str, progress_callback, cancel_event, temp_dir: str) -> Optional[str]:
        """
        将XLSX文件转换为ODS格式。
        
        Args:
            xlsx_path: 临时目录中的XLSX文件路径
            original_md_path: 原始MD文件路径
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        Returns:
            Optional[str]: 成功时返回临时目录中ODS文件的完整路径，失败时返回None
            
        Raises:
            OfficeSoftwareNotFoundError: 未找到Office或WPS软件时抛出
        """
        # 预检查：ODS格式需要Microsoft Excel或LibreOffice，WPS不支持
        available, error_msg = check_office_availability('ods')
        if not available:
            logger.error(f"ODS转换预检查失败: {error_msg}")
            raise OfficeSoftwareNotFoundError(error_msg)
        
        if progress_callback: progress_callback("转换为ODS...")
        
        try:
            from gongwen_converter.converter.formats.office import (
                xlsx_to_ods
            )
            
            # 生成最终文件名
            final_filename = os.path.basename(
                generate_output_path(original_md_path, section="", add_timestamp=True, description="fromMd", file_type="ods")
            )
            temp_output_path = os.path.join(temp_dir, final_filename)
            
            # 调用转换函数，传递完整的输出路径
            result_path = xlsx_to_ods(
                xlsx_path, 
                temp_output_path,
                cancel_event=cancel_event
            )
            
            return temp_output_path if result_path and os.path.exists(temp_output_path) else None
                
        except Exception as e:
            logger.error(f"MD转ODS失败: {e}", exc_info=True)
            return None


@register_conversion(CATEGORY_MARKDOWN, 'csv')
class MdToCsvStrategy(BaseStrategy):
    """
    将Markdown通过XLSX转换为CSV的策略。
    
    转换流程：MD → XLSX → CSV
    使用转换链方式，先转为XLSX，再将XLSX转为CSV。
    
    输出结构：所有CSV文件都放在一个子文件夹内
    ```
    document_20251107_201500_fromMd/
    ├── document_Sheet1_20251107_201500_fromMd.csv
    ├── document_Sheet2_20251107_201500_fromMd.csv
    └── document_Sheet3_20251107_201500_fromMd.csv
    ```
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行Markdown到CSV的转换
        
        Args:
            file_path: 输入的Markdown文件路径
            options: 转换选项字典，包含：
                - template_name: (必需) 使用的模板名称
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            
        Note:
            output_path 返回子文件夹中第一个CSV文件的路径，
            以便在批量模式下正确定位到转换结果
        """
        if options is None:
            options = {}
            
        cancel_event = options.get("cancel_event")
        template_name = options.get("template_name")
        
        if not template_name:
            return ConversionResult(success=False, message="错误：未提供模板名称。")
        
        try:
            if progress_callback:
                progress_callback("正在将Markdown转换为XLSX...")
            
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用临时目录管理输出
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤0：创建输入副本 input.md
                from gongwen_converter.utils.workspace_manager import prepare_input_file
                temp_md = prepare_input_file(file_path, temp_dir, 'md')
                logger.debug(f"已创建输入副本: {os.path.basename(temp_md)}")
                
                # 步骤1：MD → XLSX (在临时目录中，使用副本，生成规范文件名)
                # 使用原始文件路径生成文件名，而不是input.md
                temp_xlsx_filename = os.path.basename(
                    generate_output_path(
                        file_path, temp_dir, "", True, "fromMd", "xlsx"
                    )
                )
                temp_xlsx = os.path.join(temp_dir, temp_xlsx_filename)
                temp_xlsx_path = convert_md_to_xlsx(
                    temp_md,
                    temp_xlsx,
                    template_name=template_name,
                    progress_callback=progress_callback,
                    cancel_event=cancel_event,
                    original_source_path=file_path
                )
                
                if not temp_xlsx_path:
                    return ConversionResult(success=False, message="Markdown到XLSX转换失败。")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成统一basename和时间戳描述部分
                if progress_callback:
                    progress_callback("转换为CSV...")
                
                base_path = generate_output_path(
                    file_path,
                    section="",
                    add_timestamp=True,
                    description="fromMd",
                    file_type="csv"
                )
                basename = os.path.splitext(os.path.basename(base_path))[0]
                # 例如: "document_20251107_201500_fromMd"
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
                # basename格式: "原名_时间戳_描述"
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
                
                # 步骤4：调用底层转换函数，输出到临时子文件夹
                # 传递原始文件名和统一的时间戳描述，确保CSV文件名与子文件夹名一致
                from gongwen_converter.converter.table_converters import convert_xlsx_to_csv
                
                csv_files = convert_xlsx_to_csv(
                    temp_xlsx_path,
                    actual_format=None,
                    output_dir=temp_output_folder,  # 传递临时子文件夹
                    original_basename=original_file_basename,  # 传递原始文件basename
                    unified_timestamp_desc=unified_timestamp_desc  # 传递统一的时间戳描述
                )
                
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
            logger.error(f"MD转CSV失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )
