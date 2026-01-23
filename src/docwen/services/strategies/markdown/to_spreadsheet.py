"""
Markdown 转电子表格策略模块

将 Markdown 文件转换为各种电子表格格式（XLSX、XLS、ODS、CSV）。

转换流程：
1. MD → XLSX（使用模板）
2. 可选：XLSX → 其他表格格式（XLS/ODS/CSV）

依赖：
- converter.md2xlsx: MD → XLSX 核心转换
- converter.formats.spreadsheet: XLSX → 其他表格格式
"""
import os
import re
import shutil
import logging
import tempfile
from typing import Dict, Any, Callable, Optional

from docwen.services.result import ConversionResult
from docwen.services.strategies.base_strategy import BaseStrategy
from docwen.services.strategies import register_conversion, CATEGORY_MARKDOWN
from docwen.utils.path_utils import generate_output_path
from docwen.config.config_manager import config_manager
from docwen.i18n import t
from docwen.converter.md2xlsx.core import convert as convert_md_to_xlsx
from docwen.converter.formats.spreadsheet import (
    xlsx_to_xls,
    xlsx_to_ods,
)
from docwen.converter.formats.common import (
    OfficeSoftwareNotFoundError,
    check_office_availability,
)

logger = logging.getLogger(__name__)


class BaseMdToSpreadsheetStrategy(BaseStrategy):
    """
    MD转电子表格系列策略的基类
    
    提供了将Markdown文件转换为电子表格的通用流程：
    1. 将MD转换为XLSX（中间格式）
    2. 可选：将XLSX进一步转换为其他格式（XLS、ODS、CSV等）
    3. 根据配置决定是否保留中间文件
    
    子类需要实现：
    - _finalize_conversion: 完成从XLSX到目标格式的转换
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行Markdown到电子表格的转换
        
        参数:
            file_path: 输入的Markdown文件路径
            options: 转换选项字典，包含：
                - template_name: (必需) 使用的模板名称
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        返回:
            ConversionResult: 包含转换结果的对象
        """
        if options is None:
            options = {}
            
        cancel_event = options.get("cancel_event")
        update_progress = progress_callback or (lambda msg: None)

        try:
            template_name = options.get("template_name")
            if not template_name:
                return ConversionResult(success=False, message=t('conversion.progress.error_no_template'))

            from docwen.utils.workspace_manager import prepare_input_file, get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤0: 创建输入副本 input.md
                temp_md = prepare_input_file(file_path, temp_dir, 'md')
                logger.debug(f"已创建输入副本: {os.path.basename(temp_md)}")
                
                # 步骤 1: MD -> XLSX (在临时目录中，使用副本)
                intermediate_xlsx_path = self._convert_md_to_xlsx(
                    temp_md, template_name, progress_callback, cancel_event, temp_dir, file_path
                )
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                if not intermediate_xlsx_path:
                    return ConversionResult(success=False, message=t('conversion.messages.md_to_xlsx_failed'))

                # 步骤 2: (可选) XLSX -> 最终格式 (在临时目录中)
                update_progress(t('conversion.progress.final_conversion'))
                final_temp_path = self._finalize_conversion(
                    intermediate_xlsx_path, file_path, progress_callback, cancel_event, temp_dir
                )
                if cancel_event and cancel_event.is_set(): 
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                if not (final_temp_path and os.path.exists(final_temp_path)):
                    return ConversionResult(success=False, message=t('conversion.messages.final_format_generation_failed'))

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

            target_fmt = os.path.splitext(final_output_path)[1].lstrip('.').upper()
            return ConversionResult(
                success=True, 
                output_path=final_output_path, 
                message=t('conversion.messages.conversion_to_format_success', format=target_fmt)
            )
        
        except OfficeSoftwareNotFoundError as e:
            logger.error(f"Office/WPS软件未找到: {e}")
            # 根据具体策略类提供精确的错误提示
            if isinstance(self, MdToXlsStrategy):
                return ConversionResult(
                    success=False, 
                    message=t('conversion.messages.xls_format_requires_office'), 
                    error=e
                )
            else:
                return ConversionResult(success=False, message=str(e), error=e)
        except Exception as e:
            logger.error(f"执行电子表格策略时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"发生未知错误: {e}", error=e)

    def _convert_md_to_xlsx(
        self,
        md_path: str,
        template_name: str,
        progress_callback,
        cancel_event,
        temp_dir: str,
        original_md_path: str
    ) -> Optional[str]:
        """
        将Markdown文件转换为XLSX格式（第一步转换）
        
        参数:
            md_path: Markdown文件路径（input.md副本）
            template_name: 使用的模板名称
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径，用于存放中间文件
            original_md_path: 原始MD文件路径（用于生成文件名）
            
        返回:
            成功时返回生成的XLSX文件路径，失败时返回None
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

    def _finalize_conversion(
        self,
        xlsx_path: str,
        original_md_path: str,
        progress_callback,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """
        完成最终格式转换（模板方法）
        
        子类需要实现此方法来完成从XLSX到目标格式的转换。
        
        参数:
            xlsx_path: 中间XLSX文件路径
            original_md_path: 原始Markdown文件路径
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        返回:
            最终文件在临时目录中的完整路径，失败时返回None
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
    将Markdown直接转换为XLSX的策略
    
    这是最简单的转换策略，因为MD→XLSX已经是基类完成的，
    只需重命名文件即可。
    """
    
    def _generate_final_path(self, file_path: str) -> str:
        """生成XLSX输出路径"""
        return generate_output_path(
            file_path, section="", add_timestamp=True, 
            description="fromMd", file_type="xlsx"
        )
    
    def _finalize_conversion(
        self,
        xlsx_path: str,
        original_md_path: str,
        progress_callback,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """重命名XLSX文件为最终输出文件名"""
        final_filename = os.path.basename(
            generate_output_path(
                original_md_path, section="", add_timestamp=True, 
                description="fromMd", file_type="xlsx"
            )
        )
        final_temp_path = os.path.join(temp_dir, final_filename)
        os.rename(xlsx_path, final_temp_path)
        return final_temp_path


@register_conversion(CATEGORY_MARKDOWN, 'xls')
class MdToXlsStrategy(BaseMdToSpreadsheetStrategy):
    """
    将Markdown通过XLSX转换为XLS的策略
    
    转换流程：MD → XLSX → XLS
    需要安装WPS或Microsoft Office软件来完成XLSX到XLS的转换。
    """
    
    def _generate_final_path(self, file_path: str) -> str:
        """生成XLS输出路径"""
        return generate_output_path(
            file_path, section="", add_timestamp=True, 
            description="fromMd", file_type="xls"
        )
    
    def _finalize_conversion(
        self,
        xlsx_path: str,
        original_md_path: str,
        progress_callback,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """将XLSX文件转换为XLS格式"""
        if progress_callback: 
            progress_callback(t('conversion.progress.converting_to_format', format='XLS'))
        final_filename = os.path.basename(
            generate_output_path(
                original_md_path, section="", add_timestamp=True, 
                description="fromMd", file_type="xls"
            )
        )
        temp_output_path = os.path.join(temp_dir, final_filename)
        result_path = xlsx_to_xls(
            xlsx_path, 
            temp_output_path, 
            cancel_event=cancel_event, 
            progress_callback=progress_callback
        )
        return temp_output_path if result_path and os.path.exists(temp_output_path) else None


@register_conversion(CATEGORY_MARKDOWN, 'ods')
class MdToOdsStrategy(BaseMdToSpreadsheetStrategy):
    """
    将Markdown转换为ODS文件的策略
    
    转换流程：MD → XLSX → ODS
    ODS是OpenDocument电子表格格式，需要安装Microsoft Excel或LibreOffice软件。
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """执行Markdown到ODS的转换，开始前进行软件可用性预检查"""
        # 预检查：ODS格式需要Microsoft Excel或LibreOffice，WPS不支持
        available, error_msg = check_office_availability('ods')
        if not available:
            logger.error(f"ODS转换预检查失败: {error_msg}")
            raise OfficeSoftwareNotFoundError(error_msg)
        
        # 调用父类的execute方法
        return super().execute(file_path, options, progress_callback)
    
    def _generate_final_path(self, file_path: str) -> str:
        """生成ODS输出路径"""
        return generate_output_path(
            file_path, section="", add_timestamp=True, 
            description="fromMd", file_type="ods"
        )
    
    def _finalize_conversion(
        self,
        xlsx_path: str,
        original_md_path: str,
        progress_callback,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """将XLSX文件转换为ODS格式"""
        # 预检查：ODS格式需要Microsoft Excel或LibreOffice，WPS不支持
        available, error_msg = check_office_availability('ods')
        if not available:
            logger.error(f"ODS转换预检查失败: {error_msg}")
            raise OfficeSoftwareNotFoundError(error_msg)
        
        if progress_callback: 
            progress_callback(t('conversion.progress.converting_to_format', format='ODS'))
        
        try:
            # 生成最终文件名
            final_filename = os.path.basename(
                generate_output_path(
                    original_md_path, section="", add_timestamp=True, 
                    description="fromMd", file_type="ods"
                )
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
    将Markdown通过XLSX转换为CSV的策略
    
    转换流程：MD → XLSX → CSV
    使用转换链方式，先转为XLSX，再将XLSX转为CSV。
    
    输出结构：所有CSV文件都放在一个子文件夹内
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行Markdown到CSV的转换
        
        参数:
            file_path: 输入的Markdown文件路径
            options: 转换选项字典，包含：
                - template_name: (必需) 使用的模板名称
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        返回:
            ConversionResult: 包含转换结果的对象
        """
        if options is None:
            options = {}
            
        cancel_event = options.get("cancel_event")
        template_name = options.get("template_name")
        
        if not template_name:
            return ConversionResult(success=False, message=t('conversion.progress.error_no_template'))
        
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.converting_to_format', format='XLSX'))
            
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用临时目录管理输出
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤0：创建输入副本 input.md
                from docwen.utils.workspace_manager import prepare_input_file
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
                    return ConversionResult(success=False, message=t('conversion.messages.md_to_xlsx_failed'))
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 步骤2：生成统一basename和时间戳描述部分
                if progress_callback:
                    progress_callback(t('conversion.progress.converting_to_format', format='CSV'))
                
                base_path = generate_output_path(
                    file_path,
                    section="",
                    add_timestamp=True,
                    description="fromMd",
                    file_type="csv"
                )
                basename = os.path.splitext(os.path.basename(base_path))[0]
                logger.debug(f"统一basename: {basename}")
                
                # 提取原始文件名（不含扩展名、时间戳、描述）
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
                
                # 步骤4：调用底层转换函数，输出到临时子文件夹
                from docwen.converter.formats.spreadsheet import xlsx_to_csv
                
                csv_files = xlsx_to_csv(
                    temp_xlsx_path,
                    actual_format=None,
                    output_dir=temp_output_folder,
                    original_basename=original_file_basename,
                    unified_timestamp_desc=unified_timestamp_desc
                )
                
                if not csv_files:
                    return ConversionResult(
                        success=False,
                        message=t('conversion.messages.no_csv_generated')
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
                progress_callback(t('conversion.progress.csv_completed', count=len(csv_files)))
            
            # 返回第一个CSV文件路径用于定位（现在在子文件夹内）
            return ConversionResult(
                success=True,
                output_path=output_path,
                message=t('conversion.progress.conversion_to_csv_success', count=len(csv_files))
            )
            
        except Exception as e:
            logger.error(f"MD转CSV失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )
