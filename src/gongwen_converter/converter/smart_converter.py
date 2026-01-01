"""
智能转换链核心模块
使用规则表驱动的方式，自动规划单步或多步转换路径

核心特性:
- 使用真实格式进行转换路径规划（不依赖文件扩展名）
- 自动处理扩展名不匹配的文件
- 支持单步和多步转换
- 中间文件自动管理和清理
"""

import os
import logging
import tempfile
import shutil
from typing import Optional, Callable, List
import threading

# 导入并重新导出异常类，使其成为模块的公共API
from gongwen_converter.converter.formats.office import OfficeSoftwareNotFoundError, check_office_availability

logger = logging.getLogger(__name__)


# 转换规则表
CONVERSION_RULES = {
    'spreadsheet': {
        'hub': 'xlsx',  # 中心格式
        'formats': ['xlsx', 'xls', 'ods', 'csv', 'et'],
    },
    'document': {
        'hub': 'docx',  # 中心格式
        'formats': ['docx', 'doc', 'odt', 'rtf', 'wps'],
        # 注意：WPS 格式可以作为输入（通过 office_to_docx），但不作为输出格式
    }
}


class SmartConverter:
    """
    智能转换链：自动规划单步或多步转换路径
    
    核心特性：
    - 规则表驱动的转换路径规划
    - 自动处理 ODT/ODS 格式的 WPS 限制
    - 支持扩展名不匹配的文件
    - 中间文件自动管理和清理
    """
    
    def __init__(self):
        """初始化智能转换器"""
        pass
    
    def convert(
        self,
        input_path: str,
        target_format: str,
        category: str,
        actual_format: str,
        output_dir: Optional[str] = None,
        cancel_event: Optional[threading.Event] = None,
        progress_callback: Optional[Callable[[str], None]] = None,
        preferred_software: Optional[str] = None
    ) -> Optional[str]:
        """
        执行智能转换
        
        参数:
            input_path: 输入文件路径
            target_format: 目标格式（如 'ods', 'odt'）
            category: 文件类别（'spreadsheet' 或 'document'）
            actual_format: 文件的真实格式（必需），用于转换路径规划
                          使用此参数而非从文件名推断，确保扩展名不匹配的文件也能正确转换
            output_dir: 输出目录（可选）
            cancel_event: 取消事件
            progress_callback: 进度回调
            preferred_software: 用户偏好的Office软件
            
        返回:
            成功时返回输出文件路径，失败时返回 None
            
        说明:
            智能转换链会根据源格式和目标格式自动规划最优转换路径：
            - 单步转换：源或目标是中心格式时（如 xlsx→ods）
            - 多步转换：需要经过中心格式时（如 xls→xlsx→ods）
        """
        try:
            # 使用传入的真实格式作为源格式
            # 这确保即使文件扩展名被修改（如xls文件改名为.doc），也能正确转换
            source_format = actual_format
            logger.info(f"智能转换链: {source_format} → {target_format}")
            logger.debug(f"输入文件: {input_path}, 真实格式: {actual_format}")
            
            # 预检查：对于需要Office软件的目标格式，先检查软件可用性
            # 这样可以在转换开始前就告知用户软件不可用，避免浪费时间
            available, error_msg = check_office_availability(target_format)
            if not available:
                logger.error(f"预检查失败: {error_msg}")
                raise OfficeSoftwareNotFoundError(error_msg)
            
            # 规划转换路径
            conversion_path = self._plan_conversion_path(
                source_format, 
                target_format, 
                category
            )
            
            if not conversion_path:
                logger.error(f"无法规划转换路径: {source_format} → {target_format}")
                return None
            
            logger.info(f"规划的转换路径: {source_format} → {' → '.join(conversion_path)}")
            
            # 执行转换
            if len(conversion_path) == 1:
                # 单步转换
                return self._convert_single_step(
                    input_path,
                    source_format,
                    target_format,
                    output_dir,
                    cancel_event,
                    progress_callback,
                    preferred_software
                )
            else:
                # 多步转换
                return self._convert_multi_step(
                    input_path,
                    source_format,
                    conversion_path,
                    output_dir,
                    cancel_event,
                    progress_callback,
                    preferred_software
                )
        
        except Exception as e:
            logger.error(f"智能转换失败: {e}", exc_info=True)
            return None
    
    def _plan_conversion_path(
        self, 
        source_format: str, 
        target_format: str, 
        category: str
    ) -> List[str]:
        """
        规划转换路径
        
        规则：
        - 如果源或目标是中心格式，返回单步转换
        - 否则返回两步转换：源 → 中心 → 目标
        
        参数:
            source_format: 源格式
            target_format: 目标格式
            category: 文件类别
            
        返回:
            转换路径列表，例如 ['xlsx', 'ods'] 或 ['xlsx']
        """
        rules = CONVERSION_RULES.get(category)
        if not rules:
            logger.error(f"未知的文件类别: {category}")
            return []
        
        hub_format = rules['hub']
        supported_formats = rules['formats']
        
        # 检查格式是否支持
        if source_format not in supported_formats or target_format not in supported_formats:
            logger.error(f"不支持的格式: {source_format} → {target_format}")
            return []
        
        # 如果源或目标是中心格式，单步转换
        if source_format == hub_format or target_format == hub_format:
            return [target_format]
        
        # 否则两步转换：源 → 中心 → 目标
        return [hub_format, target_format]
    
    def _should_exclude_wps(self, source_format: str, target_format: str) -> bool:
        """
        判断此转换步骤是否需要排除 WPS
        
        规则：仅当源或目标是 ODS/ODT 时，该步骤 exclude_wps=True
        """
        ods_odt_formats = {'ods', 'odt'}
        return source_format in ods_odt_formats or target_format in ods_odt_formats
    
    def _convert_single_step(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: Optional[str],
        cancel_event: Optional[threading.Event],
        progress_callback: Optional[Callable[[str], None]],
        preferred_software: Optional[str]
    ) -> Optional[str]:
        """
        执行单步转换
        
        参数:
            source_format: 源文件的真实格式
        """
        if cancel_event and cancel_event.is_set():
            logger.info("转换被取消")
            return None
        
        if progress_callback:
            progress_callback(f"转换中: {source_format.upper()} → {target_format.upper()}...")
        
        # 判断是否需要排除 WPS
        exclude_wps = self._should_exclude_wps(source_format, target_format)
        
        # 执行转换
        return self._execute_conversion(
            input_path,
            source_format,
            target_format,
            output_dir,
            cancel_event,
            exclude_wps,
            preferred_software
        )
    
    def _convert_multi_step(
        self,
        input_path: str,
        source_format: str,
        conversion_path: List[str],
        output_dir: Optional[str],
        cancel_event: Optional[threading.Event],
        progress_callback: Optional[Callable[[str], None]],
        preferred_software: Optional[str]
    ) -> Optional[str]:
        """
        执行多步转换（统一的中间文件保留逻辑）
        
        参数:
            source_format: 原始输入文件的真实格式
            
        说明:
            - 所有中间文件都先在临时目录处理
            - 根据配置决定是否保留中间文件到输出目录
            - 与版式文件转换策略保持一致的逻辑
        """
        # 创建临时目录用于中间文件
        with tempfile.TemporaryDirectory() as temp_dir:
            current_file = input_path
            current_format = source_format  # 使用传入的真实格式
            original_format = source_format  # 保存原始格式，用于最终文件的description
            intermediate_files = []  # 记录中间文件
            
            # 执行每一步转换
            for step_index, intermediate_format in enumerate(conversion_path, 1):
                if cancel_event and cancel_event.is_set():
                    logger.info("转换被取消")
                    return None
                
                is_last_step = (step_index == len(conversion_path))
                
                if progress_callback:
                    progress_callback(
                        f"步骤 {step_index}/{len(conversion_path)}: "
                        f"{current_format.upper()} → {intermediate_format.upper()}..."
                    )
                
                # 判断当前步骤是否需要排除 WPS
                exclude_wps = self._should_exclude_wps(current_format, intermediate_format)
                
                # 所有步骤都输出到临时目录
                step_output_dir = temp_dir
                
                # 执行转换（传递当前格式和原始格式）
                # 最后一步使用原始格式生成description
                result_file = self._execute_conversion(
                    current_file,
                    current_format,
                    intermediate_format,
                    step_output_dir,
                    cancel_event,
                    exclude_wps,
                    preferred_software,
                    original_format=original_format if is_last_step else None
                )
                
                if not result_file:
                    logger.error(f"步骤 {step_index} 转换失败")
                    return None
                
                logger.info(f"步骤 {step_index}/{len(conversion_path)} 完成: {os.path.basename(result_file)}")
                
                # 记录中间文件（除了最后一步）
                if not is_last_step:
                    intermediate_files.append(result_file)
                    logger.debug(f"记录中间文件: {os.path.basename(result_file)}")
                
                # 更新当前文件和格式
                # 第二步的源格式是第一步的目标格式（intermediate_format）
                current_file = result_file
                current_format = intermediate_format
            
            # 转换完成，处理文件移动
            final_file = current_file
            
            if not output_dir:
                # 如果没有指定输出目录，使用输入文件目录
                output_dir = os.path.dirname(input_path)
            
            # 检查最后一步是否是CSV转换
            final_format = conversion_path[-1]
            if final_format == 'csv':
                # CSV转换：移动整个子文件夹到output_dir
                csv_subfolder = os.path.dirname(final_file)  # 获取CSV文件所在的子文件夹
                subfolder_name = os.path.basename(csv_subfolder)
                
                # 构建目标文件夹路径
                target_subfolder = os.path.join(output_dir, subfolder_name)
                
                # 如果目标文件夹已存在，先删除
                if os.path.exists(target_subfolder):
                    shutil.rmtree(target_subfolder)
                    logger.debug(f"已删除现有文件夹: {target_subfolder}")
                
                # 移动整个子文件夹到output_dir
                final_subfolder = shutil.move(csv_subfolder, output_dir)
                logger.info(f"CSV文件夹已移动到: {final_subfolder}")
                
                # 返回output_dir中的文件路径
                final_output_path = os.path.join(final_subfolder, os.path.basename(final_file))
                logger.debug(f"返回CSV文件路径: {final_output_path}")
            else:
                # 非CSV转换：移动单个文件到output_dir
                final_output_path = os.path.join(output_dir, os.path.basename(final_file))
                shutil.move(final_file, final_output_path)
                logger.info(f"最终文件已移动: {os.path.basename(final_output_path)}")
            
            # 检查配置决定是否保留中间文件（仅对非CSV转换）
            if final_format != 'csv':
                should_keep = self._should_keep_intermediates()
                if should_keep and intermediate_files:
                    logger.info(f"保留 {len(intermediate_files)} 个中间文件到输出目录")
                    for intermediate_file in intermediate_files:
                        if os.path.exists(intermediate_file):
                            dest_path = os.path.join(output_dir, os.path.basename(intermediate_file))
                            shutil.move(intermediate_file, dest_path)
                            logger.info(f"保留中间文件: {os.path.basename(dest_path)}")
                else:
                    if intermediate_files:
                        logger.debug(f"清理 {len(intermediate_files)} 个中间文件")
            else:
                # CSV转换：中间文件留在temp_dir，由调用者决定是否保留
                logger.debug(f"CSV转换：中间文件保留在临时目录，由调用者处理")
            
            return final_output_path
    
    def _execute_conversion(
        self,
        input_path: str,
        source_format: str,
        target_format: str,
        output_dir: Optional[str],
        cancel_event: Optional[threading.Event],
        exclude_wps: bool,
        preferred_software: Optional[str],
        original_format: Optional[str] = None
    ) -> Optional[str]:
        """
        执行实际的格式转换
        
        参数:
            input_path: 输入文件路径
            source_format: 源文件的真实格式（不从文件名推断）
            target_format: 目标格式
            output_dir: 输出目录
            cancel_event: 取消事件
            exclude_wps: 是否排除 WPS
            preferred_software: 用户偏好的Office软件
            original_format: 原始文件格式（用于多步转换中生成最终description）
            
        返回:
            转换后的文件路径，失败时返回 None
            
        说明:
            使用传入的source_format而不是从文件名推断，
            确保多步转换时每一步都使用正确的格式。
            当original_format不为None时，使用它来生成description，
            这样多步转换的最终文件能显示真实的原始格式。
        """
        from gongwen_converter.converter.formats.office import (
            office_to_xlsx,
            xlsx_to_ods,
            ods_to_xlsx,
            office_to_docx,
            docx_to_odt,
            odt_to_docx,
            rtf_to_docx,
            convert_xlsx_to_xls,
            convert_docx_to_doc,
            convert_docx_to_rtf,
            OfficeSoftwareNotFoundError
        )
        from gongwen_converter.utils.path_utils import generate_output_path
        
        logger.debug(f"执行转换: {source_format} → {target_format}, 文件: {os.path.basename(input_path)}")
        
        # 确定用于description的格式：如果提供了原始格式，使用它；否则使用当前源格式
        format_for_description = original_format if original_format else source_format
        
        try:
            # 根据源格式和目标格式选择转换函数
            # 表格格式转换
            if source_format in ['xls', 'et'] and target_format == 'xlsx':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="xlsx"
                )
                return office_to_xlsx(
                    input_path=input_path,
                    output_path=output_path,
                    actual_format=source_format,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'xlsx' and target_format == 'ods':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="ods"
                )
                return xlsx_to_ods(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'ods' and target_format == 'xlsx':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="xlsx"
                )
                return ods_to_xlsx(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'xlsx' and target_format == 'xls':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="xls"
                )
                return convert_xlsx_to_xls(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            # CSV 转换支持
            elif source_format == 'xlsx' and target_format == 'csv':
                from gongwen_converter.converter.table_converters import convert_xlsx_to_csv
                
                # 生成统一的basename（清理时间戳）
                # 如果提供了original_format，说明是多步转换的最后一步，使用它构建description
                # 否则，说明是单步转换，使用当前source_format
                desc_format = original_format if original_format else source_format
                
                # 生成基础输出路径以获取清理后的basename和时间戳
                base_output = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{desc_format.capitalize()}",
                    file_type="csv"
                )
                
                # 提取basename作为子文件夹名（包含完整的时间戳和描述）
                folder_basename = os.path.splitext(os.path.basename(base_output))[0]
                logger.debug(f"CSV转换 - 子文件夹名: {folder_basename}")
                
                # 提取清理后的basename（不含扩展名、时间戳、描述）
                base_filename = os.path.basename(base_output)
                parts = base_filename.replace('.csv', '').split('_')
                
                # 找到时间戳位置，提取原始basename
                timestamp_idx = None
                for i, part in enumerate(parts):
                    if len(part) == 8 and part.isdigit():
                        timestamp_idx = i
                        break
                
                if timestamp_idx is not None:
                    original_basename = '_'.join(parts[:timestamp_idx])
                    unified_timestamp_desc = '_'.join(parts[timestamp_idx:])
                else:
                    # 降级方案
                    original_basename = parts[0] if parts else "output"
                    unified_timestamp_desc = '_'.join(parts[1:]) if len(parts) > 1 else ""
                
                logger.debug(f"CSV转换 - 原始basename: {original_basename}, 时间戳描述: {unified_timestamp_desc}")
                
                # 创建子文件夹路径
                csv_folder = os.path.join(output_dir, folder_basename)
                os.makedirs(csv_folder, exist_ok=True)
                logger.debug(f"创建CSV子文件夹: {csv_folder}")
                
                # XLSX → CSV：传递子文件夹路径、原始basename和统一时间戳
                csv_files = convert_xlsx_to_csv(
                    xlsx_path=input_path,
                    actual_format='xlsx',
                    output_dir=csv_folder,  # 输出到子文件夹
                    original_basename=original_basename,
                    unified_timestamp_desc=unified_timestamp_desc
                )
                if csv_files and len(csv_files) > 0:
                    # 返回第一个CSV文件的路径（在子文件夹中）
                    return csv_files[0]
                else:
                    logger.error("XLSX转CSV失败：未生成CSV文件")
                    return None
            
            elif source_format == 'csv' and target_format == 'xlsx':
                from gongwen_converter.converter.table_converters import convert_csv_to_xlsx
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="xlsx"
                )
                return convert_csv_to_xlsx(
                    csv_path=input_path,
                    output_path=output_path
                )
            
            # 文档格式转换
            elif source_format in ['doc', 'wps'] and target_format == 'docx':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="docx"
                )
                return office_to_docx(
                    input_path=input_path,
                    output_path=output_path,
                    actual_format=source_format,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'rtf' and target_format == 'docx':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="docx"
                )
                return rtf_to_docx(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'docx' and target_format == 'odt':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="odt"
                )
                return docx_to_odt(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'odt' and target_format == 'docx':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="docx"
                )
                return odt_to_docx(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'docx' and target_format == 'doc':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="doc"
                )
                return convert_docx_to_doc(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'docx' and target_format == 'rtf':
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="rtf"
                )
                return convert_docx_to_rtf(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            elif source_format == 'xlsx' and target_format == 'et':
                # ET 格式实际上使用 XLS 格式
                output_path = generate_output_path(
                    input_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{format_for_description.capitalize()}",
                    file_type="et"
                )
                return convert_xlsx_to_xls(
                    input_path=input_path,
                    output_path=output_path,
                    cancel_event=cancel_event
                )
            
            else:
                logger.error(f"不支持的转换: {source_format} → {target_format}")
                return None
        
        except OfficeSoftwareNotFoundError as e:
            logger.error(f"Office软件未找到: {e}")
            return None
        except Exception as e:
            logger.error(f"转换执行失败: {e}", exc_info=True)
            return None
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件（与其他策略保持一致）"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False
