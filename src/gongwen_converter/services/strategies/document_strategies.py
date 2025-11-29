"""
处理 DOCX 文件转换和校对的策略集合。
"""
import os
import logging
from typing import Dict, Any, Callable, Optional

from gongwen_converter.services.result import ConversionResult
from gongwen_converter.services.strategies.base_strategy import BaseStrategy
from . import register_conversion, register_action, CATEGORY_DOCUMENT
from gongwen_converter.utils.path_utils import generate_output_path
import tempfile
import shutil

# 导入核心转换和处理函数
from gongwen_converter.converter.docx2md.core import convert_docx_to_md
from gongwen_converter.docx_spell.core import process_docx

logger = logging.getLogger(__name__)


def _preprocess_document_file(file_path: str, temp_dir: str, cancel_event=None, actual_format: str = None) -> str:
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
        from gongwen_converter.utils.file_type_utils import detect_actual_file_format
        actual_format = detect_actual_file_format(file_path)
        logger.debug(f"自动检测文档文件格式: {actual_format}")
    else:
        logger.debug(f"使用传入的文件格式: {actual_format}")
    
    # 步骤1：无论什么格式，都先创建输入副本 input.{ext}
    from gongwen_converter.utils.workspace_manager import prepare_input_file
    temp_input = prepare_input_file(file_path, temp_dir, actual_format)
    logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
    
    # 步骤2：如果是标准格式（DOCX），直接返回副本路径
    if actual_format == 'docx':
        logger.debug(f"文件已是DOCX格式，返回副本路径")
        return temp_input
    
    # 步骤3：需要转换的格式，从副本转换为DOCX
    if actual_format in ['doc', 'wps', 'rtf', 'odt']:
        logger.info(f"检测到{actual_format.upper()}格式，从副本转换为DOCX: {os.path.basename(temp_input)}")
        
        try:
            from gongwen_converter.converter.formats.office import (
                office_to_docx, rtf_to_docx, odt_to_docx
            )
            
            # 生成目标文件路径
            output_filename = generate_output_path(
                file_path,
                section="",
                add_timestamp=True,
                description=f'from{actual_format.capitalize()}',
                file_type='docx'
            )
            output_path = os.path.join(temp_dir, os.path.basename(output_filename))
            
            # 根据格式选择转换函数，使用副本作为输入
            if actual_format == 'rtf':
                converted_path = rtf_to_docx(
                    temp_input,  # 使用副本
                    output_path,
                    cancel_event=cancel_event
                )
            elif actual_format == 'odt':
                converted_path = odt_to_docx(
                    temp_input,  # 使用副本
                    output_path,
                    cancel_event=cancel_event
                )
            else:  # doc 或 wps
                converted_path = office_to_docx(
                    temp_input,  # 使用副本
                    output_path,
                    actual_format=actual_format,
                    cancel_event=cancel_event
                )
            
            if converted_path:
                logger.info(f"{actual_format.upper()}转DOCX成功: {os.path.basename(converted_path)}")
                return converted_path
            else:
                logger.error("格式转换失败，返回None")
                raise RuntimeError(f"{actual_format.upper()}转DOCX失败")
                
        except Exception as e:
            logger.error(f"{actual_format.upper()}转DOCX失败: {e}")
            raise RuntimeError(f"{actual_format.upper()}转DOCX失败: {e}")
    
    # 步骤4：其他不支持的格式，返回副本路径（尝试直接处理）
    logger.warning(f"不支持的文档格式: {actual_format}，返回副本路径尝试处理")
    return temp_input

@register_conversion(CATEGORY_DOCUMENT, 'md')
class DocxToMdStrategy(BaseStrategy):
    """
    将DOCX文件转换为Markdown文件的策略。
    
    功能特性：
    - 自动识别主要部分和附件部分
    - 主要部分和附件部分分别生成独立的Markdown文件
    - 支持取消操作
    - 带YAML元数据头部
    - **支持DOC/WPS/RTF格式自动转换**
    
    支持的输入格式：
    - DOCX：直接处理
    - DOC/WPS/RTF：自动转换为DOCX后处理
    
    输出文件：
    - 主要部分：文件名标记为 "主要部分"
    - 附件部分（如果存在）：文件名标记为 "附件部分"
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行DOCX到Markdown的转换（支持DOC/WPS/RTF自动转换）。
        
        Args:
            file_path: 输入的文档文件路径（DOCX/DOC/WPS/RTF）
            options: 转换选项字典，包含：
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            - success: 转换是否成功
            - output_path: 主要部分的Markdown文件路径
            - message: 成功或失败的描述信息
            - error: 失败时的错误对象
            
        Note:
            - 如果输入是DOC/WPS/RTF，会先自动转换为DOCX
            - 转换过程在临时目录中进行
            - 如果DOCX包含附件部分，将额外生成附件的Markdown文件
        """
        options = options or {}
        
        # 调试日志：查看收到的options
        logger.info(f"DocxToMdStrategy收到options: {options}")
        logger.info(f"  extract_image: {options.get('extract_image')}")
        logger.info(f"  extract_ocr: {options.get('extract_ocr')}")
        
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")  # 从options中提取actual_format
        
        if progress_callback:
            progress_callback("准备转换...")
        
        try:
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 将DOC/WPS/RTF转换为DOCX（如需要）
                if progress_callback:
                    progress_callback("检测文件格式...")
                
                processed_file = _preprocess_document_file(file_path, temp_dir, cancel_event, actual_format)
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成统一basename和创建临时子文件夹
                # 根据actual_format生成description
                description = f"from{actual_format.capitalize()}" if actual_format else "fromDocx"
                
                # 生成主要部分文件的完整路径（用于提取basename）
                base_path = generate_output_path(
                    file_path,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="md"
                )
                
                # 提取basename（不含.md扩展名）
                basename = os.path.splitext(os.path.basename(base_path))[0]
                # 例如: "document_20231107_204530_fromDocx"
                
                # 创建临时子文件夹
                temp_output_folder = os.path.join(temp_dir, basename)
                os.makedirs(temp_output_folder, exist_ok=True)
                logger.debug(f"创建临时子文件夹: {temp_output_folder}")
                
                # 步骤3：从GUI获取导出选项
                extract_image = options.get('extract_image', True)
                extract_ocr = options.get('extract_ocr', False)
                
                logger.info(f"从options提取参数 - extract_image: {extract_image}, extract_ocr: {extract_ocr}")
                
                # 步骤4：调用核心转换函数，解析DOCX结构并转换为Markdown
                if progress_callback:
                    progress_callback("转换中...")
                
                result = convert_docx_to_md(
                    processed_file,
                    extract_image=extract_image,
                    extract_ocr=extract_ocr,
                    progress_callback=progress_callback,
                    cancel_event=cancel_event,
                    output_folder=temp_output_folder,  # 传递输出文件夹路径用于图片提取
                    original_file_path=file_path  # 传递原始文件路径用于图片命名
                )

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")

                if not result['success']:
                    return ConversionResult(success=False, message=f"DOCX到MD转换失败: {result['error']}")
                
                # 步骤5：写入文件
                if progress_callback:
                    progress_callback("正在写入...")
                
                # 步骤5：构建文件名并写入主要MD
                main_filename = f"{basename}.md"
                temp_main_output = os.path.join(temp_output_folder, main_filename)
                with open(temp_main_output, 'w', encoding='utf-8') as f:
                    f.write(result['main_content'])
                
                logger.info(f"主要部分文件已写入: {temp_main_output}")
                
                # 步骤6：如果有附件内容，写入附件MD（使用特殊命名）
                if result['attachment_content']:
                    # 提取原始文件名（不含扩展名）
                    original_basename = os.path.splitext(os.path.basename(file_path))[0]
                    
                    # 从basename中提取时间戳和description部分
                    # basename格式: "document_20231107_204530_fromDocx"
                    # 需要构建: "document_附件部分_20231107_204530_fromDocx.md"
                    parts = basename.split('_')
                    if len(parts) >= 3:
                        # 找到时间戳的位置（格式: YYYYMMDD）
                        timestamp_idx = None
                        for i, part in enumerate(parts):
                            if len(part) == 8 and part.isdigit():
                                timestamp_idx = i
                                break
                        
                        if timestamp_idx is not None:
                            # 重组文件名: 原名_附件部分_时间戳_时间_description
                            attachment_filename = f"{original_basename}_附件部分_{'_'.join(parts[timestamp_idx:])}.md"
                        else:
                            # 降级方案：如果找不到时间戳，使用basename
                            attachment_filename = f"{basename}_附件部分.md"
                    else:
                        # 降级方案
                        attachment_filename = f"{basename}_附件部分.md"
                    
                    temp_attachment_output = os.path.join(temp_output_folder, attachment_filename)
                    with open(temp_attachment_output, 'w', encoding='utf-8') as f:
                        f.write(result['attachment_content'])
                    
                    logger.info(f"附件文件已写入: {temp_attachment_output}")
                
                # 步骤7：移动整个文件夹到输出目录
                final_folder = os.path.join(output_dir, basename)
                
                # 如果目标文件夹已存在，先删除
                if os.path.exists(final_folder):
                    shutil.rmtree(final_folder)
                    logger.debug(f"已删除现有文件夹: {final_folder}")
                
                shutil.move(temp_output_folder, final_folder)
                logger.info(f"已移动文件夹到: {final_folder}")
                
                # 步骤8：如果保留中间文件，移动temp_dir中的其他规范文件（如中间DOCX）
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
                
                # 准备返回路径（主要部分MD文件的完整路径）
                main_output_path = os.path.join(final_folder, main_filename)
            
            return ConversionResult(
                success=True, 
                output_path=main_output_path, 
                message="转换为Markdown成功。"
            )
                
        except Exception as e:
            logger.error(f"执行 DocxToMdStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"发生未知错误: {e}", error=e)
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False


@register_action("validate")
class DocxValidationStrategy(BaseStrategy):
    """
    对文档文件执行错别字校对的策略。
    
    支持的输入格式：
    - DOCX：直接校对
    - DOC/WPS：自动转换为DOCX后校对
    - RTF：自动转换为DOCX后校对
    - ODT：自动转换为DOCX后校对
    
    校对功能：
    - 敏感词检查
    - 错别字检查
    - 标点符号规范检查
    - 在文档中用批注标记问题位置
    
    输出：
    - 生成校对后的新DOCX文件
    - 文件名标记为 "checked"
    - 保留原文件不变
    - 如果配置保留中间文件，也会保存转换步骤的DOCX（如有）
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行文档文件的错别字校对（支持DOC/WPS/RTF/ODT自动转换）。
        
        Args:
            file_path: 输入的文档文件路径（DOCX/DOC/WPS/RTF/ODT）
            options: 校对选项字典，包含：
                - spell_check_options: (必需) 校对选项位标志
                  * 0: 不进行校对（直接返回）
                  * >0: 启用相应的校对规则
                - cancel_event: (可选) 用于取消操作的事件对象
                - actual_format: (可选) 文件的真实格式
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含校对结果的对象
            - success: 校对是否成功
            - output_path: 校对后的DOCX文件路径
            - message: 成功或失败的描述信息
            - error: 失败时的错误对象
            
        Note:
            - 如果输入是DOC/WPS/RTF/ODT，会先自动转换为DOCX
            - 转换过程在临时目录中进行
            - 转换和校对两个步骤的文件都可能被保留（根据配置）
        """
        if options is None:
            options = {}
            
        if progress_callback:
            progress_callback("准备校对...")

        try:
            # 检查是否启用了任何校对规则
            spell_check_options = options.get("spell_check_options", 0)
            if not spell_check_options > 0:
                return ConversionResult(success=True, output_path=file_path, message="未选择任何校对规则，操作跳过。")

            cancel_event = options.get("cancel_event")
            actual_format = options.get("actual_format", 'docx')
            
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 自动转换 DOC/WPS/RTF/ODT → DOCX（如需要）
                if progress_callback:
                    progress_callback("检测文件格式...")
                
                processed_docx = _preprocess_document_file(
                    file_path, temp_dir, cancel_event, actual_format
                )
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成校对输出文件名
                output_filename = os.path.basename(
                    generate_output_path(
                        file_path,
                        section="",
                        add_timestamp=True,
                        description="checked",
                        file_type="docx"
                    )
                )
                
                # 步骤3：在临时目录进行校对
                temp_output = os.path.join(temp_dir, output_filename)
                
                if progress_callback:
                    progress_callback("校对中...")
                
                result_path = process_docx(
                    processed_docx,  # 使用预处理后的DOCX
                    output_path=temp_output,  # 输出到临时目录
                    spell_check_options=spell_check_options,
                    progress_callback=progress_callback,
                    cancel_event=cancel_event
                )

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")

                if not result_path or not os.path.exists(result_path):
                    return ConversionResult(success=False, message="校对失败，请查看日志。")
                
                # 准备最终输出路径
                final_output = os.path.join(output_dir, output_filename)
                
                # 步骤4：根据配置移动文件（保留中间文件逻辑）
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
                    shutil.move(result_path, final_output)
                    logger.debug(f"已移动最终文件: {os.path.basename(result_path)}")
                
                logger.info(f"校对完成，文件已保存: {final_output}")
                
                return ConversionResult(success=True, output_path=final_output, message="校对完成。")

        except Exception as e:
            logger.error(f"执行 DocxValidationStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"发生未知错误: {e}", error=e)
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False


@register_conversion(CATEGORY_DOCUMENT, 'txt')
class DocxToTxtStrategy(BaseStrategy):
    """
    将DOCX文件转换为TXT纯文本文件的策略。
    
    转换流程：
    1. 将DOCX转换为Markdown（带YAML头部）
    2. 移除YAML元数据头部
    3. 保存为TXT格式的纯文本
    
    功能特性：
    - 自动分离主要部分和附件部分
    - 移除格式标记，提取纯文本内容
    - 支持取消操作
    
    输出文件：
    - 主要部分TXT文件
    - 附件部分TXT文件（如果存在）
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行DOCX到TXT的转换。
        
        Args:
            file_path: 输入的DOCX文件路径
            options: 转换选项字典，包含：
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            - success: 转换是否成功
            - output_path: 主要部分的TXT文件路径
            - message: 成功或失败的描述信息
            - error: 失败时的错误对象
        """
        try:
            if progress_callback:
                progress_callback("正在转换为TXT...")
            
            options = options or {}
            cancel_event = options.get("cancel_event")
            actual_format = options.get("actual_format")  # 从options中提取actual_format
            
            # 根据actual_format生成description
            description = f"from{actual_format.capitalize()}" if actual_format else "fromDocx"
            
            # 先转换为Markdown格式（中间格式）
            result = convert_docx_to_md(
                file_path,
                config=None,
                progress_callback=progress_callback,
                cancel_event=cancel_event
            )
            
            if cancel_event and cancel_event.is_set():
                return ConversionResult(success=False, message="操作已取消")
            
            if not result['success']:
                return ConversionResult(success=False, message=f"转换失败: {result['error']}")
            
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 生成主要部分的TXT输出路径
            main_txt_output = generate_output_path(
                file_path,
                output_dir=output_dir,
                section="主要部分",
                add_timestamp=True,
                description=description,
                file_type="txt"
            )
            
            # 提取主要部分纯文本内容（移除YAML头部）
            main_plain_text = self._extract_plain_text(result['main_content'])
            
            # 写入主要部分TXT文件
            if progress_callback:
                progress_callback("正在写入...")
            
            with open(main_txt_output, 'w', encoding='utf-8') as f:
                f.write(main_plain_text)
            
            logger.info(f"主要部分TXT文件已写入: {main_txt_output}")
            
            # 如果有附件内容，写入附件TXT文件
            if result['attachment_content']:
                attachment_txt_output = generate_output_path(
                    file_path,
                    output_dir=output_dir,
                    section="附件部分",
                    add_timestamp=True,
                    description=description,
                    file_type="txt"
                )
                
                # 提取附件纯文本内容（移除YAML头部）
                attachment_plain_text = self._extract_plain_text(result['attachment_content'])
                
                if progress_callback:
                    progress_callback("正在写入...")
                
                with open(attachment_txt_output, 'w', encoding='utf-8') as f:
                    f.write(attachment_plain_text)
                
                logger.info(f"附件TXT文件已写入: {attachment_txt_output}")
            
            return ConversionResult(
                success=True, 
                output_path=main_txt_output, 
                message="转换为TXT成功。"
            )
            
        except Exception as e:
            logger.error(f"DOCX转TXT失败: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {e}", error=e)
    
    def _extract_plain_text(self, markdown_content: str) -> str:
        """
        从Markdown内容中提取纯文本。
        
        当前配置：保留YAML元数据头部。
        
        Args:
            markdown_content: 包含YAML头部的Markdown内容
            
        Returns:
            str: 完整的文本内容（包含YAML头部）
        """
        # 直接返回完整内容，保留YAML头部
        return markdown_content.strip()
        
        # 以下代码已注释，未来如需移除YAML头部可取消注释
        # # 找到第二个 --- 之后的内容
        # parts = markdown_content.split('---', 2)
        # if len(parts) >= 3:
        #     # 有YAML头部，取第三部分（正文）
        #     return parts[2].strip()
        # else:
        #     # 没有YAML头部，直接返回全部内容
        #     return markdown_content.strip()


@register_conversion(CATEGORY_DOCUMENT, 'pdf')
class DocumentToPdfStrategy(BaseStrategy):
    """
    将文档文件转换为PDF文件的策略。
    
    功能特性：
    - 使用本地Office软件（WPS或Microsoft Office）进行转换
    - 支持DOCX、DOC、WPS、RTF、ODT等文档格式
    - 转换质量高，能保持文档格式和样式
    - 生成不可编辑的PDF文档，适合最终版本归档
    - 使用 office_to_pdf 配置的软件优先级
    
    支持的输入格式：
    - DOCX (Word 2007+)
    - DOC (Word 97-2003)
    - WPS (WPS文字格式)
    - RTF (富文本格式)
    - ODT (OpenDocument文本)
    
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
        执行文档到PDF的转换。
        
        Args:
            file_path: 输入的文档文件路径（支持 DOCX/DOC/WPS/RTF/ODT）
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
            actual_format = options.get('actual_format', 'docx')
            
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
                
                # 步骤2：生成输出文件名（根据实际格式动态生成description）
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
                from gongwen_converter.converter.formats.office import docx_to_pdf
                
                result_path = docx_to_pdf(
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
            logger.error(f"文档转PDF失败 - 未找到Office软件: {e}")
            return ConversionResult(
                success=False,
                message="未找到Office软件（WPS、Microsoft Office或LibreOffice），无法转换为PDF。",
                error=str(e)
            )
        except Exception as e:
            logger.error(f"文档转PDF失败: {e}", exc_info=True)
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
            intermediate_settings = config_manager.get_intermediate_files_settings()
            return intermediate_settings.get("save_to_output", False)
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False


@register_conversion(CATEGORY_DOCUMENT, 'ofd')
class DocumentToOfdStrategy(BaseStrategy):
    """
    将DOCX文件转换为OFD文件的策略（占位实现）。
    
    当前状态：
    - 功能尚未实现
    - 调用时返回错误信息
    
    规划用途：
    - OFD是中国电子文件标准格式
    - 用于符合国产化文档标准的场景
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        OFD转换功能占位方法。
        
        Args:
            file_path: 输入的DOCX文件路径（暂未使用）
            options: 转换选项（暂未使用）
            progress_callback: 进度回调（暂未使用）
            
        Returns:
            ConversionResult: 返回失败结果，提示功能未实现
        """
        return ConversionResult(success=False, message="OFD转换功能暂未实现")


# ==================== 智能转换链：策略工厂 ====================

def _create_document_conversion_strategy(source_fmt: str, target_fmt: str):
    """
    策略工厂：动态创建文档格式转换策略
    
    使用智能转换链自动处理单步或多步转换
    
    参数:
        source_fmt: 源格式（如 'doc', 'odt'）
        target_fmt: 目标格式（如 'docx', 'odt'）
    
    返回:
        动态生成的策略类
    """
    
    @register_conversion(source_fmt, target_fmt)
    class DynamicDocumentConversionStrategy(BaseStrategy):
        """动态生成的文档转换策略"""
        
        def execute(
            self,
            file_path: str,
            options: Optional[Dict[str, Any]] = None,
            progress_callback: Optional[Callable[[str], None]] = None
        ) -> ConversionResult:
            """
            执行文档格式转换
            
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
                
                logger.debug(f"文档转换策略: {source_fmt}→{target_fmt}, 真实格式: {actual_format}")
                
                # 使用智能转换链
                from gongwen_converter.converter.smart_converter import SmartConverter, OfficeSoftwareNotFoundError
                
                converter = SmartConverter()
                output_dir = os.path.dirname(file_path)
                
                # 使用临时目录管理中间文件
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 调用SmartConverter，输出到临时目录
                    result_path = converter.convert(
                        input_path=file_path,
                        target_format=target_fmt,
                        category='document',
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
                    
                    # 准备最终输出路径
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
                intermediate_settings = config_manager.get_intermediate_files_settings()
                return intermediate_settings.get("save_to_output", False)
            except Exception as e:
                logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
                return False
    
    return DynamicDocumentConversionStrategy


# 批量注册文档格式转换策略
DOCUMENT_FORMATS = ['docx', 'doc', 'odt', 'rtf', 'wps']
for source in DOCUMENT_FORMATS:
    for target in DOCUMENT_FORMATS:
        if source != target:
            # 注意：这里会覆盖上面手写的策略类
            _create_document_conversion_strategy(source, target)

logger.info("文档格式转换策略已通过智能转换链批量注册")
