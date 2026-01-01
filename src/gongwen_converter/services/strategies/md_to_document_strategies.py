"""
处理 Markdown 文件转换的策略集合。
"""
import os
import shutil
import logging
import tempfile
from typing import Dict, Any, Callable, Optional, Tuple

from gongwen_converter.services.result import ConversionResult
from gongwen_converter.services.strategies.base_strategy import BaseStrategy
from . import register_conversion, CATEGORY_MARKDOWN
from gongwen_converter.utils.path_utils import generate_output_path
from gongwen_converter.config.config_manager import config_manager

# 导入核心转换和处理函数
from gongwen_converter.converter.md2docx.core import convert as convert_md_to_docx
from gongwen_converter.docx_spell.core import process_docx
from gongwen_converter.converter.formats.office import convert_docx_to_doc, convert_docx_to_rtf, docx_to_odt, OfficeSoftwareNotFoundError, check_office_availability

logger = logging.getLogger(__name__)


class BaseMdToDocumentStrategy(BaseStrategy):
    """
    MD转Word系列策略的基类。
    
    使用模板方法模式封装通用的转换流程：
    1. MD → DOCX（使用模板）
    2. 可选：错别字校对
    3. 可选：DOCX → 其他Word格式（DOC/WPS）
    4. 根据配置决定是否保留中间文件
    
    设计模式：
    - 模板方法模式：定义转换骨架，子类实现具体格式转换
    - 临时目录管理：确保转换过程的原子性和清洁性
    
    子类需要实现：
    - _finalize_conversion: 完成从DOCX到目标格式的转换
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行Markdown到Word文档的转换。
        
        Args:
            file_path: 输入的Markdown文件路径
            options: 转换选项字典，包含：
                - template_name: (必需) 使用的模板名称
                - spell_check_options: (可选) 校对选项位标志，默认为0（不校对）
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            - success: 转换是否成功
            - output_path: 最终输出文件路径
            - message: 成功或失败的描述信息
            - error: 失败时的错误对象
            
        Raises:
            OfficeSoftwareNotFoundError: 当转换DOC/WPS格式时，未找到所需软件
        """
        if options is None:
            options = {}
            
        cancel_event = options.get("cancel_event")

        def update_progress(message: str):
            if progress_callback:
                progress_callback(message)

        try:
            template_name = options.get("template_name")
            if not template_name:
                return ConversionResult(success=False, message="错误：未提供模板名称。")

            spell_check_options = options.get("spell_check_options", 0)
            
            from gongwen_converter.utils.workspace_manager import prepare_input_file, get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤0: 创建输入副本 input.md
                temp_md = prepare_input_file(file_path, temp_dir, 'md')
                logger.debug(f"已创建输入副本: {os.path.basename(temp_md)}")
                
                # 步骤1: MD转为初始DOCX (在临时目录中，使用副本)
                update_progress("正在转换MD到DOCX...")
                try:
                    intermediate_docx_path = self._convert_md_to_initial_docx(
                        temp_md, template_name, progress_callback, cancel_event, temp_dir, file_path, options
                    )
                except ValueError as ve:
                    # 捕获已知的ValueError（如模板样式缺失），直接返回具体错误信息给界面
                    return ConversionResult(success=False, message=str(ve), error=ve)
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                if not intermediate_docx_path:
                    return ConversionResult(success=False, message="MD到DOCX的初始转换失败。")

                # 步骤2: 运行错别字检查 (在临时目录中)
                final_docx_path, checked = self._run_spell_check(
                    intermediate_docx_path, file_path, spell_check_options, progress_callback, cancel_event, temp_dir
                )
                if cancel_event and cancel_event.is_set(): 
                    return ConversionResult(success=False, message="操作已取消")
                if not final_docx_path:
                    return ConversionResult(success=False, message="错别字检查过程失败。")

                # 步骤3: 子类实现的最终转换 (在临时目录中)
                update_progress("正在完成最终格式转换...")
                final_temp_path = self._finalize_conversion(final_docx_path, file_path, checked, cancel_event, temp_dir)
                if cancel_event and cancel_event.is_set(): 
                    return ConversionResult(success=False, message="操作已取消")
                if not (final_temp_path and os.path.exists(final_temp_path)):
                    return ConversionResult(success=False, message="最终格式转换失败。")

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

            return ConversionResult(success=True, output_path=final_output_path, message="转换成功。")

        except OfficeSoftwareNotFoundError as e:
            logger.error(f"Office/WPS软件未找到: {e}")
            # DOC/WPS格式需要精确提示需要安装Office或WPS任一软件
            return ConversionResult(success=False, message="DOC/WPS格式转换需要安装WPS或Microsoft Office任一软件", error=e)
        except Exception as e:
            logger.error(f"执行MD to Word策略时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"发生未知错误: {e}", error=e)

    def _convert_md_to_initial_docx(self, md_path: str, template_name: str, progress_callback, cancel_event, temp_dir: str, original_md_path: str, options: Optional[Dict[str, Any]] = None) -> Optional[str]:
        """
        将Markdown文件转换为初始的DOCX文件（第一步转换）。
        
        Args:
            md_path: Markdown文件路径（input.md副本）
            template_name: 使用的模板名称
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            original_md_path: 原始MD文件路径（用于生成文件名）
            options: 转换选项字典，包含序号配置等
            
        Returns:
            Optional[str]: 成功时返回生成的DOCX文件路径，失败时返回None
        """
        try:
            # 生成规范的临时输出路径（基于原始MD文件名，不是input.md）
            temp_output_filename = os.path.basename(
                generate_output_path(
                    original_md_path, temp_dir, "", True, "fromMd", "docx"
                )
            )
            temp_output = os.path.join(temp_dir, temp_output_filename)
            
            # 调用底层转换函数（传递原始源文件路径用于嵌入功能的路径解析）
            output_path = convert_md_to_docx(
                md_path,
                temp_output,
                template_name=template_name,
                spell_check_option=0,
                progress_callback=progress_callback,
                cancel_event=cancel_event,
                original_source_path=original_md_path,
                options=options
            )
            return output_path
        except ValueError:
            # 如果是ValueError（如模板样式错误），直接向上抛出，不要在这里吞掉
            raise
        except Exception as e:
            logger.error(f"_convert_md_to_initial_docx失败: {e}", exc_info=True)
            return None

    def _run_spell_check(self, docx_path: str, original_md_path: str, spell_check_options: int, progress_callback, cancel_event, temp_dir: str) -> Tuple[Optional[str], bool]:
        """
        运行错别字检查（第二步，可选）。
        
        Args:
            docx_path: 输入的DOCX文件路径
            original_md_path: 原始MD文件路径（用于生成输出路径）
            spell_check_options: 校对选项位标志（0表示不校对）
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        Returns:
            Tuple[Optional[str], bool]: (处理后的docx路径, 是否尝试了检查)
            - 如果未启用校对：返回原路径和False
            - 如果校对成功：返回新路径和True
            - 如果校对失败：返回原路径和True（继续后续流程）
        """
        has_spell_check = spell_check_options > 0
        if not has_spell_check:
            return docx_path, False
        try:
            if progress_callback: progress_callback("错别字检查中...")
            output_path = generate_output_path(original_md_path, output_dir=temp_dir, section="", add_timestamp=True, description="checked", file_type="docx")
            result_path = process_docx(
                docx_path,
                output_path=output_path,
                spell_check_options=spell_check_options,
                progress_callback=progress_callback,
                cancel_event=cancel_event
            )
            
            if result_path and os.path.exists(result_path):
                return result_path, True
            else:
                logger.warning("错别字检查失败，将使用未检查的DOCX文件。")
                return docx_path, True # 即使失败，也认为“尝试过”检查
        except Exception as e:
            logger.error(f"_run_spell_check失败: {e}", exc_info=True)
            return docx_path, True # 返回原始路径继续流程

    def _finalize_conversion(self, docx_path: str, original_md_path: str, was_checked: bool, cancel_event, temp_dir: str) -> Optional[str]:
        """
        完成最终格式转换（模板方法，第三步）。
        
        子类需要实现此方法来完成从DOCX到目标格式的转换。
        
        Args:
            docx_path: 临时目录中的DOCX文件路径
            original_md_path: 原始Markdown文件路径
            was_checked: 是否进行了错别字检查
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        Returns:
            Optional[str]: 最终文件在临时目录中的完整路径，失败时返回None
            
        Note:
            - 实现时应在temp_dir中生成文件
            - 返回temp_dir中的文件完整路径（不是输出目录的路径）
            - was_checked参数用于决定输出文件的description标记
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


@register_conversion(CATEGORY_MARKDOWN, 'docx')
class MdToDocxStrategy(BaseMdToDocumentStrategy):
    """
    将Markdown转换为DOCX文件的策略。
    
    这是最直接的转换策略，因为基类已经完成了MD→DOCX的转换，
    只需处理文件命名即可。
    
    输出文件命名：
    - 未校对：标记为 "fromMd"
    - 已校对：标记为 "checked"
    """
    def _generate_final_path(self, file_path: str, was_checked: bool) -> str:
        """生成DOCX输出路径"""
        description = "checked" if was_checked else "fromMd"
        return generate_output_path(file_path, section="", add_timestamp=True, description=description, file_type="docx")
    
    def _finalize_conversion(self, docx_path: str, original_md_path: str, was_checked: bool, cancel_event, temp_dir: str) -> Optional[str]:
        """
        重命名DOCX文件为最终输出文件名。
        
        Args:
            docx_path: 临时目录中的DOCX文件路径
            original_md_path: 原始MD文件路径
            was_checked: 是否进行了错别字检查
            cancel_event: 取消事件对象（未使用）
            temp_dir: 临时目录路径
            
        Returns:
            str: 临时目录中最终文件的完整路径
        """
        description = "checked" if was_checked else "fromMd"
        final_filename = os.path.basename(
            generate_output_path(original_md_path, section="", add_timestamp=True, description=description, file_type="docx")
        )
        # 将临时文件重命名为最终文件名，仍在临时目录中
        final_temp_path = os.path.join(temp_dir, final_filename)
        os.rename(docx_path, final_temp_path)
        return final_temp_path


@register_conversion(CATEGORY_MARKDOWN, 'doc')
class MdToDocStrategy(BaseMdToDocumentStrategy):
    """
    将Markdown转换为DOC文件的策略。
    
    转换流程：MD → DOCX → DOC
    DOC是Word 97-2003格式，需要安装WPS或Microsoft Office软件。
    
    输出文件命名：
    - 未校对：标记为 "fromMd"
    - 已校对：标记为 "checked"
    """
    def _generate_final_path(self, file_path: str, was_checked: bool) -> str:
        """生成DOC输出路径"""
        description = "checked" if was_checked else "fromMd"
        return generate_output_path(file_path, section="", add_timestamp=True, description=description, file_type="doc")
    
    def _finalize_conversion(self, docx_path: str, original_md_path: str, was_checked: bool, cancel_event, temp_dir: str) -> Optional[str]:
        """
        将DOCX文件转换为DOC格式。
        
        Args:
            docx_path: 临时目录中的DOCX文件路径
            original_md_path: 原始MD文件路径
            was_checked: 是否进行了错别字检查
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        Returns:
            Optional[str]: 成功时返回临时目录中DOC文件的完整路径，失败时返回None
            
        Raises:
            OfficeSoftwareNotFoundError: 未找到Office或WPS软件时抛出
        """
        try:
            description = "checked" if was_checked else "fromMd"
            # 生成最终文件名
            final_filename = os.path.basename(
                generate_output_path(original_md_path, section="", add_timestamp=True, description=description, file_type="doc")
            )
            # 在临时目录中生成 .doc 文件
            temp_output_path = os.path.join(temp_dir, final_filename)
            
            result_path = convert_docx_to_doc(docx_path, temp_output_path, cancel_event=cancel_event)
            
            # 返回临时目录中的路径
            return temp_output_path if result_path and os.path.exists(temp_output_path) else None
        except Exception as e:
            logger.error(f"MdToDocStrategy._finalize_conversion失败: {e}", exc_info=True)
            return None


@register_conversion(CATEGORY_MARKDOWN, 'odt')
class MdToOdtStrategy(BaseMdToDocumentStrategy):
    """
    将Markdown转换为ODT文件的策略。
    
    转换流程：MD → DOCX → ODT
    ODT是OpenDocument文本格式，需要安装Microsoft Office或LibreOffice软件。
    
    输出文件命名：
    - 未校对：标记为 "fromMd"
    - 已校对：标记为 "checked"
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行Markdown到ODT的转换，开始前进行软件可用性预检查。
        """
        # 预检查：ODT格式需要Microsoft Word或LibreOffice，WPS不支持
        available, error_msg = check_office_availability('odt')
        if not available:
            logger.error(f"ODT转换预检查失败: {error_msg}")
            raise OfficeSoftwareNotFoundError(error_msg)
        
        # 调用父类的execute方法
        return super().execute(file_path, options, progress_callback)
    
    def _generate_final_path(self, file_path: str, was_checked: bool) -> str:
        """生成ODT输出路径"""
        description = "checked" if was_checked else "fromMd"
        return generate_output_path(file_path, section="", add_timestamp=True, description=description, file_type="odt")
    
    def _finalize_conversion(self, docx_path: str, original_md_path: str, was_checked: bool, cancel_event, temp_dir: str) -> Optional[str]:
        """
        将DOCX文件转换为ODT格式。
        
        Args:
            docx_path: 临时目录中的DOCX文件路径
            original_md_path: 原始MD文件路径
            was_checked: 是否进行了错别字检查
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        Returns:
            Optional[str]: 成功时返回临时目录中ODT文件的完整路径，失败时返回None
            
        Raises:
            OfficeSoftwareNotFoundError: 未找到Office或LibreOffice软件时抛出
        """
        # 预检查：ODT格式需要Microsoft Word或LibreOffice，WPS不支持
        available, error_msg = check_office_availability('odt')
        if not available:
            logger.error(f"ODT转换预检查失败: {error_msg}")
            raise OfficeSoftwareNotFoundError(error_msg)
        
        try:
            description = "checked" if was_checked else "fromMd"
            # 生成最终文件名
            final_filename = os.path.basename(
                generate_output_path(original_md_path, section="", add_timestamp=True, description=description, file_type="odt")
            )
            # 在临时目录中生成 .odt 文件
            temp_output_path = os.path.join(temp_dir, final_filename)
            
            result_path = docx_to_odt(docx_path, temp_output_path, cancel_event=cancel_event)
            
            # 返回临时目录中的路径
            return temp_output_path if result_path and os.path.exists(temp_output_path) else None
        except OfficeSoftwareNotFoundError:
            raise  # 预检查异常向上抛出
        except Exception as e:
            logger.error(f"MdToOdtStrategy._finalize_conversion失败: {e}", exc_info=True)
            return None


@register_conversion(CATEGORY_MARKDOWN, 'rtf')
class MdToRtfStrategy(BaseMdToDocumentStrategy):
    """
    将Markdown转换为RTF文件的策略。
    
    转换流程：MD → DOCX → RTF
    RTF是富文本格式，需要安装WPS或Microsoft Office软件来完成转换。
    
    输出文件命名：
    - 未校对：标记为 "fromMd"
    - 已校对：标记为 "checked"
    """
    def _generate_final_path(self, file_path: str, was_checked: bool) -> str:
        """生成RTF输出路径"""
        description = "checked" if was_checked else "fromMd"
        return generate_output_path(file_path, section="", add_timestamp=True, description=description, file_type="rtf")
    
    def _finalize_conversion(self, docx_path: str, original_md_path: str, was_checked: bool, cancel_event, temp_dir: str) -> Optional[str]:
        """
        将DOCX文件转换为RTF格式。
        
        Args:
            docx_path: 临时目录中的DOCX文件路径
            original_md_path: 原始MD文件路径
            was_checked: 是否进行了错别字检查
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        Returns:
            Optional[str]: 成功时返回临时目录中RTF文件的完整路径，失败时返回None
            
        Raises:
            OfficeSoftwareNotFoundError: 未找到Office或WPS软件时抛出
        """
        try:
            description = "checked" if was_checked else "fromMd"
            # 生成目标文件路径
            final_filename = os.path.basename(
                generate_output_path(original_md_path, section="", add_timestamp=True, description=description, file_type="rtf")
            )
            # 在临时目录中生成 .rtf 文件
            temp_output_path = os.path.join(temp_dir, final_filename)
            
            result_path = convert_docx_to_rtf(docx_path, temp_output_path, cancel_event=cancel_event)
            
            # 返回临时目录中的路径
            return temp_output_path if result_path and os.path.exists(temp_output_path) else None
        except Exception as e:
            logger.error(f"MdToRtfStrategy._finalize_conversion失败: {e}", exc_info=True)
            return None
