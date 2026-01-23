"""
Markdown 转文档策略模块

将 Markdown 文件转换为各种文档格式（DOCX、DOC、ODT、RTF）。

转换流程：
1. MD → DOCX（使用模板）
2. 可选：错别字校对
3. 可选：DOCX → 其他Word格式（DOC/ODT/RTF）

依赖：
- converter.md2docx: MD → DOCX 核心转换
- converter.formats.document: DOCX → 其他文档格式
- docx_spell: 错别字校对
"""
import os
import shutil
import logging
import tempfile
from typing import Dict, Any, Callable, Optional, Tuple

from docwen.services.result import ConversionResult
from docwen.services.strategies.base_strategy import BaseStrategy
from docwen.services.strategies import register_conversion, CATEGORY_MARKDOWN
from docwen.utils.path_utils import generate_output_path
from docwen.config.config_manager import config_manager
from docwen.i18n import t

# 导入核心转换和处理函数
from docwen.converter.md2docx.core import convert as convert_md_to_docx
from docwen.docx_spell.core import process_docx
from docwen.converter.formats.document import (
    docx_to_doc,
    docx_to_rtf,
    docx_to_odt,
)
from docwen.converter.formats.common import (
    OfficeSoftwareNotFoundError,
    check_office_availability,
)

logger = logging.getLogger(__name__)


class BaseMdToDocumentStrategy(BaseStrategy):
    """
    MD转Word系列策略的基类
    
    使用模板方法模式封装通用的转换流程：
    1. MD → DOCX（使用模板）
    2. 可选：错别字校对
    3. 可选：DOCX → 其他Word格式（DOC/WPS）
    4. 根据配置决定是否保留中间文件
    
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
        执行Markdown到Word文档的转换
        
        参数:
            file_path: 输入的Markdown文件路径
            options: 转换选项字典，包含：
                - template_name: (必需) 使用的模板名称
                - proofread_options: (可选) 校对选项字典，默认为None（使用配置）
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        返回:
            ConversionResult: 包含转换结果的对象
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
                return ConversionResult(success=False, message=t('conversion.progress.error_no_template'))

            proofread_options = options.get("proofread_options")
            
            from docwen.utils.workspace_manager import prepare_input_file, get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤0: 创建输入副本 input.md
                temp_md = prepare_input_file(file_path, temp_dir, 'md')
                logger.debug(f"已创建输入副本: {os.path.basename(temp_md)}")
                
                # 步骤1: MD转为初始DOCX (在临时目录中，使用副本)
                try:
                    intermediate_docx_path = self._convert_md_to_initial_docx(
                        temp_md, template_name, progress_callback, cancel_event, temp_dir, file_path, options
                    )
                except ValueError as ve:
                    # 捕获已知的ValueError（如模板样式缺失），直接返回具体错误信息给界面
                    return ConversionResult(success=False, message=str(ve), error=ve)
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                if not intermediate_docx_path:
                    return ConversionResult(success=False, message=t('conversion.messages.md_to_docx_failed'))

                # 步骤2: 运行错别字检查 (在临时目录中)
                final_docx_path, checked = self._run_spell_check(
                    intermediate_docx_path, file_path, proofread_options, progress_callback, cancel_event, temp_dir
                )
                if cancel_event and cancel_event.is_set(): 
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                if not final_docx_path:
                    return ConversionResult(success=False, message=t('conversion.messages.spell_check_failed'))

                # 步骤3: 子类实现的最终转换 (在临时目录中)
                update_progress(t('conversion.progress.final_conversion'))
                final_temp_path = self._finalize_conversion(final_docx_path, file_path, checked, cancel_event, temp_dir)
                if cancel_event and cancel_event.is_set(): 
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                if not (final_temp_path and os.path.exists(final_temp_path)):
                    return ConversionResult(success=False, message=t('conversion.messages.final_format_failed'))

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

            # 准备返回消息
            target_fmt = os.path.splitext(final_output_path)[1].lstrip('.').upper()
            return ConversionResult(
                success=True, 
                output_path=final_output_path, 
                message=t('conversion.messages.conversion_to_format_success', format=target_fmt)
            )

        except OfficeSoftwareNotFoundError as e:
            logger.error(f"Office/WPS软件未找到: {e}")
            # DOC/WPS格式需要精确提示需要安装Office或WPS任一软件
            return ConversionResult(success=False, message=t('conversion.messages.doc_format_requires_office'), error=e)
        except Exception as e:
            logger.error(f"执行MD to Word策略时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"发生未知错误: {e}", error=e)

    def _convert_md_to_initial_docx(
        self,
        md_path: str,
        template_name: str,
        progress_callback,
        cancel_event,
        temp_dir: str,
        original_md_path: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Optional[str]:
        """
        将Markdown文件转换为初始的DOCX文件（第一步转换）
        
        参数:
            md_path: Markdown文件路径（input.md副本）
            template_name: 使用的模板名称
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            original_md_path: 原始MD文件路径（用于生成文件名）
            options: 转换选项字典，包含序号配置等
            
        返回:
            成功时返回生成的DOCX文件路径，失败时返回None
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

    def _run_spell_check(
        self,
        docx_path: str,
        original_md_path: str,
        proofread_options: Optional[Dict[str, bool]],
        progress_callback,
        cancel_event,
        temp_dir: str
    ) -> Tuple[Optional[str], bool]:
        """
        运行错别字检查（第二步，可选）
        
        参数:
            docx_path: 输入的DOCX文件路径
            original_md_path: 原始MD文件路径（用于生成输出路径）
            proofread_options: 校对选项字典，None或空字典表示不校对
            progress_callback: 进度回调函数
            cancel_event: 取消事件对象
            temp_dir: 临时目录路径
            
        返回:
            (处理后的docx路径, 是否尝试了检查)
        """
        # 检查是否需要校对
        has_spell_check = proofread_options and any(proofread_options.values())
        if not has_spell_check:
            return docx_path, False
        try:
            if progress_callback: 
                progress_callback(t('conversion.progress.spell_checking'))
            output_path = generate_output_path(
                original_md_path, output_dir=temp_dir, section="", 
                add_timestamp=True, description="checked", file_type="docx"
            )
            result_path = process_docx(
                docx_path,
                output_path=output_path,
                proofread_options=proofread_options,
                progress_callback=progress_callback,
                cancel_event=cancel_event
            )
            
            if result_path and os.path.exists(result_path):
                return result_path, True
            else:
                logger.warning("错别字检查失败，将使用未检查的DOCX文件。")
                return docx_path, True  # 即使失败，也认为"尝试过"检查
        except Exception as e:
            logger.error(f"_run_spell_check失败: {e}", exc_info=True)
            return docx_path, True  # 返回原始路径继续流程

    def _finalize_conversion(
        self,
        docx_path: str,
        original_md_path: str,
        was_checked: bool,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """
        完成最终格式转换（模板方法，第三步）
        
        子类需要实现此方法来完成从DOCX到目标格式的转换。
        
        参数:
            docx_path: 临时目录中的DOCX文件路径
            original_md_path: 原始Markdown文件路径
            was_checked: 是否进行了错别字检查
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


@register_conversion(CATEGORY_MARKDOWN, 'docx')
class MdToDocxStrategy(BaseMdToDocumentStrategy):
    """
    将Markdown转换为DOCX文件的策略
    
    这是最直接的转换策略，因为基类已经完成了MD→DOCX的转换，
    只需处理文件命名即可。
    """
    
    def _generate_final_path(self, file_path: str, was_checked: bool) -> str:
        """生成DOCX输出路径"""
        description = "checked" if was_checked else "fromMd"
        return generate_output_path(
            file_path, section="", add_timestamp=True, 
            description=description, file_type="docx"
        )
    
    def _finalize_conversion(
        self,
        docx_path: str,
        original_md_path: str,
        was_checked: bool,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """重命名DOCX文件为最终输出文件名"""
        description = "checked" if was_checked else "fromMd"
        final_filename = os.path.basename(
            generate_output_path(
                original_md_path, section="", add_timestamp=True, 
                description=description, file_type="docx"
            )
        )
        # 将临时文件重命名为最终文件名，仍在临时目录中
        final_temp_path = os.path.join(temp_dir, final_filename)
        os.rename(docx_path, final_temp_path)
        return final_temp_path


@register_conversion(CATEGORY_MARKDOWN, 'doc')
class MdToDocStrategy(BaseMdToDocumentStrategy):
    """
    将Markdown转换为DOC文件的策略
    
    转换流程：MD → DOCX → DOC
    DOC是Word 97-2003格式，需要安装WPS或Microsoft Office软件。
    """
    
    def _generate_final_path(self, file_path: str, was_checked: bool) -> str:
        """生成DOC输出路径"""
        description = "checked" if was_checked else "fromMd"
        return generate_output_path(
            file_path, section="", add_timestamp=True, 
            description=description, file_type="doc"
        )
    
    def _finalize_conversion(
        self,
        docx_path: str,
        original_md_path: str,
        was_checked: bool,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """将DOCX文件转换为DOC格式"""
        try:
            description = "checked" if was_checked else "fromMd"
            # 生成最终文件名
            final_filename = os.path.basename(
                generate_output_path(
                    original_md_path, section="", add_timestamp=True, 
                    description=description, file_type="doc"
                )
            )
            # 在临时目录中生成 .doc 文件
            temp_output_path = os.path.join(temp_dir, final_filename)
            
            result_path = docx_to_doc(
                docx_path, temp_output_path, cancel_event=cancel_event
            )
            
            # 返回临时目录中的路径
            return temp_output_path if result_path and os.path.exists(temp_output_path) else None
        except Exception as e:
            logger.error(f"MdToDocStrategy._finalize_conversion失败: {e}", exc_info=True)
            return None


@register_conversion(CATEGORY_MARKDOWN, 'odt')
class MdToOdtStrategy(BaseMdToDocumentStrategy):
    """
    将Markdown转换为ODT文件的策略
    
    转换流程：MD → DOCX → ODT
    ODT是OpenDocument文本格式，需要安装Microsoft Office或LibreOffice软件。
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """执行Markdown到ODT的转换，开始前进行软件可用性预检查"""
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
        return generate_output_path(
            file_path, section="", add_timestamp=True, 
            description=description, file_type="odt"
        )
    
    def _finalize_conversion(
        self,
        docx_path: str,
        original_md_path: str,
        was_checked: bool,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """将DOCX文件转换为ODT格式"""
        # 预检查：ODT格式需要Microsoft Word或LibreOffice，WPS不支持
        available, error_msg = check_office_availability('odt')
        if not available:
            logger.error(f"ODT转换预检查失败: {error_msg}")
            raise OfficeSoftwareNotFoundError(error_msg)
        
        try:
            description = "checked" if was_checked else "fromMd"
            # 生成最终文件名
            final_filename = os.path.basename(
                generate_output_path(
                    original_md_path, section="", add_timestamp=True, 
                    description=description, file_type="odt"
                )
            )
            # 在临时目录中生成 .odt 文件
            temp_output_path = os.path.join(temp_dir, final_filename)
            
            result_path = docx_to_odt(
                docx_path, temp_output_path, cancel_event=cancel_event
            )
            
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
    将Markdown转换为RTF文件的策略
    
    转换流程：MD → DOCX → RTF
    RTF是富文本格式，需要安装WPS或Microsoft Office软件来完成转换。
    """
    
    def _generate_final_path(self, file_path: str, was_checked: bool) -> str:
        """生成RTF输出路径"""
        description = "checked" if was_checked else "fromMd"
        return generate_output_path(
            file_path, section="", add_timestamp=True, 
            description=description, file_type="rtf"
        )
    
    def _finalize_conversion(
        self,
        docx_path: str,
        original_md_path: str,
        was_checked: bool,
        cancel_event,
        temp_dir: str
    ) -> Optional[str]:
        """将DOCX文件转换为RTF格式"""
        try:
            description = "checked" if was_checked else "fromMd"
            # 生成目标文件路径
            final_filename = os.path.basename(
                generate_output_path(
                    original_md_path, section="", add_timestamp=True, 
                    description=description, file_type="rtf"
                )
            )
            # 在临时目录中生成 .rtf 文件
            temp_output_path = os.path.join(temp_dir, final_filename)
            
            result_path = docx_to_rtf(
                docx_path, temp_output_path, cancel_event=cancel_event
            )
            
            # 返回临时目录中的路径
            return temp_output_path if result_path and os.path.exists(temp_output_path) else None
        except Exception as e:
            logger.error(f"MdToRtfStrategy._finalize_conversion失败: {e}", exc_info=True)
            return None
