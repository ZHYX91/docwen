"""
版式文件转换策略集合

处理PDF、CAJ、OFD等版式文件的转换
版式文件特点：固定布局、不可编辑、适合归档和分发
"""

import os
import logging
import tempfile
import shutil
from typing import Dict, Any, Callable, Optional

from gongwen_converter.services.result import ConversionResult
from gongwen_converter.services.strategies.base_strategy import BaseStrategy
from . import register_conversion, register_action, CATEGORY_LAYOUT

logger = logging.getLogger(__name__)


def _preprocess_layout_file(file_path: str, temp_dir: str, cancel_event=None, actual_format: str = None) -> tuple:
    """
    预处理版式文件：创建输入副本，并将非标准格式转换为PDF
    
    参数:
        file_path: 原始文件路径
        temp_dir: 临时目录路径，转换后的中间文件将输出到此目录
        cancel_event: 用于取消操作的事件对象
        actual_format: 实际文件格式（可选，如果不提供则自动检测）
        
    返回:
        tuple: (处理后的PDF路径, 中间文件路径或None)
            - 所有格式都会先创建 input.{ext} 副本
            - 如果需要转换（XPS/CAJ/OFD），返回转换后的PDF路径和中间PDF文件路径
            - 如果已是标准格式（PDF），返回副本路径和None
            
    说明:
        - 使用actual_format参数避免重复检测文件格式
        - 所有中间文件都输出到temp_dir，由调用者的上下文管理器统一清理
        - 返回的中间文件路径可用于保留中间文件功能
    """
    # 如果没有提供actual_format，则检测
    if actual_format is None:
        from gongwen_converter.utils.file_type_utils import detect_actual_file_format
        actual_format = detect_actual_file_format(file_path)
        logger.debug(f"自动检测版式文件格式: {actual_format}")
    else:
        logger.debug(f"使用传入的文件格式: {actual_format}")
    
    # 步骤1：无论什么格式，都先创建输入副本 input.{ext}
    from gongwen_converter.utils.workspace_manager import prepare_input_file
    temp_input = prepare_input_file(file_path, temp_dir, actual_format)
    logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
    
    # 步骤2：如果是标准格式（PDF），直接返回副本路径，无中间文件
    if actual_format == 'pdf':
        logger.debug(f"文件已是PDF格式，返回副本路径")
        return temp_input, None
    
    # 步骤3：需要转换的格式 - CAJ
    if actual_format == 'caj':
        logger.info(f"检测到CAJ格式，从副本转换为PDF: {os.path.basename(temp_input)}")
        
        try:
            from gongwen_converter.converter.formats.layout import caj_to_pdf
            
            # 调用转换函数，使用副本作为输入
            converted_path = caj_to_pdf(temp_input, cancel_event, output_dir=temp_dir)
            
            if converted_path:
                logger.info(f"CAJ转PDF成功: {os.path.basename(converted_path)}")
                return converted_path, converted_path  # 返回PDF路径和中间文件路径
            else:
                logger.error("CAJ转PDF失败，返回None")
                raise RuntimeError("CAJ转PDF失败")
                
        except NotImplementedError as e:
            logger.error(f"CAJ转PDF功能尚未实现: {e}")
            raise RuntimeError(f"CAJ转PDF功能尚未实现: {e}")
        except Exception as e:
            logger.error(f"CAJ转PDF失败: {e}")
            raise RuntimeError(f"CAJ转PDF失败: {e}")
    
    # 步骤4：需要转换的格式 - XPS
    elif actual_format == 'xps':
        logger.info(f"检测到XPS格式，从副本转换为PDF: {os.path.basename(temp_input)}")
        
        try:
            from gongwen_converter.converter.formats.layout import xps_to_pdf
            
            # 调用转换函数，使用副本作为输入
            converted_path = xps_to_pdf(temp_input, cancel_event, output_dir=temp_dir)
            
            if converted_path:
                logger.info(f"XPS转PDF成功: {os.path.basename(converted_path)}")
                return converted_path, converted_path  # 返回PDF路径和中间文件路径
            else:
                logger.error("XPS转PDF失败，返回None")
                raise RuntimeError("XPS转PDF失败")
                
        except Exception as e:
            logger.error(f"XPS转PDF失败: {e}")
            raise RuntimeError(f"XPS转PDF失败: {e}")
    
    # 步骤5：需要转换的格式 - OFD
    elif actual_format == 'ofd':
        logger.info(f"检测到OFD格式，从副本转换为PDF: {os.path.basename(temp_input)}")
        
        try:
            from gongwen_converter.converter.formats.layout import ofd_to_pdf
            
            # 调用转换函数，使用副本作为输入
            converted_path = ofd_to_pdf(temp_input, cancel_event, output_dir=temp_dir)
            
            if converted_path:
                logger.info(f"OFD转PDF成功: {os.path.basename(converted_path)}")
                return converted_path, converted_path  # 返回PDF路径和中间文件路径
            else:
                logger.error("OFD转PDF失败，返回None")
                raise RuntimeError("OFD转PDF失败")
                
        except NotImplementedError as e:
            logger.error(f"OFD转PDF功能尚未实现: {e}")
            raise RuntimeError(f"OFD转PDF功能尚未实现: {e}")
        except Exception as e:
            logger.error(f"OFD转PDF失败: {e}")
            raise RuntimeError(f"OFD转PDF失败: {e}")
    
    # 步骤6：其他不支持的格式
    else:
        logger.warning(f"不支持的版式文件格式: {actual_format}")
        raise RuntimeError(f"不支持的版式文件格式: {actual_format}")


@register_conversion(CATEGORY_LAYOUT, 'txt')
class LayoutToTxtStrategy(BaseStrategy):
    """
    将版式文件转换为TXT文件的策略（占位实现）
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF
    2. 核心转换：PDF → TXT（OCR提取纯文本）
    3. 清理：根据配置决定中间文件去向
    
    支持的输入格式：
    - PDF：直接处理（待实现）
    - XPS：先转为PDF再处理
    - CAJ：先转为PDF再处理（待实现）
    - OFD：先转为PDF再处理（待实现）
    
    当前状态：
    - 功能框架已建立
    - PDF转TXT核心功能待实现
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行版式文件到TXT的转换
        
        Args:
            file_path: 输入的版式文件路径
            options: 转换选项字典
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            # TODO: 实现PDF转TXT功能
            # 可能需要OCR
            
            return ConversionResult(
                success=False,
                message="版式文件转TXT功能开发中，敬请期待。"
            )
            
        except Exception as e:
            logger.error(f"执行 LayoutToTxtStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"发生错误: {e}", error=e)


@register_conversion(CATEGORY_LAYOUT, 'docx')
class LayoutToDocxStrategy(BaseStrategy):
    """
    使用外部工具将PDF转换为DOCX的策略（v2.0改进）
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 核心转换：PDF → DOCX（使用外部工具：Word > LibreOffice > pdf2docx，在临时目录）
    3. 移动到输出目录
    
    改进：
    - 使用 generate_output_path() 统一命名
    - description 统一为 "fromPdf"
    - 所有转换在临时目录完成
    - 文件命名规范：document_20251111_123456_fromPdf.docx
    
    特点：
    - 转换质量高（使用原生工具）
    - 图片自动内嵌在DOCX中
    - 不需要额外的提取选项
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - CAJ：先转为PDF再处理（待实现）
    - OFD：先转为PDF再处理（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF到DOCX的转换（使用外部工具）
        
        Args:
            file_path: 输入的PDF文件路径
            options: 转换选项字典
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 预处理：确保文件是PDF格式
                if actual_format and actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, _ = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 使用标准化路径生成（在临时目录）
                if progress_callback:
                    progress_callback("使用外部工具转换PDF...")
                
                from gongwen_converter.utils.path_utils import generate_output_path
                
                # 生成标准化DOCX路径（在临时目录）
                docx_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type="docx"
                )
                
                from gongwen_converter.converter.pdf_processing.pdf_converter_utils import convert_pdf_to_docx
                
                docx_path = convert_pdf_to_docx(
                    pdf_path,
                    docx_temp_path,
                    cancel_event
                )
                
                if not docx_path or not os.path.exists(docx_path):
                    return ConversionResult(
                        success=False,
                        message="外部工具转换失败，请安装Word、LibreOffice或pdf2docx库"
                    )
                
                logger.info(f"DOCX文件已生成: {os.path.basename(docx_path)}")
                
                # 移动文件到输出目录
                from gongwen_converter.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                final_docx_path = os.path.join(output_dir, os.path.basename(docx_path))
                shutil.move(docx_path, final_docx_path)
                logger.info(f"DOCX文件已移动到: {final_docx_path}")
                
                return ConversionResult(
                    success=True,
                    output_path=final_docx_path,
                    message="已成功转换为DOCX"
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except ImportError as e:
            error_msg = str(e)
            logger.error(f"缺少必要的库: {error_msg}")
            return ConversionResult(success=False, message=error_msg)
        except Exception as e:
            logger.error(f"执行 LayoutToDocxStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)


@register_conversion(CATEGORY_LAYOUT, 'doc')
class LayoutToDocStrategy(BaseStrategy):
    """
    将PDF转换为DOC格式的策略（v2.0改进）
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. PDF → DOCX（使用外部工具，在临时目录）
    3. DOCX → DOC（使用Office支持模块，在临时目录）
    4. 根据配置决定是否保留中间DOCX文件
    
    改进：
    - 使用 generate_output_path() 统一命名
    - description 统一为 "fromPdf"（即使经过DOCX中转）
    - 支持中间文件保留配置
    - 所有转换在临时目录完成
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - CAJ：先转为PDF再处理（待实现）
    - OFD：先转为PDF再处理（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF到DOC的转换
        
        Args:
            file_path: 输入的PDF文件路径
            options: 转换选项字典
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 预处理：确保文件是PDF格式
                if actual_format and actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, _ = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 第1步：PDF → DOCX（在临时目录，使用标准命名）
                if progress_callback:
                    progress_callback("步骤1/2: PDF转DOCX...")
                
                from gongwen_converter.utils.path_utils import generate_output_path
                
                # 生成中间DOCX的标准化路径（在临时目录）
                docx_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type="docx"
                )
                
                from gongwen_converter.converter.pdf_processing.pdf_converter_utils import convert_pdf_to_docx
                
                docx_path = convert_pdf_to_docx(
                    pdf_path,
                    docx_temp_path,
                    cancel_event
                )
                
                if not docx_path or not os.path.exists(docx_path):
                    return ConversionResult(
                        success=False,
                        message="PDF转DOCX失败，请安装Word、LibreOffice或pdf2docx库"
                    )
                
                logger.info(f"中间DOCX已生成: {os.path.basename(docx_path)}")
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 第2步：DOCX → DOC（在临时目录，使用标准命名）
                if progress_callback:
                    progress_callback("步骤2/2: DOCX转DOC...")
                
                # 生成最终DOC的标准化路径（在临时目录）
                doc_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type="doc"
                )
                
                from gongwen_converter.converter.formats.office import convert_docx_to_doc
                
                doc_path = convert_docx_to_doc(docx_path, doc_temp_path, cancel_event)
                
                if not doc_path or not os.path.exists(doc_path):
                    return ConversionResult(
                        success=False,
                        message="DOCX转DOC失败，请安装Word或LibreOffice"
                    )
                
                logger.info(f"最终DOC已生成: {os.path.basename(doc_path)}")
                
                # 移动文件到输出目录
                from gongwen_converter.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                final_doc_path = os.path.join(output_dir, os.path.basename(doc_path))
                shutil.move(doc_path, final_doc_path)
                logger.info(f"DOC文件已移动到: {final_doc_path}")
                
                # 根据配置决定是否保留中间DOCX文件
                if self._should_keep_intermediates():
                    final_docx_path = os.path.join(output_dir, os.path.basename(docx_path))
                    shutil.move(docx_path, final_docx_path)
                    logger.info(f"保留中间文件: {final_docx_path}")
                else:
                    logger.debug("清理中间DOCX文件")
                
                return ConversionResult(
                    success=True,
                    output_path=final_doc_path,
                    message="已成功转换为DOC"
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except ImportError as e:
            error_msg = str(e)
            logger.error(f"缺少必要的库: {error_msg}")
            return ConversionResult(success=False, message=error_msg)
        except Exception as e:
            logger.error(f"执行 LayoutToDocStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False


@register_conversion(CATEGORY_LAYOUT, 'odt')
class LayoutToOdtStrategy(BaseStrategy):
    """
    将PDF转换为ODT格式的策略
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. PDF → DOCX（使用外部工具，在临时目录）
    3. DOCX → ODT（使用Office支持模块，在临时目录）
    4. 根据配置决定是否保留中间DOCX文件
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - CAJ：先转为PDF再处理（待实现）
    - OFD：先转为PDF再处理（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF到ODT的转换
        
        Args:
            file_path: 输入的PDF文件路径
            options: 转换选项字典
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 预处理：确保文件是PDF格式
                if actual_format and actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, _ = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 第1步：PDF → DOCX（在临时目录，使用标准命名）
                if progress_callback:
                    progress_callback("步骤1/2: PDF转DOCX...")
                
                from gongwen_converter.utils.path_utils import generate_output_path
                
                # 生成中间DOCX的标准化路径（在临时目录）
                docx_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type="docx"
                )
                
                from gongwen_converter.converter.pdf_processing.pdf_converter_utils import convert_pdf_to_docx
                
                docx_path = convert_pdf_to_docx(
                    pdf_path,
                    docx_temp_path,
                    cancel_event
                )
                
                if not docx_path or not os.path.exists(docx_path):
                    return ConversionResult(
                        success=False,
                        message="PDF转DOCX失败，请安装Word、LibreOffice或pdf2docx库"
                    )
                
                logger.info(f"中间DOCX已生成: {os.path.basename(docx_path)}")
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 第2步：DOCX → ODT（在临时目录，使用标准命名）
                if progress_callback:
                    progress_callback("步骤2/2: DOCX转ODT...")
                
                # 生成最终ODT的标准化路径（在临时目录）
                odt_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type="odt"
                )
                
                from gongwen_converter.converter.formats.office import convert_docx_to_odt
                
                odt_path = convert_docx_to_odt(docx_path, odt_temp_path, cancel_event)
                
                if not odt_path or not os.path.exists(odt_path):
                    return ConversionResult(
                        success=False,
                        message="DOCX转ODT失败，请安装LibreOffice或Word"
                    )
                
                logger.info(f"最终ODT已生成: {os.path.basename(odt_path)}")
                
                # 移动文件到输出目录
                from gongwen_converter.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                final_odt_path = os.path.join(output_dir, os.path.basename(odt_path))
                shutil.move(odt_path, final_odt_path)
                logger.info(f"ODT文件已移动到: {final_odt_path}")
                
                # 根据配置决定是否保留中间DOCX文件
                if self._should_keep_intermediates():
                    final_docx_path = os.path.join(output_dir, os.path.basename(docx_path))
                    shutil.move(docx_path, final_docx_path)
                    logger.info(f"保留中间文件: {final_docx_path}")
                else:
                    logger.debug("清理中间DOCX文件")
                
                return ConversionResult(
                    success=True,
                    output_path=final_odt_path,
                    message="已成功转换为ODT"
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except ImportError as e:
            error_msg = str(e)
            logger.error(f"缺少必要的库: {error_msg}")
            return ConversionResult(success=False, message=error_msg)
        except Exception as e:
            logger.error(f"执行 LayoutToOdtStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)
    
        @staticmethod
        def _should_keep_intermediates() -> bool:
            """判断是否应该保留中间文件"""
            try:
                from gongwen_converter.config.config_manager import config_manager
                return config_manager.get_save_intermediate_files()
            except Exception as e:
                logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
                return False


@register_conversion(CATEGORY_LAYOUT, 'rtf')
class LayoutToRtfStrategy(BaseStrategy):
    """
    将PDF转换为RTF格式的策略
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. PDF → DOCX（使用外部工具，在临时目录）
    3. DOCX → RTF（使用Office支持模块，在临时目录）
    4. 根据配置决定是否保留中间DOCX文件
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - CAJ：先转为PDF再处理（待实现）
    - OFD：先转为PDF再处理（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF到RTF的转换
        
        Args:
            file_path: 输入的PDF文件路径
            options: 转换选项字典
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 预处理：确保文件是PDF格式
                if actual_format and actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, _ = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 第1步：PDF → DOCX（在临时目录，使用标准命名）
                if progress_callback:
                    progress_callback("步骤1/2: PDF转DOCX...")
                
                from gongwen_converter.utils.path_utils import generate_output_path
                
                # 生成中间DOCX的标准化路径（在临时目录）
                docx_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type="docx"
                )
                
                from gongwen_converter.converter.pdf_processing.pdf_converter_utils import convert_pdf_to_docx
                
                docx_path = convert_pdf_to_docx(
                    pdf_path,
                    docx_temp_path,
                    cancel_event
                )
                
                if not docx_path or not os.path.exists(docx_path):
                    return ConversionResult(
                        success=False,
                        message="PDF转DOCX失败，请安装Word、LibreOffice或pdf2docx库"
                    )
                
                logger.info(f"中间DOCX已生成: {os.path.basename(docx_path)}")
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 第2步：DOCX → RTF（在临时目录，使用标准命名）
                if progress_callback:
                    progress_callback("步骤2/2: DOCX转RTF...")
                
                # 生成最终RTF的标准化路径（在临时目录）
                rtf_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type="rtf"
                )
                
                from gongwen_converter.converter.formats.office import convert_docx_to_rtf
                
                rtf_path = convert_docx_to_rtf(docx_path, rtf_temp_path, cancel_event)
                
                if not rtf_path or not os.path.exists(rtf_path):
                    return ConversionResult(
                        success=False,
                        message="DOCX转RTF失败，请安装LibreOffice、Word或WPS"
                    )
                
                logger.info(f"最终RTF已生成: {os.path.basename(rtf_path)}")
                
                # 移动文件到输出目录
                from gongwen_converter.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                final_rtf_path = os.path.join(output_dir, os.path.basename(rtf_path))
                shutil.move(rtf_path, final_rtf_path)
                logger.info(f"RTF文件已移动到: {final_rtf_path}")
                
                # 根据配置决定是否保留中间DOCX文件
                if self._should_keep_intermediates():
                    final_docx_path = os.path.join(output_dir, os.path.basename(docx_path))
                    shutil.move(docx_path, final_docx_path)
                    logger.info(f"保留中间文件: {final_docx_path}")
                else:
                    logger.debug("清理中间DOCX文件")
                
                return ConversionResult(
                    success=True,
                    output_path=final_rtf_path,
                    message="已成功转换为RTF"
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except ImportError as e:
            error_msg = str(e)
            logger.error(f"缺少必要的库: {error_msg}")
            return ConversionResult(success=False, message=error_msg)
        except Exception as e:
            logger.error(f"执行 LayoutToRtfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False


@register_conversion(CATEGORY_LAYOUT, 'md')
class LayoutToMarkdownPymupdf4llmStrategy(BaseStrategy):
    """
    使用pymupdf4llm将PDF转换为Markdown的策略（v2.2简化版）
    
    新设计：所有输出（MD和图片）都放在一个文件夹内
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 生成标准输出路径（含时间戳）
    3. 核心转换：PDF → Markdown（使用pymupdf4llm）
    4. 根据提取选项处理图片和OCR
    
    支持4种提取组合（简化为实际3种）：
    1. ❌图片 ❌OCR：纯文本MD（放在文件夹内）
    2. ✅图片 ❌OCR：MD + 图片（同文件夹）
    3. ✅图片 ✅OCR：MD + 图片 + OCR（同文件夹）
    
    注意：内部总是提取文本，GUI不再显示"提取文字"选项
    
    输出结构：
    ```
    document_20251107_201500_fromPdf/
    ├── document_20251107_201500_fromPdf.md
    ├── image_1.png
    ├── image_2.png
    └── image_3.png
    ```
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - CAJ：先转为PDF再处理（待实现）
    - OFD：先转为PDF再处理（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF到Markdown的转换（使用pymupdf4llm，v2.2简化版）
        
        Args:
            file_path: 输入的PDF文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - actual_format: 实际文件格式
                - extract_images: 是否提取图片（布尔值，默认False）
                - extract_ocr: 是否OCR识别（布尔值，默认False）
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")
        
        # 获取提取选项（只有2个）
        extract_images = options.get("extract_images", False)
        extract_ocr = options.get("extract_ocr", False)
        
        logger.info(
            f"PDF转Markdown - pymupdf4llm模式，"
            f"提取图片={extract_images}, OCR={extract_ocr}"
        )
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 预处理：确保文件是PDF格式
                if actual_format and actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, _ = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 确定输出目录
                from gongwen_converter.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                
                # 生成标准输出路径（不含扩展名，作为文件夹名）
                from gongwen_converter.utils.path_utils import generate_output_path
                
                folder_path_with_ext = generate_output_path(
                    file_path,
                    output_dir,
                    section="",
                    add_timestamp=True,
                    description='fromPdf',
                    file_type='md'
                )
                # 去掉.md扩展名，作为文件夹名
                basename_for_output = os.path.splitext(os.path.basename(folder_path_with_ext))[0]
                
                logger.info(f"输出文件夹基础名: {basename_for_output}")
                
                # 使用pymupdf4llm提取内容
                if progress_callback:
                    progress_callback("提取PDF内容...")
                
                from gongwen_converter.converter.pdf_processing.pdf_pymupdf4llm import extract_pdf_with_pymupdf4llm
                
                result_data = extract_pdf_with_pymupdf4llm(
                    pdf_path,
                    extract_images,
                    extract_ocr,
                    output_dir,
                    basename_for_output,  # 传递标准化的文件夹名
                    cancel_event,
                    progress_callback  # 传递进度回调
                )
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 获取结果路径
                md_path = result_data['md_path']
                folder_path = result_data['folder_path']
                
                logger.info(f"Markdown文件已生成: {md_path}")
                logger.info(f"输出文件夹: {folder_path}")
                logger.info(
                    f"统计信息 - 图片: {result_data['image_count']}, "
                    f"OCR: {result_data['ocr_count']}"
                )
                
                # 构建成功消息
                message_parts = ["已成功导出为Markdown"]
                if result_data['image_count'] > 0:
                    message_parts.append(f"提取了{result_data['image_count']}张图片")
                if result_data['ocr_count'] > 0:
                    message_parts.append(f"识别了{result_data['ocr_count']}个OCR文本")
                
                return ConversionResult(
                    success=True,
                    output_path=md_path,  # 返回MD文件路径
                    message="，".join(message_parts)
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except ImportError as e:
            error_msg = str(e)
            logger.error(f"缺少必要的库: {error_msg}")
            return ConversionResult(success=False, message=error_msg)
        except Exception as e:
            logger.error(f"执行 LayoutToMarkdownPymupdf4llmStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)


@register_conversion(CATEGORY_LAYOUT, 'pdf')
class LayoutToPdfStrategy(BaseStrategy):
    """
    将版式文件（OFD/XPS/CAJ）统一转换为PDF的策略
    
    这是一个统一的入口策略，会根据实际文件格式自动分发到对应的转换函数。
    解决了GUI中convert_layout_to_pdf策略缺失的问题。
    
    转换流程：
    1. 检测或使用提供的actual_format
    2. 根据格式调用对应的转换函数：
       - PDF: 直接返回（无需转换）
       - OFD: 调用ofd_to_pdf
       - XPS: 调用xps_to_pdf
       - CAJ: 调用caj_to_pdf（待实现）
    3. 保存PDF到输出目录
    
    支持的输入格式：
    - PDF：直接返回，无需转换
    - OFD：转换为PDF
    - XPS：转换为PDF
    - CAJ：转换为PDF（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行版式文件到PDF的统一转换
        
        Args:
            file_path: 输入的版式文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - actual_format: 实际文件格式（可选，如果不提供则自动检测）
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")
        
        try:
            # 自动检测格式（如果未提供）
            if not actual_format:
                from gongwen_converter.utils.file_type_utils import detect_actual_file_format
                actual_format = detect_actual_file_format(file_path)
                logger.debug(f"自动检测版式文件格式: {actual_format}")
            else:
                logger.debug(f"使用提供的文件格式: {actual_format}")
            
            # 如果已经是PDF，直接返回成功
            if actual_format == 'pdf':
                logger.info("文件已是PDF格式，无需转换")
                return ConversionResult(
                    success=True,
                    output_path=file_path,
                    message="文件已是PDF格式"
                )
            
            if progress_callback:
                progress_callback(f"准备转换{actual_format.upper()}为PDF...")
            
            # 确定输出目录
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 根据格式调用对应的转换函数
            if actual_format == 'ofd':
                if progress_callback:
                    progress_callback("转换OFD为PDF...")
                
                from gongwen_converter.converter.formats.layout import ofd_to_pdf
                
                result_path = ofd_to_pdf(
                    file_path,
                    cancel_event,
                    output_dir=output_dir
                )
                
            elif actual_format == 'xps':
                if progress_callback:
                    progress_callback("转换XPS为PDF...")
                
                from gongwen_converter.converter.formats.layout import xps_to_pdf
                
                result_path = xps_to_pdf(
                    file_path,
                    cancel_event,
                    output_dir=output_dir
                )
                
            elif actual_format == 'caj':
                if progress_callback:
                    progress_callback("转换CAJ为PDF...")
                
                from gongwen_converter.converter.formats.layout import caj_to_pdf
                
                result_path = caj_to_pdf(
                    file_path,
                    cancel_event,
                    output_dir=output_dir
                )
                
            else:
                return ConversionResult(
                    success=False,
                    message=f"不支持的版式文件格式: {actual_format}"
                )
            
            # 检查取消
            if cancel_event and cancel_event.is_set():
                return ConversionResult(success=False, message="操作已取消")
            
            # 检查转换结果
            if not result_path or not os.path.exists(result_path):
                return ConversionResult(
                    success=False,
                    message=f"{actual_format.upper()}转PDF失败"
                )
            
            logger.info(f"{actual_format.upper()}转PDF成功: {result_path}")
            
            return ConversionResult(
                success=True,
                output_path=result_path,
                message=f"已成功转换为PDF"
            )
            
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except NotImplementedError as e:
            logger.error(f"{actual_format.upper() if actual_format else '版式文件'}转PDF功能尚未实现: {e}")
            return ConversionResult(
                success=False,
                message=f"{actual_format.upper() if actual_format else '该格式'}转PDF功能尚未实现"
            )
        except Exception as e:
            logger.error(f"执行 LayoutToPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {str(e)}",
                error=e
            )


@register_conversion('ofd', 'pdf')
class OfdToPdfStrategy(BaseStrategy):
    """
    将OFD文件转换为PDF的策略
    
    转换流程：
    1. 调用底层ofd_to_pdf转换函数
    2. 保存PDF到输出目录（与原文件同目录）
    
    特点：
    - 直接保存为PDF，不作为中间文件
    - 支持取消操作
    - 提供进度反馈
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行OFD到PDF的转换
        
        Args:
            file_path: 输入的OFD文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        
        try:
            if progress_callback:
                progress_callback("准备转换OFD为PDF...")
            
            # 确定输出目录和文件名
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            basename = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(output_dir, f"{basename}.pdf")
            
            # 调用底层转换函数
            if progress_callback:
                progress_callback("转换OFD为PDF...")
            
            from gongwen_converter.converter.formats.layout import ofd_to_pdf
            
            result_path = ofd_to_pdf(
                file_path,
                cancel_event,
                output_dir=output_dir
            )
            
            # 检查取消
            if cancel_event and cancel_event.is_set():
                return ConversionResult(success=False, message="操作已取消")
            
            if not result_path or not os.path.exists(result_path):
                return ConversionResult(
                    success=False,
                    message="OFD转PDF失败"
                )
            
            # 如果输出路径不同，移动文件
            if result_path != output_path:
                shutil.move(result_path, output_path)
                logger.info(f"PDF文件已移动到: {output_path}")
            
            logger.info(f"OFD转PDF成功: {output_path}")
            
            return ConversionResult(
                success=True,
                output_path=output_path,
                message="已成功转换为PDF"
            )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except NotImplementedError as e:
            logger.error(f"OFD转PDF功能尚未实现: {e}")
            return ConversionResult(success=False, message=f"OFD转PDF功能尚未实现: {str(e)}")
        except Exception as e:
            logger.error(f"执行 OfdToPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)


@register_conversion('xps', 'pdf')
class XpsToPdfStrategy(BaseStrategy):
    """
    将XPS文件转换为PDF的策略
    
    转换流程：
    1. 调用底层xps_to_pdf转换函数
    2. 保存PDF到输出目录（与原文件同目录）
    
    特点：
    - 直接保存为PDF，不作为中间文件
    - 支持取消操作
    - 提供进度反馈
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行XPS到PDF的转换
        
        Args:
            file_path: 输入的XPS文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        
        try:
            if progress_callback:
                progress_callback("准备转换XPS为PDF...")
            
            # 确定输出目录和文件名
            from gongwen_converter.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            basename = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(output_dir, f"{basename}.pdf")
            
            # 调用底层转换函数
            if progress_callback:
                progress_callback("转换XPS为PDF...")
            
            from gongwen_converter.converter.formats.layout import xps_to_pdf
            
            result_path = xps_to_pdf(
                file_path,
                cancel_event,
                output_dir=output_dir
            )
            
            # 检查取消
            if cancel_event and cancel_event.is_set():
                return ConversionResult(success=False, message="操作已取消")
            
            if not result_path or not os.path.exists(result_path):
                return ConversionResult(
                    success=False,
                    message="XPS转PDF失败"
                )
            
            # 如果输出路径不同，移动文件
            if result_path != output_path:
                shutil.move(result_path, output_path)
                logger.info(f"PDF文件已移动到: {output_path}")
            
            logger.info(f"XPS转PDF成功: {output_path}")
            
            return ConversionResult(
                success=True,
                output_path=output_path,
                message="已成功转换为PDF"
            )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except Exception as e:
            logger.error(f"执行 XpsToPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)


@register_action("merge_pdfs")
class MergePdfsStrategy(BaseStrategy):
    """
    合并多个PDF/版式文件为一个PDF的策略
    
    触发条件：批量模式下，版式文件数量 > 1
    
    转换流程：
    1. 预处理：将所有非PDF格式文件转换为PDF（XPS/OFD/CAJ）
    2. 按文件列表顺序合并PDF
    3. 输出到第一个文件所在目录
    
    命名规则：
    - 输出文件名：合并文档_{时间戳}.pdf
    
    支持的输入格式：
    - PDF：直接合并
    - XPS：先转为PDF再合并
    - OFD：先转为PDF再合并（待实现）
    - CAJ：先转为PDF再合并（待实现）
    - 混合：支持以上格式混合合并
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF合并操作
        
        Args:
            file_path: 第一个文件的路径（用于确定输出目录）
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - file_list: 要合并的文件列表（必需）
                - actual_formats: 各文件的实际格式列表（可选）
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        file_list = options.get("file_list", [])
        actual_formats = options.get("actual_formats", [])
        
        if not file_list:
            return ConversionResult(
                success=False,
                message="未提供要合并的文件列表"
            )
        
        if len(file_list) < 2:
            return ConversionResult(
                success=False,
                message="至少需要2个文件才能合并"
            )
        
        try:
            if progress_callback:
                progress_callback(f"准备合并{len(file_list)}个文件...")
            
            logger.info(f"开始合并{len(file_list)}个PDF/版式文件")
            
            # 使用临时目录处理中间文件
            with tempfile.TemporaryDirectory() as temp_dir:
                pdf_paths = []
                
                # 步骤1：预处理所有文件，确保都是PDF格式
                for i, input_file in enumerate(file_list):
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(success=False, message="操作已取消")
                    
                    file_name = os.path.basename(input_file)
                    actual_format = actual_formats[i] if i < len(actual_formats) else None
                    
                    if progress_callback:
                        progress_callback(f"处理文件 {i+1}/{len(file_list)}: {file_name}")
                    
                    # 自动检测格式（如果未提供）
                    if not actual_format:
                        from gongwen_converter.utils.file_type_utils import detect_actual_file_format
                        actual_format = detect_actual_file_format(input_file)
                    
                    # 如果是PDF，直接使用；否则先转换
                    if actual_format == 'pdf':
                        pdf_paths.append(input_file)
                        logger.debug(f"文件{i+1}已是PDF: {file_name}")
                    else:
                        logger.info(f"文件{i+1}需要转换: {file_name} ({actual_format} -> PDF)")
                        pdf_path, _ = _preprocess_layout_file(
                            input_file, temp_dir, cancel_event, actual_format
                        )
                        pdf_paths.append(pdf_path)
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：合并所有PDF
                if progress_callback:
                    progress_callback("合并PDF文件...")
                
                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    return ConversionResult(
                        success=False,
                        message="缺少PyMuPDF库，请安装：pip install PyMuPDF"
                    )
                
                # 创建输出PDF
                merged_pdf = fitz.open()
                
                try:
                    for i, pdf_path in enumerate(pdf_paths):
                        if cancel_event and cancel_event.is_set():
                            merged_pdf.close()
                            return ConversionResult(success=False, message="操作已取消")
                        
                        if progress_callback:
                            progress_callback(f"合并文件 {i+1}/{len(pdf_paths)}")
                        
                        # 打开并插入PDF
                        with fitz.open(pdf_path) as pdf:
                            merged_pdf.insert_pdf(pdf)
                        
                        logger.debug(f"已合并文件{i+1}: {os.path.basename(pdf_path)}")
                    
                    # 生成输出路径（使用选中文件所在目录）
                    from gongwen_converter.utils.workspace_manager import get_output_directory
                    output_dir = get_output_directory(file_path)
                    import datetime
                    # 生成输出路径（使用选中文件所在目录）
                    from gongwen_converter.utils.workspace_manager import get_output_directory
                    selected_file = options.get("selected_file", file_path)
                    output_dir = get_output_directory(selected_file)
                    logger.info(f"输出目录基于: {os.path.basename(selected_file)}")
                    
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_path = os.path.join(output_dir, f"合并PDF_{timestamp}.pdf")
                    
                    # 保存合并后的PDF
                    merged_pdf.save(output_path)
                    logger.info(f"合并PDF成功，共{len(merged_pdf)}页: {output_path}")
                    
                    return ConversionResult(
                        success=True,
                        output_path=output_path,
                        message=f"已成功合并{len(file_list)}个文件，共{len(merged_pdf)}页"
                    )
                
                finally:
                    merged_pdf.close()
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except Exception as e:
            logger.error(f"执行 MergePdfsStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"合并失败: {str(e)}", error=e)


@register_action("split_pdf")
class SplitPdfStrategy(BaseStrategy):
    """
    拆分PDF/版式文件为两个PDF的策略
    
    触发条件：单个版式文件 + 页码输入合法
    
    转换流程：
    1. 预处理：如果非PDF格式，先转换为PDF（XPS/OFD/CAJ）
    2. 根据用户输入的页码拆分为两个PDF文件：
       - 第1个文件：用户输入的页码
       - 第2个文件：剩余页码（如果有）
    3. 输出到原文件所在目录
    
    命名规则：
    - 第1个文件：原文件名_拆分1_{时间戳}_split.pdf
    - 第2个文件：原文件名_拆分2_{时间戳}_split.pdf（如果有剩余页）
    
    支持的输入格式：
    - PDF：直接拆分
    - XPS：先转为PDF再拆分
    - OFD：先转为PDF再拆分（待实现）
    - CAJ：先转为PDF再拆分（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF拆分操作
        
        Args:
            file_path: 输入的PDF/版式文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - pages: 要提取的页码列表（必需，已排序去重）
                - total_pages: PDF总页数（可选，用于验证）
                - actual_format: 实际文件格式（可选）
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        pages = options.get("pages", [])
        total_pages = options.get("total_pages", 0)
        actual_format = options.get("actual_format")
        
        if not pages:
            return ConversionResult(
                success=False,
                message="未提供要拆分的页码"
            )
        
        try:
            if progress_callback:
                progress_callback("准备拆分PDF...")
            
            logger.info(f"开始拆分PDF，页码: {pages}")
            
            # 使用临时目录处理中间文件
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理，确保文件是PDF格式
                if actual_format and actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, _ = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：打开PDF并验证页码
                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    return ConversionResult(
                        success=False,
                        message="缺少PyMuPDF库，请安装：pip install PyMuPDF"
                    )
                
                with fitz.open(pdf_path) as pdf:
                    actual_total_pages = len(pdf)
                    logger.info(f"PDF总页数: {actual_total_pages}")
                    
                    # 验证并过滤页码
                    valid_pages = [p for p in pages if 1 <= p <= actual_total_pages]
                    if not valid_pages:
                        return ConversionResult(
                            success=False,
                            message=f"所有页码均无效（PDF共{actual_total_pages}页）"
                        )
                    
                    # 计算剩余页码
                    all_pages = set(range(1, actual_total_pages + 1))
                    remaining_pages = sorted(list(all_pages - set(valid_pages)))
                    
                    logger.info(f"第1个文件页码: {valid_pages}")
                    logger.info(f"第2个文件页码: {remaining_pages if remaining_pages else '无'}")
                    
                    # 验证：确保有剩余页码（不能拆分全部页码）
                    if not remaining_pages:
                        return ConversionResult(
                            success=False,
                            message="拆分失败：输入的页码涵盖了全部页码，无法拆分"
                        )
                    
                    # 检查取消
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(success=False, message="操作已取消")
                    
                    # 步骤3：生成输出路径
                    from gongwen_converter.utils.path_utils import generate_output_path
                    from gongwen_converter.utils.workspace_manager import get_output_directory
                    output_dir = get_output_directory(file_path)
                    
                    # 第1个文件路径
                    output_path1 = generate_output_path(
                        file_path,
                        output_dir=output_dir,
                        section="拆分1",
                        add_timestamp=True,
                        description="split",
                        file_type="pdf"
                    )
                    
                    # 步骤4：创建第1个PDF（用户选择的页码）
                    if progress_callback:
                        progress_callback("创建第1个PDF...")
                    
                    pdf1 = fitz.open()
                    try:
                        for page_num in valid_pages:
                            pdf1.insert_pdf(pdf, from_page=page_num-1, to_page=page_num-1)
                        
                        pdf1.save(output_path1)
                        logger.info(f"第1个PDF已保存: {output_path1}，共{len(pdf1)}页")
                    finally:
                        pdf1.close()
                    
                    # 检查取消
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(success=False, message="操作已取消")
                    
                    # 步骤5：创建第2个PDF（剩余页码，如果有）
                    output_path2 = None
                    if remaining_pages:
                        if progress_callback:
                            progress_callback("创建第2个PDF...")
                        
                        # 第2个文件路径
                        output_path2 = generate_output_path(
                            file_path,
                            output_dir=output_dir,
                            section="拆分2",
                            add_timestamp=True,
                            description="split",
                            file_type="pdf"
                        )
                        
                        pdf2 = fitz.open()
                        try:
                            for page_num in remaining_pages:
                                pdf2.insert_pdf(pdf, from_page=page_num-1, to_page=page_num-1)
                            
                            pdf2.save(output_path2)
                            logger.info(f"第2个PDF已保存: {output_path2}，共{len(pdf2)}页")
                        finally:
                            pdf2.close()
                    else:
                        logger.info("无剩余页码，不生成第2个PDF")
                
                # 构建结果消息
                if output_path2:
                    message = f"已成功拆分为2个文件（第1个: {len(valid_pages)}页，第2个: {len(remaining_pages)}页）"
                else:
                    message = f"已成功拆分（仅1个文件，{len(valid_pages)}页）"
                
                return ConversionResult(
                    success=True,
                    output_path=output_path1,  # 返回第1个文件路径
                    message=message
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except Exception as e:
            logger.error(f"执行 SplitPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"拆分失败: {str(e)}", error=e)


@register_conversion(CATEGORY_LAYOUT, 'png')
class LayoutToPngStrategy(BaseStrategy):
    """
    将版式文件的每一页转换为PNG图片的策略
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 创建输出子文件夹（在临时目录）
    3. 使用PyMuPDF统一渲染所有页面为PNG图片（按指定DPI）
    4. 将子文件夹移动到目标目录
    5. 根据配置决定是否保留中间PDF文件
    
    输出结构：
    原文件名_时间戳_from原格式/
    ├── page_01.png
    ├── page_02.png
    └── page_03.png
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - OFD：先转为PDF再处理
    - CAJ：先转为PDF再处理
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行版式文件到PNG图片的转换
        
        Args:
            file_path: 输入的版式文件路径
            options: 转换选项字典，包含：
                - dpi: 渲染DPI（150/300/600）
                - actual_format: 实际文件格式
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        dpi = options.get('dpi', 150)
        actual_format = options.get('actual_format', 'pdf')
        cancel_event = options.get('cancel_event')
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            logger.info(f"开始转换 {actual_format.upper()} 为PNG图片，DPI: {dpi}")
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 确保文件是PDF格式
                if actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, intermediate_pdf = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                    intermediate_pdf = None
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成输出文件夹路径
                from gongwen_converter.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                
                from gongwen_converter.utils.path_utils import generate_output_path
                
                # 生成标准化文件夹名（使用原格式）
                description = f"from{actual_format.capitalize()}"
                folder_path_with_ext = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="png"
                )
                
                # 去掉.png扩展名，作为文件夹名
                folder_name = os.path.splitext(os.path.basename(folder_path_with_ext))[0]
                temp_output_folder = os.path.join(temp_dir, folder_name)
                os.makedirs(temp_output_folder, exist_ok=True)
                
                logger.info(f"输出文件夹: {folder_name}")
                
                # 步骤3：打开PDF并渲染所有页面
                if progress_callback:
                    progress_callback("渲染PDF页面为PNG...")
                
                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    return ConversionResult(
                        success=False,
                        message="缺少PyMuPDF库，请安装：pip install PyMuPDF"
                    )
                
                with fitz.open(pdf_path) as doc:
                    total_pages = len(doc)
                    width = len(str(total_pages))  # 计算前导零位数
                    
                    logger.info(f"PDF共{total_pages}页，开始渲染（DPI: {dpi}）")
                    
                    # 计算缩放比例
                    zoom = dpi / 72.0
                    
                    for page_num in range(total_pages):
                        # 检查取消
                        if cancel_event and cancel_event.is_set():
                            return ConversionResult(success=False, message="操作已取消")
                        
                        if progress_callback:
                            progress_callback(f"渲染第 {page_num + 1}/{total_pages} 页...")
                        
                        page = doc[page_num]
                        
                        # 渲染页面
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=True)
                        
                        # 生成文件名（带前导零）
                        image_filename = f"page_{str(page_num + 1).zfill(width)}.png"
                        image_path = os.path.join(temp_output_folder, image_filename)
                        
                        # 保存图片
                        pix.save(image_path)
                        
                        logger.debug(f"已保存: {image_filename} ({pix.width}x{pix.height})")
                
                logger.info(f"所有页面已渲染完成，共{total_pages}张图片")
                
                # 步骤4：移动文件夹到目标目录
                final_folder = os.path.join(output_dir, folder_name)
                
                if os.path.exists(final_folder):
                    shutil.rmtree(final_folder)
                    logger.debug(f"已删除现有文件夹: {final_folder}")
                
                shutil.move(temp_output_folder, final_folder)
                logger.info(f"文件夹已移动到: {final_folder}")
                
                # 步骤5：处理中间文件
                if intermediate_pdf and self._should_keep_intermediates():
                    intermediate_output = os.path.join(output_dir, os.path.basename(intermediate_pdf))
                    shutil.move(intermediate_pdf, intermediate_output)
                    logger.info(f"保留中间PDF文件: {intermediate_output}")
                
                # 返回文件夹路径作为输出路径
                return ConversionResult(
                    success=True,
                    output_path=final_folder,
                    message=f"已成功转换为PNG图片，共{total_pages}张"
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except Exception as e:
            logger.error(f"执行 LayoutToPngStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from gongwen_converter.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False


@register_conversion(CATEGORY_LAYOUT, 'jpg')
class LayoutToJpgStrategy(BaseStrategy):
    """
    将版式文件的每一页转换为JPG图片的策略
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 创建输出子文件夹（在临时目录）
    3. 使用PyMuPDF统一渲染所有页面为JPG图片（按指定DPI，处理透明度）
    4. 将子文件夹移动到目标目录
    5. 根据配置决定是否保留中间PDF文件
    
    输出结构：
    原文件名_时间戳_from原格式/
    ├── page_01.jpg
    ├── page_02.jpg
    └── page_03.jpg
    
    特点：
    - JPG不支持透明，透明背景会转为白色
    - 文件体积比PNG小
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - OFD：先转为PDF再处理
    - CAJ：先转为PDF再处理
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行版式文件到JPG图片的转换
        
        Args:
            file_path: 输入的版式文件路径
            options: 转换选项字典，包含：
                - dpi: 渲染DPI（150/300/600）
                - actual_format: 实际文件格式
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        dpi = options.get('dpi', 150)
        actual_format = options.get('actual_format', 'pdf')
        cancel_event = options.get('cancel_event')
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            logger.info(f"开始转换 {actual_format.upper()} 为JPG图片，DPI: {dpi}")
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 确保文件是PDF格式
                if actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, intermediate_pdf = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                    intermediate_pdf = None
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成输出文件夹路径
                from gongwen_converter.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                
                from gongwen_converter.utils.path_utils import generate_output_path
                
                # 生成标准化文件夹名（使用原格式）
                description = f"from{actual_format.capitalize()}"
                folder_path_with_ext = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="jpg"
                )
                
                # 去掉.jpg扩展名，作为文件夹名
                folder_name = os.path.splitext(os.path.basename(folder_path_with_ext))[0]
                temp_output_folder = os.path.join(temp_dir, folder_name)
                os.makedirs(temp_output_folder, exist_ok=True)
                
                logger.info(f"输出文件夹: {folder_name}")
                
                # 步骤3：打开PDF并渲染所有页面
                if progress_callback:
                    progress_callback("渲染PDF页面为JPG...")
                
                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    return ConversionResult(
                        success=False,
                        message="缺少PyMuPDF库，请安装：pip install PyMuPDF"
                    )
                
                with fitz.open(pdf_path) as doc:
                    total_pages = len(doc)
                    width = len(str(total_pages))  # 计算前导零位数
                    
                    logger.info(f"PDF共{total_pages}页，开始渲染（DPI: {dpi}）")
                    
                    # 计算缩放比例
                    zoom = dpi / 72.0
                    
                    for page_num in range(total_pages):
                        # 检查取消
                        if cancel_event and cancel_event.is_set():
                            return ConversionResult(success=False, message="操作已取消")
                        
                        if progress_callback:
                            progress_callback(f"渲染第 {page_num + 1}/{total_pages} 页...")
                        
                        page = doc[page_num]
                        
                        # 渲染页面（JPG不支持alpha通道，设置为False）
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        
                        # 生成文件名（带前导零）
                        image_filename = f"page_{str(page_num + 1).zfill(width)}.jpg"
                        image_path = os.path.join(temp_output_folder, image_filename)
                        
                        # 保存图片
                        pix.save(image_path)
                        
                        logger.debug(f"已保存: {image_filename} ({pix.width}x{pix.height})")
                
                logger.info(f"所有页面已渲染完成，共{total_pages}张图片")
                
                # 步骤4：移动文件夹到目标目录
                final_folder = os.path.join(output_dir, folder_name)
                
                if os.path.exists(final_folder):
                    shutil.rmtree(final_folder)
                    logger.debug(f"已删除现有文件夹: {final_folder}")
                
                shutil.move(temp_output_folder, final_folder)
                logger.info(f"文件夹已移动到: {final_folder}")
                
                # 步骤5：处理中间文件
                if intermediate_pdf and self._should_keep_intermediates():
                    intermediate_output = os.path.join(output_dir, os.path.basename(intermediate_pdf))
                    shutil.move(intermediate_pdf, intermediate_output)
                    logger.info(f"保留中间PDF文件: {intermediate_output}")
                
                # 返回文件夹路径作为输出路径
                return ConversionResult(
                    success=True,
                    output_path=final_folder,
                    message=f"已成功转换为JPG图片，共{total_pages}张"
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except Exception as e:
            logger.error(f"执行 LayoutToJpgStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)
    
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


@register_conversion(CATEGORY_LAYOUT, 'tif')
class LayoutToTifStrategy(BaseStrategy):
    """
    将版式文件转换为TIF图片的策略（支持多页TIFF）
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 使用PyMuPDF渲染所有页面为图片
    3. 使用PIL保存为TIF（自动多页）
    4. 输出单个TIF文件（无需子文件夹）
    
    特点：
    - 自动判断单页/多页
    - 多页自动打包为单个TIFF文件
    - 无需子文件夹，直接输出文件
    - alpha=False，透明背景转为白色
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - OFD：先转为PDF再处理
    - CAJ：先转为PDF再处理
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行版式文件到TIF图片的转换
        
        Args:
            file_path: 输入的版式文件路径
            options: 转换选项字典，包含：
                - dpi: 渲染DPI（150/300/600）
                - actual_format: 实际文件格式
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        dpi = options.get('dpi', 150)
        actual_format = options.get('actual_format', 'pdf')
        cancel_event = options.get('cancel_event')
        
        try:
            if progress_callback:
                progress_callback("准备转换...")
            
            logger.info(f"开始转换 {actual_format.upper()} 为TIF图片，DPI: {dpi}")
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 确保文件是PDF格式
                if actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(f"转换{actual_format.upper()}为PDF...")
                    pdf_path, intermediate_pdf = _preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                    intermediate_pdf = None
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤2：生成输出文件路径（直接在临时目录，无需子文件夹）
                from gongwen_converter.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                
                from gongwen_converter.utils.path_utils import generate_output_path
                
                # 生成标准化TIF路径（在临时目录）
                description = f"from{actual_format.capitalize()}"
                tif_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="tif"
                )
                
                logger.info(f"输出TIF文件: {os.path.basename(tif_temp_path)}")
                
                # 步骤3：打开PDF并渲染所有页面
                if progress_callback:
                    progress_callback("渲染PDF页面...")
                
                try:
                    import fitz  # PyMuPDF
                    from PIL import Image
                except ImportError as e:
                    missing_lib = "PyMuPDF" if "fitz" in str(e) else "PIL"
                    return ConversionResult(
                        success=False,
                        message=f"缺少{missing_lib}库，请安装：pip install {missing_lib}"
                    )
                
                images = []
                
                with fitz.open(pdf_path) as doc:
                    total_pages = len(doc)
                    logger.info(f"PDF共{total_pages}页，开始渲染（DPI: {dpi}）")
                    
                    # 计算缩放比例
                    zoom = dpi / 72.0
                    
                    for page_num in range(total_pages):
                        # 检查取消
                        if cancel_event and cancel_event.is_set():
                            return ConversionResult(success=False, message="操作已取消")
                        
                        if progress_callback:
                            progress_callback(f"渲染第 {page_num + 1}/{total_pages} 页...")
                        
                        page = doc[page_num]
                        
                        # 渲染页面（TIF不需要透明通道，设置alpha=False）
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        
                        # 转换为PIL Image对象
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        images.append(img)
                        
                        logger.debug(f"已渲染第{page_num + 1}页 ({pix.width}x{pix.height})")
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message="操作已取消")
                
                # 步骤4：保存为TIF（自动多页）
                if progress_callback:
                    progress_callback("保存TIF文件...")
                
                if len(images) > 1:
                    # 多页TIF：使用save_all参数
                    images[0].save(
                        tif_temp_path,
                        format='TIFF',
                        save_all=True,
                        append_images=images[1:],
                        compression='tiff_lzw'  # 使用LZW无损压缩
                    )
                    logger.info(f"多页TIF已保存: {os.path.basename(tif_temp_path)}，共{len(images)}页")
                else:
                    # 单页TIF
                    images[0].save(
                        tif_temp_path,
                        format='TIFF',
                        compression='tiff_lzw'
                    )
                    logger.info(f"单页TIF已保存: {os.path.basename(tif_temp_path)}")
                
                # 步骤5：移动文件到目标目录
                final_tif_path = os.path.join(output_dir, os.path.basename(tif_temp_path))
                shutil.move(tif_temp_path, final_tif_path)
                logger.info(f"TIF文件已移动到: {final_tif_path}")
                
                # 步骤6：处理中间文件
                if intermediate_pdf and self._should_keep_intermediates():
                    intermediate_output = os.path.join(output_dir, os.path.basename(intermediate_pdf))
                    shutil.move(intermediate_pdf, intermediate_output)
                    logger.info(f"保留中间PDF文件: {intermediate_output}")
                
                # 返回文件路径作为输出路径
                page_desc = f"{total_pages}页" if total_pages > 1 else "1页"
                return ConversionResult(
                    success=True,
                    output_path=final_tif_path,
                    message=f"已成功转换为TIF图片（{page_desc}）"
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message="操作已取消")
        except Exception as e:
            logger.error(f"执行 LayoutToTifStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)
    
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


# 模块测试
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s"
    )
    
    logger.info("版式文件策略模块测试开始")
    
    # 测试预处理函数（预期PDF直接返回）
    test_pdf = "test.pdf"
    if os.path.exists(test_pdf):
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                result = _preprocess_layout_file(test_pdf, temp_dir)
                logger.info(f"预处理测试成功: {result}")
        except Exception as e:
            logger.error(f"预处理测试失败: {e}")
    else:
        logger.warning(f"测试文件不存在: {test_pdf}")
    
    logger.info("版式文件策略模块测试结束")
