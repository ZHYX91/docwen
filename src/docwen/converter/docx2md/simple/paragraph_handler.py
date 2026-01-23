"""
段落处理器模块

负责处理各种类型的Word段落，将其转换为Markdown格式。

主要功能：
- ParagraphHandler: 段落处理器类，统一处理所有段落类型
- 支持的段落类型：代码块、引用块、标题、列表、公式、普通段落

设计说明：
    此模块从 converter.py 的主循环中提取段落处理逻辑，
    使主循环更加简洁，段落处理逻辑更易于测试和维护。

依赖：
- shared/conversion_context.py: 代码块状态管理
- shared/list_processor.py: 列表上下文管理
- shared/break_processor.py: 边框组处理
- utils/docx_utils.py: 段落检测和转换函数
"""

import re
import logging
from typing import Optional, Callable, List, Tuple
import threading

from ..shared.conversion_context import ConversionContext
from ..shared.list_processor import (
    detect_list_item,
    get_list_marker,
    ListCounterManager,
    ListContextManager,
)
from ..shared.break_processor import (
    BorderGroupTracker,
    detect_page_break_in_run,
    detect_section_break_in_paragraph,
)
from ..shared.style_detector import detect_paragraph_style_type
from ..shared.markdown_converter import (
    convert_paragraph_to_markdown,
    convert_paragraph_to_markdown_skip_prefix,
    convert_paragraph_to_markdown_with_styles,
    has_paragraph_gray_shading,
    apply_format_markers,
)
from ..shared.formula_processor import (
    has_formulas_in_paragraph,
    process_paragraph_with_formulas,
    is_formula_supported
)
from ..shared.image_processor import process_image_with_ocr, get_paragraph_images

logger = logging.getLogger(__name__)


class ParagraphHandler:
    """
    段落处理器
    
    统一处理各种类型的Word段落，将其转换为Markdown格式。
    
    支持的段落类型：
    - 代码块（段落样式或灰色底纹）
    - 引用块（quote 1-9 样式）
    - 标题（Heading 1-6, Title, Subtitle）
    - 列表（有序和无序）
    - 公式（OMML格式）
    - 普通段落
    
    使用方式：
        handler = ParagraphHandler(
            ctx=ConversionContext(config_manager),
            list_context_manager=ListContextManager(list_ranges),
            list_counter_manager=ListCounterManager(),
            border_tracker=BorderGroupTracker(),
            config_manager=config_manager,
            note_extractor=note_extractor
        )
        
        for idx, para in enumerate(doc.paragraphs):
            results = handler.process_paragraph(para, idx, skip_indices)
            content_lines.extend(results)
    
    注意：
        - 处理器维护内部状态（代码块、边框组等）
        - 每个文档应创建新的处理器实例
    """
    
    def __init__(
        self,
        ctx: ConversionContext,
        list_context_manager: ListContextManager,
        list_counter_manager: ListCounterManager,
        border_tracker: BorderGroupTracker,
        config_manager,
        note_extractor=None,
        list_indent_spaces: int = 4,
        preserve_formatting: bool = True,
        preserve_heading_formatting: bool = True,
        syntax_config: dict = None,
        wps_shading_enabled: bool = True,
        word_shading_enabled: bool = True,
        remove_numbering: bool = False,
        add_numbering: bool = False,
        formatter=None
    ):
        """
        初始化段落处理器
        
        参数:
            ctx: ConversionContext - 转换上下文，管理代码块状态
            list_context_manager: ListContextManager - 列表上下文管理器
            list_counter_manager: ListCounterManager - 列表编号计数器
            border_tracker: BorderGroupTracker - 边框组跟踪器
            config_manager: ConfigManager - 配置管理器
            note_extractor: NoteExtractor - 脚注/尾注提取器（可选）
            list_indent_spaces: int - 列表缩进空格数（默认4）
            preserve_formatting: bool - 是否保留正文格式（默认True）
            preserve_heading_formatting: bool - 是否保留标题格式（默认True）
            syntax_config: dict - 语法配置字典（可选）
            wps_shading_enabled: bool - 是否启用WPS底纹检测（默认True）
            word_shading_enabled: bool - 是否启用Word底纹检测（默认True）
            remove_numbering: bool - 是否清除原有序号（默认False）
            add_numbering: bool - 是否添加新序号（默认False）
            formatter: HeadingFormatter - 标题序号格式化器（可选）
        """
        self.ctx = ctx
        self.list_context_manager = list_context_manager
        self.list_counter_manager = list_counter_manager
        self.border_tracker = border_tracker
        self.config_manager = config_manager
        self.note_extractor = note_extractor
        self.list_indent_spaces = list_indent_spaces
        self.preserve_formatting = preserve_formatting
        self.preserve_heading_formatting = preserve_heading_formatting
        self.syntax_config = syntax_config or {}
        self.wps_shading_enabled = wps_shading_enabled
        self.word_shading_enabled = word_shading_enabled
        self.remove_numbering = remove_numbering
        self.add_numbering = add_numbering
        self.formatter = formatter
        
        logger.debug(f"ParagraphHandler 初始化完成: list_indent={list_indent_spaces}, "
                     f"preserve_formatting={preserve_formatting}")
    
    def process_paragraph(
        self,
        para,
        idx: int,
        skip_indices: list,
        images_info: list = None,
        keep_images: bool = True,
        enable_ocr: bool = False,
        output_folder: str = None,
        progress_callback: Callable = None,
        cancel_event: threading.Event = None,
        current_image_index: int = 0,
        total_images: int = 0
    ) -> Tuple[List[str], int]:
        """
        处理单个段落，返回Markdown内容行列表
        
        参数:
            para: Paragraph - Word段落对象
            idx: int - 段落索引
            skip_indices: list - 需要跳过的段落索引列表
            images_info: list - 图片信息列表（可选）
            keep_images: bool - 是否保留图片（默认True）
            enable_ocr: bool - 是否启用OCR（默认False）
            output_folder: str - 输出文件夹路径（可选）
            progress_callback: Callable - 进度回调函数（可选）
            cancel_event: threading.Event - 取消事件（可选）
            current_image_index: int - 当前图片索引（用于进度显示）
            total_images: int - 图片总数（用于进度显示）
        
        返回:
            Tuple[List[str], int]: (内容行列表, 更新后的current_image_index)
        
        注意:
            - 跳过skip_indices中的段落
            - 处理器会自动维护代码块、边框组等状态
        """
        if images_info is None:
            images_info = []
        
        content_lines = []
        
        # 跳过已提取到YAML的段落
        if idx in skip_indices:
            logger.debug(f"跳过已提取到YAML的段落: {idx+1}")
            return content_lines, current_image_index
        
        # ==== 边框组处理 ====
        border_separators = self.border_tracker.process_paragraph(para, self.config_manager)
        content_lines.extend(border_separators)
        
        # ==== 分节符检测 ====
        pending_section_break = self._detect_section_break(para, idx)
        
        # ==== 分页符检测 ====
        has_page_break, page_break_separator = self._detect_page_break(para, idx)
        
        # ==== 公式检测 ====
        result, current_image_index = self._try_handle_formula(
            para, idx, current_image_index
        )
        if result is not None:
            content_lines.extend(result)
            # 公式段落也要处理分节符
            if pending_section_break:
                content_lines.append(pending_section_break)
            return content_lines, current_image_index
        
        # ==== 段落样式检测（代码块/引用块）====
        result = self._try_handle_style(para, idx)
        if result is not None:
            content_lines.extend(result)
            # 代码块/引用块也要处理分节符
            if pending_section_break:
                content_lines.append(pending_section_break)
            return content_lines, current_image_index
        
        # ==== 底纹检测（兼容旧文档）====
        result = self._try_handle_shading(para, idx)
        if result is not None:
            content_lines.extend(result)
            if pending_section_break:
                content_lines.append(pending_section_break)
            return content_lines, current_image_index
        
        # ==== 结束代码块（如果有）====
        if self.ctx.in_code_block:
            code_block = self.ctx.end_code_block()
            if code_block:
                content_lines.append(code_block)
                logger.debug("代码块结束")
        
        # ==== 获取段落文本 ====
        para_text_raw = para.text.strip()
        para_text_formatted, inline_style_type, inline_style_value = self._get_formatted_text(para)
        
        # 如果整段都是代码字符样式，作为代码块处理
        if inline_style_type == 'code_block':
            list_context = self.list_context_manager.get_context_for_para(idx, para)
            self.ctx.start_code_block(list_context)
            self.ctx.add_code_line(para.text)
            logger.debug(f"段落 {idx+1}: 整段字符样式检测为代码块")
            if pending_section_break:
                content_lines.append(pending_section_break)
            return content_lines, current_image_index
        
        # ==== 处理空段落 ====
        if not para_text_raw:
            result, current_image_index = self._handle_empty_paragraph(
                para, idx, images_info, keep_images, enable_ocr, output_folder,
                progress_callback, cancel_event, current_image_index, total_images,
                has_page_break, page_break_separator, pending_section_break
            )
            content_lines.extend(result)
            return content_lines, current_image_index
        
        # ==== 处理非空段落 ====
        style_name = para.style.name if para.style else None
        
        # 标题处理
        if style_name and style_name.startswith('Heading'):
            result = self._handle_heading(para, para_text_raw, para_text_formatted, style_name, idx)
            content_lines.extend(result)
        elif style_name == 'Title':
            result = self._handle_title(para, para_text_raw, idx)
            content_lines.extend(result)
        elif style_name == 'Subtitle':
            result = self._handle_subtitle(para, para_text_raw, idx)
            content_lines.extend(result)
        else:
            # 列表或普通段落
            result = self._handle_list_or_normal(
                para, para_text_formatted, idx,
                has_page_break, page_break_separator
            )
            content_lines.extend(result)
        
        # ==== 处理段落后的图片 ====
        para_images = get_paragraph_images(para, images_info)
        for img in para_images:
            current_image_index += 1
            filename = img.get('filename', 'unknown')
            logger.info(f"在段落 {idx + 1} 后匹配到图片: {filename}")
            
            image_link = process_image_with_ocr(
                img, keep_images, enable_ocr, output_folder,
                progress_callback, current_image_index, total_images, cancel_event
            )
            content_lines.append(image_link)
        
        # ==== 输出分节符 ====
        if pending_section_break:
            content_lines.append(pending_section_break)
            logger.debug(f"段落 {idx+1} 输出分节符: {pending_section_break}")
        
        return content_lines, current_image_index
    
    def finalize(self) -> List[str]:
        """
        处理器收尾，返回剩余内容
        
        处理以下情况：
        - 文档末尾未关闭的代码块
        - 文档末尾的边框组
        
        返回:
            List[str]: 剩余内容行列表
        
        注意:
            应在所有段落处理完成后调用
        """
        content_lines = []
        
        # 处理未关闭的代码块
        if self.ctx.in_code_block:
            code_block = self.ctx.end_code_block()
            if code_block:
                content_lines.append(code_block)
                logger.debug("文档末尾代码块结束")
        
        # 处理文档末尾边框组
        final_separator = self.border_tracker.finalize(self.config_manager)
        if final_separator:
            content_lines.append(final_separator)
            logger.debug(f"文档末尾边框组分隔线: {final_separator}")
        
        return content_lines
    
    # ========== 私有方法 ==========
    
    def _detect_section_break(self, para, idx: int) -> Optional[str]:
        """检测分节符"""
        if not self.config_manager.is_horizontal_rule_enabled():
            return None
        
        section_break_type, _ = detect_section_break_in_paragraph(para)
        if section_break_type:
            md_separator = self.config_manager.get_md_separator_for_break_type(section_break_type)
            if md_separator:
                logger.debug(f"段落 {idx+1} 包含分节符 {section_break_type}")
                return md_separator
        return None
    
    def _detect_page_break(self, para, idx: int) -> Tuple[bool, Optional[str]]:
        """检测分页符"""
        if not self.config_manager.is_horizontal_rule_enabled():
            return False, None
        
        for run in para.runs:
            if detect_page_break_in_run(run):
                separator = self.config_manager.get_md_separator_for_break_type('page_break')
                logger.debug(f"段落 {idx+1} 检测到分页符")
                return True, separator
        
        return False, None
    
    def _try_handle_formula(self, para, idx: int, current_image_index: int) -> Tuple[Optional[List[str]], int]:
        """尝试处理公式段落"""
        if not (is_formula_supported() and has_formulas_in_paragraph(para)):
            return None, current_image_index
        
        formula_md = process_paragraph_with_formulas(
            para, 
            preserve_formatting=self.preserve_formatting, 
            syntax_config=self.syntax_config
        )
        if not formula_md:
            return None, current_image_index
        
        content_lines = []
        
        # 检查是否为列表项
        list_type, list_level, num_id = detect_list_item(para)
        
        if list_type:
            # 列表项中的公式
            marker_type = self.config_manager.get_unordered_list_marker()
            marker = get_list_marker(marker_type)
            indent = " " * self.list_indent_spaces * list_level
            
            # 转为行内公式
            formula_content = self._convert_to_inline_formula(formula_md)
            
            if list_type == "bullet":
                list_line = f"{indent}{marker} {formula_content}"
            else:
                number = self.list_counter_manager.increment(num_id, list_level)
                self.list_counter_manager.reset_deeper_levels(num_id, list_level)
                list_line = f"{indent}{number}. {formula_content}"
            
            content_lines.append(list_line)
            logger.debug(f"段落 {idx+1} 列表项包含公式")
        else:
            # 非列表项公式
            formula_list_context = self.list_context_manager.get_context_for_para(idx, para)
            if formula_list_context:
                _, f_level = formula_list_context
                f_indent = " " * self.list_indent_spaces * (f_level + 1)
                indented_formula = "\n".join(f_indent + line for line in formula_md.split("\n"))
                content_lines.append(indented_formula)
            else:
                content_lines.append(formula_md)
            
            logger.debug(f"段落 {idx+1} 包含公式")
        
        return content_lines, current_image_index
    
    def _convert_to_inline_formula(self, formula_md: str) -> str:
        """将公式块转为行内公式"""
        formula_content = formula_md.strip()
        if formula_content.startswith('$$\n') and formula_content.endswith('\n$$'):
            inner = formula_content[3:-3].strip()
            return f"${inner}$"
        elif formula_content.startswith('$$') and formula_content.endswith('$$'):
            inner = formula_content[2:-2].strip()
            return f"${inner}$"
        return formula_content
    
    def _try_handle_style(self, para, idx: int) -> Optional[List[str]]:
        """尝试处理样式段落（代码块/引用块）"""
        style_type, style_value = detect_paragraph_style_type(para, self.config_manager)
        
        if style_type == 'code_block':
            if not self.ctx.in_code_block:
                list_context = self.list_context_manager.get_context_for_para(idx, para)
                self.ctx.start_code_block(list_context)
            self.ctx.add_code_line(para.text)
            logger.debug(f"段落 {idx+1}: 代码块样式")
            return []  # 返回空列表，代码块结束时统一输出
        
        elif style_type == 'quote':
            content_lines = []
            
            # 先结束代码块
            if self.ctx.in_code_block:
                code_block = self.ctx.end_code_block()
                if code_block:
                    content_lines.append(code_block)
            
            # 生成引用块
            quote_prefix = '>' * style_value + ' '
            quote_text, _, _ = convert_paragraph_to_markdown_with_styles(
                para, self.config_manager, self.preserve_formatting,
                self.syntax_config, self.wps_shading_enabled, self.word_shading_enabled
            )
            
            # 检查引用块是否在列表上下文中
            quote_list_context = self.list_context_manager.get_context_for_para(idx, para)
            if quote_list_context:
                _, q_level = quote_list_context
                q_indent = " " * self.list_indent_spaces * (q_level + 1)
                content_lines.append(f"{q_indent}{quote_prefix}{quote_text}")
            else:
                content_lines.append(f"{quote_prefix}{quote_text}")
            
            logger.debug(f"段落 {idx+1}: {style_value}级引用样式")
            return content_lines
        
        return None
    
    def _try_handle_shading(self, para, idx: int) -> Optional[List[str]]:
        """尝试处理底纹段落（兼容处理）"""
        is_code_block = has_paragraph_gray_shading(
            para, self.wps_shading_enabled, self.word_shading_enabled
        )
        
        if is_code_block:
            if not self.ctx.in_code_block:
                list_context = self.list_context_manager.get_context_for_para(idx, para)
                self.ctx.start_code_block(list_context)
            self.ctx.add_code_line(para.text)
            logger.debug(f"段落 {idx+1}: 灰色底纹（代码块兼容）")
            return []  # 代码块结束时统一输出
        
        return None
    
    def _get_formatted_text(self, para) -> Tuple[str, Optional[str], Optional[int]]:
        """获取段落的格式化文本"""
        para_text_raw = para.text.strip()
        
        if self.preserve_formatting:
            return convert_paragraph_to_markdown_with_styles(
                para, self.config_manager, self.preserve_formatting,
                self.syntax_config, self.wps_shading_enabled,
                self.word_shading_enabled, self.note_extractor
            )
        else:
            return para_text_raw, None, None
    
    def _handle_empty_paragraph(
        self, para, idx: int, images_info: list, keep_images: bool,
        enable_ocr: bool, output_folder: str, progress_callback: Callable,
        cancel_event, current_image_index: int, total_images: int,
        has_page_break: bool, page_break_separator: str,
        pending_section_break: str
    ) -> Tuple[List[str], int]:
        """处理空段落"""
        content_lines = []
        
        # 检查空段落中的图片
        para_images = get_paragraph_images(para, images_info)
        
        if para_images:
            for img in para_images:
                current_image_index += 1
                filename = img.get('filename', 'unknown')
                logger.info(f"在空段落 {idx + 1} 匹配到图片: {filename}")
                
                image_link = process_image_with_ocr(
                    img, keep_images, enable_ocr, output_folder,
                    progress_callback, current_image_index, total_images, cancel_event
                )
                content_lines.append(image_link)
        else:
            logger.debug(f"跳过空段落: {idx+1}")
        
        # 空段落也要处理分页符和分节符
        if has_page_break and page_break_separator:
            content_lines.append(page_break_separator)
        if pending_section_break:
            content_lines.append(pending_section_break)
        
        return content_lines, current_image_index
    
    def _handle_heading(
        self, para, para_text_raw: str, para_text_formatted: str,
        style_name: str, idx: int
    ) -> List[str]:
        """处理标题段落"""
        content_lines = []
        
        match = re.match(r'Heading (\d)', style_name)
        if match:
            level = int(match.group(1))
            final_text = self._apply_numbering(para, para_text_raw, level)
            
            markdown_line = '#' * level + ' ' + final_text
            content_lines.append(markdown_line)
            logger.debug(f"段落 {idx+1} Heading {level}")
        else:
            content_lines.append(para_text_formatted)
            logger.warning(f"段落 {idx+1} 无法解析标题级别: {style_name}")
        
        return content_lines
    
    def _handle_title(self, para, para_text_raw: str, idx: int) -> List[str]:
        """处理Title样式"""
        final_text = self._apply_numbering(para, para_text_raw, 1)
        markdown_line = '# ' + final_text
        logger.debug(f"段落 {idx+1} Title样式")
        return [markdown_line]
    
    def _handle_subtitle(self, para, para_text_raw: str, idx: int) -> List[str]:
        """处理Subtitle样式"""
        final_text = self._apply_numbering(para, para_text_raw, 2)
        markdown_line = '## ' + final_text
        logger.debug(f"段落 {idx+1} Subtitle样式")
        return [markdown_line]
    
    def _apply_numbering(self, para, para_text_raw: str, level: int) -> str:
        """应用序号配置（清除/添加）"""
        skip_chars = 0
        
        if self.remove_numbering:
            from docwen.utils.heading_utils import remove_heading_numbering
            cleaned_raw = remove_heading_numbering(para_text_raw)
            if cleaned_raw.strip() and cleaned_raw != para_text_raw:
                skip_chars = len(para_text_raw) - len(cleaned_raw)
        
        # 使用Run级别跳过序号
        final_text = convert_paragraph_to_markdown_skip_prefix(
            para, skip_chars, self.preserve_heading_formatting,
            self.syntax_config, self.wps_shading_enabled, self.word_shading_enabled
        )
        
        # 添加新序号
        if self.add_numbering and self.formatter:
            self.formatter.increment_level(level)
            numbering = self.formatter.format_heading(level)
            final_text = numbering + final_text
        
        return final_text
    
    def _handle_list_or_normal(
        self, para, para_text_formatted: str, idx: int,
        has_page_break: bool, page_break_separator: str
    ) -> List[str]:
        """处理列表或普通段落"""
        content_lines = []
        
        list_type, list_level, num_id = detect_list_item(para)
        
        if list_type:
            # 列表项处理
            marker_type = self.config_manager.get_unordered_list_marker()
            marker = get_list_marker(marker_type)
            indent = " " * self.list_indent_spaces * list_level
            
            if list_type == "bullet":
                list_line = f"{indent}{marker} {para_text_formatted}"
            else:
                number = self.list_counter_manager.increment(num_id, list_level)
                self.list_counter_manager.reset_deeper_levels(num_id, list_level)
                list_line = f"{indent}{number}. {para_text_formatted}"
            
            content_lines.append(list_line)
            logger.debug(f"段落 {idx+1} {list_type}列表项")
        else:
            # 检查列表续行
            list_context = self.list_context_manager.get_context_for_para(idx, para)
            
            if list_context:
                num_id, continuation_level = list_context
                indent = " " * self.list_indent_spaces * (continuation_level + 1)
                
                if has_page_break:
                    split_parts = self._split_by_page_breaks(para)
                    for part in split_parts:
                        content_lines.append(f"{indent}{part}")
                else:
                    content_lines.append(f"{indent}{para_text_formatted}")
                
                logger.debug(f"段落 {idx+1} 列表续行")
            else:
                # 普通段落
                if has_page_break:
                    split_parts = self._split_by_page_breaks(para)
                    content_lines.extend(split_parts)
                else:
                    content_lines.append(para_text_formatted)
                
                # 非列表段落重置计数器
                self.list_counter_manager.reset_all()
                logger.debug(f"段落 {idx+1} 普通段落")
        
        return content_lines
    
    def _split_by_page_breaks(self, para) -> List[str]:
        """按分页符分割段落"""
        
        results = []
        current_text_parts = []
        
        for run in para.runs:
            if detect_page_break_in_run(run):
                if current_text_parts:
                    text = ''.join(current_text_parts)
                    current_text_parts = []
                    if text.strip():
                        results.append(text)
                
                md_sep = self.config_manager.get_md_separator_for_break_type('page_break')
                if md_sep:
                    results.append(md_sep)
            
            if run.text:
                if self.preserve_formatting and self.syntax_config:
                    text = apply_format_markers(
                        run, run.text, self.syntax_config,
                        self.wps_shading_enabled, self.word_shading_enabled
                    )
                else:
                    text = run.text
                current_text_parts.append(text)
        
        if current_text_parts:
            text = ''.join(current_text_parts)
            if text.strip():
                results.append(text)
        
        return results
