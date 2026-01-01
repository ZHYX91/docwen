"""
边框组处理模块

负责边框组状态跟踪、分隔线输出和分页/分节符检测。

主要组件：
- extract_paragraph_border_info(): 提取段落边框信息
- detect_horizontal_rule_in_paragraph(): 检测分隔线
- detect_page_or_section_break(): 检测分页/分节符
- detect_page_break_in_run(): 检测 Run 中的分页符
- detect_section_break_in_paragraph(): 检测段落中的分节符
- detect_all_breaks_in_paragraph(): 检测所有分隔符
- BorderGroupTracker: 边框组状态跟踪器
- is_valid_border(): 边框有效性检查

边框组规则：
- 边框组第一个段落的顶部边框（top=single）→ 输出分隔线
- 边框组最后一个段落的底部边框（bottom=single）→ 输出分隔线
- 边框组内部 between=single → 段落间输出分隔线
- 两个相邻边框组边界合并（上一组 bottom + 下一组 top）→ 只输出1条
"""

import logging

logger = logging.getLogger(__name__)


# ======================================
#  边框信息提取函数
# ======================================

def extract_paragraph_border_info(para) -> dict:
    """
    提取段落的完整边框信息（包括样式继承的边框）
    
    用于边框组检测和可视分隔线计算。返回边框各方向的状态和 between 属性。
    
    注意：此函数会同时检查：
    1. 段落直接定义的边框（<w:pPr><w:pBdr>）
    2. 段落样式中继承的边框（样式定义中的 <w:pBdr>）
    
    参数:
        para: python-docx Paragraph 对象
    
    返回:
        dict: 边框信息字典
            - 'has_border': bool - 是否有有效边框（至少一个方向有实线边框，非none/nil）
            - 'top': str|None - 顶部边框类型 ('single', 'none', None)
            - 'bottom': str|None - 底部边框类型
            - 'left': str|None - 左侧边框类型
            - 'right': str|None - 右侧边框类型
            - 'between': str|None - 段落间边框类型（边框组内部分隔线）
            - 'is_box': bool - 是否是框（有左或右边框）
    """
    WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    result = {
        'has_border': False,
        'top': None,
        'bottom': None,
        'left': None,
        'right': None,
        'between': None,
        'is_box': False
    }
    
    def extract_border_from_pBdr(pBdr, target_result):
        """从 pBdr 元素提取边框信息到 target_result"""
        if pBdr is None:
            return False
        
        found_any = False
        for direction in ['top', 'bottom', 'left', 'right', 'between']:
            elem = pBdr.find(f'{WORD_NS}{direction}')
            if elem is not None:
                val = elem.get(f'{WORD_NS}val')
                if val and target_result[direction] is None:  # 不覆盖已有值
                    target_result[direction] = val
                    found_any = True
        return found_any
    
    try:
        para_xml = para._element
        pPr = para_xml.find(f'{WORD_NS}pPr')
        
        # 1. 首先检查段落直接定义的边框
        if pPr is not None:
            pBdr = pPr.find(f'{WORD_NS}pBdr')
            extract_border_from_pBdr(pBdr, result)
        
        # 2. 如果没有找到完整边框，尝试从样式中获取
        # 检查是否需要从样式中获取边框（任何方向都没有定义时）
        needs_style_check = not any(
            result[direction] is not None
            for direction in ['top', 'bottom', 'left', 'right', 'between']
        )
        
        if needs_style_check:
            # 尝试从段落样式获取边框定义
            try:
                style = para.style
                if style is not None:
                    # 通过样式的 _element 访问样式XML
                    style_xml = style._element
                    if style_xml is not None:
                        style_pPr = style_xml.find(f'{WORD_NS}pPr')
                        if style_pPr is not None:
                            style_pBdr = style_pPr.find(f'{WORD_NS}pBdr')
                            if extract_border_from_pBdr(style_pBdr, result):
                                logger.debug(f"从样式 '{style.name}' 中提取到边框信息")
            except Exception as e:
                logger.debug(f"从样式提取边框失败: {e}")
        
        # has_border 应检查是否有至少一个有效边框（非none/nil）
        has_valid_border = any(
            result[direction] and result[direction] not in ('none', 'nil')
            for direction in ['top', 'bottom', 'left', 'right', 'between']
        )
        result['has_border'] = has_valid_border
        
        # 判断是否是框（有左或右实线边框）
        if result['left'] and result['left'] not in ('none', 'nil'):
            result['is_box'] = True
        if result['right'] and result['right'] not in ('none', 'nil'):
            result['is_box'] = True
        
        return result
        
    except Exception as e:
        logger.debug(f"提取边框信息失败: {e}")
        return result


def detect_horizontal_rule_in_paragraph(para) -> bool:
    """
    检测段落是否包含分隔线（只有底部边框，没有其他边框）
    
    Word 文档中的分隔线通常是通过段落底部边框实现的。当用户在空白行输入 --- 或 ***
    并按回车时，Word 会自动将其转换为前一个段落的底部边框。
    
    检测条件：
    1. 段落属性中包含 <w:pBdr><w:bottom> 元素
    2. 没有顶部边框 <w:top>
    3. 没有左侧边框 <w:left>
    4. 没有右侧边框 <w:right>
    
    参数:
        para: python-docx Paragraph 对象
    
    返回:
        bool: 是否为分隔线段落
    """
    WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    try:
        para_xml = para._element
        pPr = para_xml.find(f'{WORD_NS}pPr')
        if pPr is None:
            return False
        
        pBdr = pPr.find(f'{WORD_NS}pBdr')
        if pBdr is None:
            return False
        
        has_bottom = pBdr.find(f'{WORD_NS}bottom') is not None
        has_top = pBdr.find(f'{WORD_NS}top') is not None
        has_left = pBdr.find(f'{WORD_NS}left') is not None
        has_right = pBdr.find(f'{WORD_NS}right') is not None
        
        if has_bottom and not has_top and not has_left and not has_right:
            logger.debug("检测到分隔线段落（只有底部边框）")
            return True
        
        return False
        
    except Exception as e:
        logger.debug(f"检测分隔线时出错: {e}")
        return False


# ======================================
#  分页符/分节符检测函数
# ======================================

def detect_page_or_section_break(para):
    """
    检测段落中是否包含分页符或分节符
    
    分页符: <w:br w:type="page"/>
    分节符: <w:pPr><w:sectPr><w:type w:val="..."/></w:sectPr></w:pPr>
    
    参数:
        para: python-docx Paragraph 对象
    
    返回:
        tuple: (break_type, break_value)
            - break_type: "page_break" | "section_next" | "section_continuous" | 
                         "section_even" | "section_odd" | None
            - break_value: 原始 Word 值（用于调试）
        
        如果不包含分页/分节符，返回 (None, None)
    """
    WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    try:
        para_xml = para._element
        
        # 1. 检测分页符 <w:br w:type="page"/>
        for br in para_xml.iter(f'{WORD_NS}br'):
            br_type = br.get(f'{WORD_NS}type')
            if br_type == 'page':
                return ('page_break', 'page')
        
        # 2. 检测分节符
        pPr = para_xml.find(f'{WORD_NS}pPr')
        if pPr is not None:
            sectPr = pPr.find(f'{WORD_NS}sectPr')
            if sectPr is not None:
                type_elem = sectPr.find(f'{WORD_NS}type')
                if type_elem is not None:
                    sect_val = type_elem.get(f'{WORD_NS}val', '')
                    
                    section_type_mapping = {
                        'nextPage': 'section_next',
                        'continuous': 'section_continuous',
                        'evenPage': 'section_even',
                        'oddPage': 'section_odd'
                    }
                    
                    break_type = section_type_mapping.get(sect_val)
                    if break_type:
                        return (break_type, sect_val)
                else:
                    return ('section_next', 'nextPage')
        
        return (None, None)
        
    except Exception as e:
        return (None, None)


def detect_page_break_in_run(run):
    """
    检测 Run 中是否包含分页符
    
    分页符存储在 Run 内部: <w:r><w:br w:type="page"/></w:r>
    
    参数:
        run: python-docx Run 对象
    
    返回:
        bool: 是否包含分页符
    """
    WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    try:
        run_xml = run._r
        
        for br in run_xml.iter(f'{WORD_NS}br'):
            br_type = br.get(f'{WORD_NS}type')
            if br_type == 'page':
                return True
        
        return False
        
    except Exception as e:
        logger.debug(f"检测 Run 中分页符时出错: {e}")
        return False


def detect_section_break_in_paragraph(para):
    """
    检测段落中是否包含分节符（只检测分节符，不检测分页符）
    
    参数:
        para: python-docx Paragraph 对象
    
    返回:
        tuple: (break_type, break_value)
            - break_type: "section_next" | "section_continuous" | 
                         "section_even" | "section_odd" | None
            - break_value: 原始 Word 值（用于调试）
    """
    WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    try:
        para_xml = para._element
        
        pPr = para_xml.find(f'{WORD_NS}pPr')
        if pPr is not None:
            sectPr = pPr.find(f'{WORD_NS}sectPr')
            if sectPr is not None:
                type_elem = sectPr.find(f'{WORD_NS}type')
                if type_elem is not None:
                    sect_val = type_elem.get(f'{WORD_NS}val', '')
                    
                    section_type_mapping = {
                        'nextPage': 'section_next',
                        'continuous': 'section_continuous',
                        'evenPage': 'section_even',
                        'oddPage': 'section_odd'
                    }
                    
                    break_type = section_type_mapping.get(sect_val)
                    if break_type:
                        return (break_type, sect_val)
                else:
                    return ('section_next', 'nextPage')
        
        return (None, None)
        
    except Exception as e:
        logger.debug(f"检测分节符时出错: {e}")
        return (None, None)


def detect_all_breaks_in_paragraph(para):
    """
    检测段落中的所有分页符和分节符
    
    参数:
        para: python-docx Paragraph 对象
    
    返回:
        list: 分隔符列表，每个元素为 (break_type, break_value) 元组
    """
    WORD_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    breaks = []
    
    try:
        para_xml = para._element
        
        # 1. 检测所有分页符
        for br in para_xml.iter(f'{WORD_NS}br'):
            br_type = br.get(f'{WORD_NS}type')
            if br_type == 'page':
                breaks.append(('page_break', 'page'))
        
        # 2. 检测分节符
        pPr = para_xml.find(f'{WORD_NS}pPr')
        if pPr is not None:
            sectPr = pPr.find(f'{WORD_NS}sectPr')
            if sectPr is not None:
                type_elem = sectPr.find(f'{WORD_NS}type')
                if type_elem is not None:
                    sect_val = type_elem.get(f'{WORD_NS}val', '')
                    
                    section_type_mapping = {
                        'nextPage': 'section_next',
                        'continuous': 'section_continuous',
                        'evenPage': 'section_even',
                        'oddPage': 'section_odd'
                    }
                    
                    break_type = section_type_mapping.get(sect_val)
                    if break_type:
                        breaks.append((break_type, sect_val))
                else:
                    breaks.append(('section_next', 'nextPage'))
        
        return breaks
        
    except Exception as e:
        logger.debug(f"检测分页/分节符时出错: {e}")
        return breaks


def is_valid_border(border_val) -> bool:
    """
    检查边框值是否为有效的实线
    
    Word 边框类型包括：
    - 'single': 单线（有效）
    - 'double': 双线（有效）
    - 'none': 无边框（无效）
    - 'nil': 无边框（无效）
    - None: 未设置（无效）
    
    参数:
        border_val: 边框值（如 'single', 'none', 'nil', None）
    
    返回:
        bool: 是否为有效的实线边框
    
    示例:
        >>> is_valid_border('single')
        True
        >>> is_valid_border('none')
        False
        >>> is_valid_border(None)
        False
    """
    return border_val and border_val not in ('none', 'nil')


class BorderGroupTracker:
    """
    边框组状态跟踪器
    
    用于处理 Word 段落边框 → Markdown 分隔线的转换。
    边框组：相邻的、有边框设置的段落视为一组（不含框，即没有左右边框）。
    
    边框组规则：
    1. 进入边框组：输出 top 边框对应的分隔线
    2. 边框组内部：输出 between 边框对应的分隔线
    3. 退出边框组：输出 bottom 边框对应的分隔线
    4. 相邻边框组边界合并：上一组 bottom + 下一组 top → 只输出1条
    
    使用方式：
        tracker = BorderGroupTracker()
        
        for para in paragraphs:
            # 处理段落前的分隔线
            separators = tracker.process_paragraph(para, config_manager)
            for sep in separators:
                output(sep)
            
            # 输出段落内容
            output(para_content)
        
        # 文档末尾处理
        final_sep = tracker.finalize(config_manager)
        if final_sep:
            output(final_sep)
    
    注意:
        - 框（有左或右边框的段落）不视为边框组成员
        - 需要配合 config_manager.is_horizontal_rule_enabled() 使用
    """
    
    def __init__(self):
        """初始化边框组跟踪器"""
        self.prev_border_info = None
        self.prev_had_border = False
    
    def reset(self):
        """
        重置状态
        
        当处理新文档时应调用此方法。
        """
        self.prev_border_info = None
        self.prev_had_border = False
        logger.debug("BorderGroupTracker 状态已重置")
    
    def process_paragraph(self, para, config_manager) -> list:
        """
        处理段落边框，返回需要在段落前输出的分隔线列表
        
        参数:
            para: Word段落对象
            config_manager: 配置管理器实例
        
        返回:
            list: 需要输出的分隔线列表（可能为空、1个或多个）
        
        注意:
            此方法会更新内部状态（prev_border_info, prev_had_border）。
        """
        separators = []
        
        if not config_manager.is_horizontal_rule_enabled():
            return separators
        
        # 提取当前段落边框信息
        curr_border_info = extract_paragraph_border_info(para)
        # 有边框且不是框（没有左右边框）才视为边框组
        curr_has_border = curr_border_info['has_border'] and not curr_border_info['is_box']
        
        md_separator = config_manager.get_md_separator_for_break_type('horizontal_rule')
        
        if curr_has_border and not self.prev_had_border:
            # 【进入边框组】：当前有边框，上一个没有 → 输出 top 边框
            if is_valid_border(curr_border_info['top']):
                if md_separator:
                    separators.append(md_separator)
                    logger.debug(f"进入边框组，输出 top 分隔线: {md_separator}")
        
        elif curr_has_border and self.prev_had_border:
            # 【边框组内部切换】：上一个和当前都有边框 → 检查 between
            if self.prev_border_info and is_valid_border(self.prev_border_info['between']):
                if md_separator:
                    separators.append(md_separator)
                    logger.debug(f"边框组内部，输出 between 分隔线: {md_separator}")
        
        elif not curr_has_border and self.prev_had_border:
            # 【退出边框组】：当前没边框，上一个有 → 输出上一个的 bottom 边框
            if self.prev_border_info and is_valid_border(self.prev_border_info['bottom']):
                if md_separator:
                    separators.append(md_separator)
                    logger.debug(f"退出边框组，输出上一段 bottom 分隔线: {md_separator}")
        
        # 更新状态
        self.prev_border_info = curr_border_info
        self.prev_had_border = curr_has_border
        
        return separators
    
    def finalize(self, config_manager) -> str:
        """
        文档末尾处理，返回最后的分隔线（如有）
        
        如果最后一个段落在边框组内，需要输出其 bottom 边框。
        
        参数:
            config_manager: 配置管理器实例
        
        返回:
            str: 分隔线字符串，如果不需要则返回 None
        
        注意:
            应在所有段落处理完成后调用此方法。
        """
        if not self.prev_had_border or not self.prev_border_info:
            return None
        
        if not config_manager.is_horizontal_rule_enabled():
            return None
        
        md_separator = config_manager.get_md_separator_for_break_type('horizontal_rule')
        
        if is_valid_border(self.prev_border_info['bottom']):
            if md_separator:
                logger.debug(f"文档末尾边框组，输出最后段落 bottom 分隔线: {md_separator}")
                return md_separator
        
        return None
    
    def update_state(self, para, config_manager):
        """
        仅更新状态，不返回分隔线
        
        用于空段落等特殊情况，需要更新边框组状态但不输出分隔线。
        
        参数:
            para: Word段落对象
            config_manager: 配置管理器实例
        """
        if not config_manager.is_horizontal_rule_enabled():
            return
        
        curr_border_info = extract_paragraph_border_info(para)
        curr_has_border = curr_border_info['has_border'] and not curr_border_info['is_box']
        
        self.prev_border_info = curr_border_info
        self.prev_had_border = curr_has_border
    
    def __str__(self):
        """调试用字符串表示"""
        return (
            f"BorderGroupTracker("
            f"prev_had_border={self.prev_had_border}, "
            f"prev_border_info={self.prev_border_info})"
        )
