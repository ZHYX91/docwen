"""
列表处理模块

负责列表预扫描、续行检测和编号计数器管理。
可被 simple/ 和 gongwen/ 转换器共享使用。

主要组件：
- detect_list_item(): 检测段落是否为列表项
- get_list_marker(): 获取无序列表标记字符
- preprocess_list_ranges(): 预扫描文档中的列表范围
- ListCounterManager: 列表编号计数器管理
- ListContextManager: 列表上下文检测（封装闭包函数）
"""

import logging
from docwen.utils.docx_utils import NAMESPACES

logger = logging.getLogger(__name__)


# ======================================
#  列表检测函数
# ======================================

def detect_list_item(para):
    """
    检测段落是否为Word列表项（项目符号/编号列表）
    
    参数:
        para: Word段落对象
        
    返回:
        tuple: (list_type, level, num_id)
            - list_type: 列表类型 ("bullet", "number", None)
            - level: 列表级别 (0-8, 0为第一级)
            - num_id: 编号ID（用于判断是否同一列表）
        
    说明:
        Word列表通过 numPr 元素标识:
        - numId: 列表编号ID
        - ilvl: 列表级别（indentation level）
        
        通过 numFmt 判断类型:
        - bullet: 无序列表（项目符号）
        - decimal, upperLetter, lowerLetter, 
          upperRoman, lowerRoman 等: 有序列表
    """
    try:
        # 获取段落属性
        pPr = para._element.pPr
        if pPr is None:
            return None, None, None
        
        # 查找 numPr 元素（列表属性）
        numPr = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr')
        if numPr is None:
            return None, None, None
        
        # 获取列表ID
        numId_elem = numPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId')
        if numId_elem is None:
            return None, None, None
        
        num_id = numId_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        if num_id is None or num_id == '0':
            return None, None, None
        
        # 获取列表级别
        ilvl_elem = numPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl')
        level = 0
        if ilvl_elem is not None:
            level_val = ilvl_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if level_val:
                try:
                    level = int(level_val)
                except ValueError:
                    level = 0
        
        # 判断列表类型
        list_type = _get_list_type_from_numbering(para, num_id, level)
        
        logger.debug(f"检测到列表项: type={list_type}, level={level}, numId={num_id}")
        return list_type, level, num_id
        
    except Exception as e:
        logger.debug(f"检测列表项失败: {e}")
        return None, None, None


def _get_list_type_from_numbering(para, num_id, level):
    """
    从numbering.xml获取列表类型
    
    参数:
        para: 段落对象
        num_id: 编号ID
        level: 列表级别
        
    返回:
        str: "bullet" 或 "number"
    """
    try:
        doc = para.part.document
        numbering_part = doc.part.numbering_part
        
        if numbering_part is None:
            return "bullet"
        
        numbering_xml = numbering_part._element
        
        # 查找对应的 num 元素
        num_elem = numbering_xml.find(
            f'.//{{{NAMESPACES["w"]}}}num[@{{{NAMESPACES["w"]}}}numId="{num_id}"]'
        )
        
        if num_elem is None:
            for num in numbering_xml.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num'):
                if num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId') == num_id:
                    num_elem = num
                    break
        
        if num_elem is None:
            return "bullet"
        
        # 获取 abstractNumId
        abstract_num_id_elem = num_elem.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId')
        if abstract_num_id_elem is None:
            return "bullet"
        
        abstract_num_id = abstract_num_id_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        
        # 查找对应的 abstractNum 元素
        abstract_num = None
        for an in numbering_xml.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNum'):
            if an.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId') == abstract_num_id:
                abstract_num = an
                break
        
        if abstract_num is None:
            return "bullet"
        
        # 查找对应级别的 lvl 元素
        for lvl in abstract_num.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvl'):
            lvl_ilvl = lvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl')
            if lvl_ilvl == str(level):
                num_fmt = lvl.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numFmt')
                if num_fmt is not None:
                    fmt_val = num_fmt.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    if fmt_val == 'bullet':
                        return "bullet"
                    else:
                        return "number"
                break
        
        return "number"
        
    except Exception as e:
        logger.debug(f"获取列表类型失败: {e}")
        return "bullet"


def get_list_marker(marker_type="dash"):
    """
    获取无序列表标记字符
    
    参数:
        marker_type: 标记类型 ("dash", "asterisk", "plus")
        
    返回:
        str: 列表标记字符
    """
    markers = {
        "dash": "-",
        "asterisk": "*",
        "plus": "+"
    }
    return markers.get(marker_type, "-")


# 列表续行检测相关样式名（英文和中文）
# 同时支持 Word 内置的 List Paragraph 和自定义的 List Block
LIST_PARAGRAPH_STYLE_NAMES = [
    'List Paragraph', '列表段落', 'ListParagraph',
    'List Block', '列表块', 'ListBlock'
]


def preprocess_list_ranges(doc, skip_indices: list) -> dict:
    """
    预扫描文档中的列表范围
    
    为每个 numId 记录列表的 [first_index, last_index] 范围，
    用于后续识别列表续行内容。
    
    参数:
        doc: Document对象
        skip_indices: 需要跳过的段落索引列表（如已提取到YAML的标题）
    
    返回:
        dict: {numId: {'first': int, 'last': int, 'levels': {idx: level}}}
            - numId: 列表编号ID
            - first: 列表第一个段落索引
            - last: 列表最后一个段落索引
            - levels: {段落索引: 级别} 映射
    
    注意:
        - 必须在主循环之前调用
        - 返回结果用于 ListContextManager 判断续行
    
    示例:
        >>> list_ranges = preprocess_list_ranges(doc, [0, 1])
        >>> print(list_ranges)
        {'1': {'first': 5, 'last': 10, 'levels': {5: 0, 6: 1, 7: 0, ...}}}
    """
    list_ranges = {}
    
    for idx, para in enumerate(doc.paragraphs):
        if idx in skip_indices:
            continue
        
        list_type, list_level, num_id = detect_list_item(para)
        
        if list_type and num_id:
            if num_id not in list_ranges:
                list_ranges[num_id] = {
                    'first': idx,
                    'last': idx,
                    'levels': {}
                }
            else:
                list_ranges[num_id]['last'] = idx
            
            list_ranges[num_id]['levels'][idx] = list_level
    
    logger.debug(f"列表预扫描完成: 发现 {len(list_ranges)} 个独立列表")
    return list_ranges


class ListCounterManager:
    """
    列表编号计数器管理
    
    管理有序列表的编号计数，支持多级嵌套列表。
    每个 numId 独立计数，每个级别独立计数。
    
    使用方式：
        manager = ListCounterManager()
        
        # 有序列表项
        number = manager.increment(num_id, level)
        markdown_line = f"{indent}{number}. {text}"
        
        # 返回更浅级别时，自动重置更深级别
        manager.reset_deeper_levels(num_id, level)
        
        # 遇到非列表内容，重置所有计数器
        manager.reset_all()
    """
    
    def __init__(self):
        """初始化列表计数器管理器"""
        self.counters = {}  # {num_id: {level: counter}}
    
    def increment(self, num_id: str, level: int) -> int:
        """
        递增计数器并返回当前编号
        
        参数:
            num_id: 列表编号ID（从 Word 文档中提取）
            level: 列表级别（0为第一级）
        
        返回:
            int: 当前编号值（从1开始）
        
        示例:
            >>> manager = ListCounterManager()
            >>> manager.increment('1', 0)  # 第一个一级列表项
            1
            >>> manager.increment('1', 0)  # 第二个一级列表项
            2
            >>> manager.increment('1', 1)  # 第一个二级列表项
            1
        """
        if num_id not in self.counters:
            self.counters[num_id] = {}
        
        if level not in self.counters[num_id]:
            self.counters[num_id][level] = 0
        
        self.counters[num_id][level] += 1
        return self.counters[num_id][level]
    
    def reset_deeper_levels(self, num_id: str, level: int):
        """
        重置更深级别的计数器
        
        当返回更浅级别时，需要重置更深级别的计数器。
        
        参数:
            num_id: 列表编号ID
            level: 当前级别
        
        示例:
            # 假设当前在 level=2，现在回到 level=0
            >>> manager.reset_deeper_levels('1', 0)
            # 会重置 level=1, level=2 等所有更深级别的计数器
        """
        if num_id not in self.counters:
            return
        
        for lvl in list(self.counters[num_id].keys()):
            if lvl > level:
                del self.counters[num_id][lvl]
    
    def reset_all(self):
        """
        重置所有计数器
        
        遇到非列表内容时调用，因为根据 Markdown 语法规范，
        被非列表内容中断后的列表项属于新列表。
        """
        self.counters = {}
        logger.debug("列表计数器已全部重置")
    
    def __str__(self):
        """调试用字符串表示"""
        return f"ListCounterManager(counters={self.counters})"


class ListContextManager:
    """
    列表上下文管理器
    
    封装列表续行检测逻辑，替代原来的闭包函数。
    用于判断段落或位置是否在某个列表的范围内。
    
    使用方式：
        list_ranges = preprocess_list_ranges(doc, skip_indices)
        ctx_mgr = ListContextManager(list_ranges)
        
        # 检查表格是否在列表中
        table_context = ctx_mgr.get_context_for_position(table_position)
        if table_context:
            num_id, level = table_context
            # 为表格添加列表缩进
        
        # 检查段落是否为列表续行
        para_context = ctx_mgr.get_context_for_para(para_idx, para)
        if para_context:
            num_id, level = para_context
            # 为段落添加列表缩进
    """
    
    def __init__(self, list_ranges: dict, list_style_names: list = None):
        """
        初始化列表上下文管理器
        
        参数:
            list_ranges: 预扫描的列表范围字典（由 preprocess_list_ranges 返回）
            list_style_names: 列表样式名列表（默认使用模块常量 LIST_PARAGRAPH_STYLE_NAMES）
        """
        self.list_ranges = list_ranges
        self.list_style_names = list_style_names or LIST_PARAGRAPH_STYLE_NAMES
    
    def get_context_for_position(self, pos_idx: int) -> tuple:
        """
        检查给定位置是否在某个列表的范围内（用于表格等特殊元素）
        
        此方法不需要段落对象，适用于检查表格、图片等元素的位置。
        
        参数:
            pos_idx: 位置索引
        
        返回:
            tuple: (numId, level) 如果在列表范围内
            None: 如果不在列表范围内
        
        示例:
            >>> ctx = ctx_mgr.get_context_for_position(7)
            >>> if ctx:
            ...     num_id, level = ctx
            ...     indent = " " * 4 * (level + 1)
        """
        for num_id, range_info in self.list_ranges.items():
            first = range_info['first']
            last = range_info['last']
            
            # 在列表范围内（不包括边界，边界是列表项本身）
            if first < pos_idx < last:
                # 找到最近的前一个列表项的级别
                prev_level = self._find_previous_level(range_info, pos_idx)
                return (num_id, prev_level)
        
        return None
    
    def get_context_for_para(self, para_idx: int, para) -> tuple:
        """
        检查段落是否在某个列表的范围内（用于续行检测）
        
        此方法需要段落对象，用于检查段落样式（List Paragraph 等）。
        
        参数:
            para_idx: 段落索引
            para: 段落对象
        
        返回:
            tuple: (numId, level) 如果在列表范围内或是列表样式段落
            None: 如果不在列表范围内
        
        注意:
            - 段落在列表范围内（first < para_idx < last）
            - 或者段落使用 List Paragraph 样式且紧跟在列表最后一项之后
        """
        # 检查是否在列表范围内
        for num_id, range_info in self.list_ranges.items():
            first = range_info['first']
            last = range_info['last']
            
            if first < para_idx < last:
                # 找到最近的前一个列表项的级别
                prev_level = self._find_previous_level(range_info, para_idx)
                return (num_id, prev_level)
        
        # 检查是否是 List Paragraph 样式（即使超出范围也视为续行）
        if para.style and para.style.name in self.list_style_names:
            # 找最近的列表范围
            for num_id, range_info in self.list_ranges.items():
                last = range_info['last']
                if para_idx == last + 1:  # 紧跟在列表最后一项之后
                    last_level = range_info['levels'].get(last, 0)
                    logger.debug(f"段落 {para_idx} 使用 List Paragraph 样式，视为列表续行")
                    return (num_id, last_level)
        
        return None
    
    def _find_previous_level(self, range_info: dict, pos_idx: int) -> int:
        """
        找到最近的前一个列表项的级别
        
        参数:
            range_info: 列表范围信息字典
            pos_idx: 当前位置索引
        
        返回:
            int: 前一个列表项的级别，如果找不到则返回0
        """
        prev_level = 0
        for list_idx in sorted(range_info['levels'].keys(), reverse=True):
            if list_idx < pos_idx:
                prev_level = range_info['levels'][list_idx]
                break
        return prev_level
    
    def __str__(self):
        """调试用字符串表示"""
        return f"ListContextManager(ranges={len(self.list_ranges)} lists)"
