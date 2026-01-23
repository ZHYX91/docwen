"""
列表格式管理器模块

负责 MD→DOCX 转换中的列表处理：
- 扫描模板构建格式表
- 按需拼接列表定义
- 创建 Word 原生列表

核心策略：格式合并 + 按需拼接
- 从模板收集有序和无序格式（跳过 tentative 级别）
- 缺失级别用预设填充
- 根据 MD 列表结构拼接新定义
"""

import logging
from typing import Dict, List, Optional, Tuple
from lxml import etree
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .numbering_handler import (
    get_numbering_xml_root,
    get_max_abstract_num_id,
    get_max_num_id,
    WORD_NS
)

logger = logging.getLogger(__name__)

# ============================================
# 预设定义
# ============================================

# 有序列表预设（所有级别统一使用 decimal）
ORDERED_PRESET = {
    'start': '1',
    'numFmt': 'decimal',
    'lvlText': '%{level}.',  # 占位符，运行时替换
    'lvlJc': 'left',
}

# 无序列表预设（所有级别统一使用 bullet + Unicode 字符）
# 使用 Unicode 字符，无需特殊字体
UNORDERED_PRESET = {
    'start': '1',
    'numFmt': 'bullet',
    'lvlText': '•',  # Unicode U+2022 实心圆
    'lvlJc': 'left',
}

# 缩进配置（单位：twips，1英寸=1440 twips）
BASE_INDENT = 420       # 基础缩进
INDENT_INCREMENT = 420  # 每级递增


class ListFormatManager:
    """
    列表格式管理器
    
    负责扫描模板构建格式表，按需拼接列表定义。
    """
    
    def __init__(self, doc):
        """
        初始化列表格式管理器
        
        参数:
            doc: python-docx Document 对象
        """
        self.doc = doc
        self._numbering_root = None
        self._ordered_formats = None    # 有序格式表 {level: format_dict}
        self._unordered_formats = None  # 无序格式表 {level: format_dict}
        self._next_abstract_num_id = None
        self._next_num_id = None
        
        # 获取 numbering.xml 根元素
        self._numbering_root = get_numbering_xml_root(doc)
        if self._numbering_root is not None:
            self._next_abstract_num_id = get_max_abstract_num_id(self._numbering_root) + 1
            self._next_num_id = get_max_num_id(self._numbering_root) + 1
        else:
            self._next_abstract_num_id = 0
            self._next_num_id = 1
        
        logger.debug(f"ListFormatManager 初始化: next_abstractNumId={self._next_abstract_num_id}, next_numId={self._next_num_id}")
    
    def build_format_tables(self):
        """
        扫描模板，构建有序和无序格式表
        
        遍历顺序：abstractNumId 从大到小（最新的优先，ID 越大代表定义越新）
        收集规则：
            - 跳过 tentative="1" 的级别（占位符级别）
            - 智能检测格式属性，不盲目信任模板中的归类
            - 已有的有效级别不被覆盖（保证最新优先）
        """
        self._ordered_formats = {}
        self._unordered_formats = {}
        
        if self._numbering_root is None:
            logger.info("无 numbering.xml，使用预设填充所有级别")
            self._fill_missing_with_presets()
            return
        
        # 1. 收集所有 abstractNum，按 ID 降序排列（最新优先）
        abstract_nums = []
        for abstract_num in self._numbering_root.findall('.//{%s}abstractNum' % WORD_NS):
            abstract_num_id = abstract_num.get('{%s}abstractNumId' % WORD_NS)
            if abstract_num_id:
                try:
                    abstract_nums.append((int(abstract_num_id), abstract_num))
                except ValueError:
                    continue
        
        # 按 abstractNumId 从大到小排序
        abstract_nums.sort(key=lambda x: x[0], reverse=True)
        logger.debug(f"找到 {len(abstract_nums)} 个 abstractNum 定义，将按 ID 逆序扫描")
        
        # 2. 遍历收集格式
        for abstract_num_id, abstract_num in abstract_nums:
            for lvl in abstract_num.findall('.//{%s}lvl' % WORD_NS):
                # 获取级别
                ilvl = lvl.get('{%s}ilvl' % WORD_NS)
                if ilvl is None:
                    continue
                
                try:
                    level = int(ilvl)
                    if level < 0 or level > 8:
                        continue
                except ValueError:
                    continue
                
                # 关键过滤：跳过 tentative 级别
                # Word 会生成一些带 tentative="1" 的级别作为占位符，这些级别通常不是用户定义的格式
                tentative = lvl.get('{%s}tentative' % WORD_NS)
                if tentative == '1':
                    logger.debug(f"忽略占位级别 (tentative=1): id={abstract_num_id}, level={level}")
                    continue
                
                # 提取格式定义
                level_def = self._extract_level_definition(lvl)
                num_fmt = level_def.get('numFmt', 'decimal')
                lvl_text = level_def.get('lvlText', '')
                
                # 智能分类逻辑：不再仅凭 num_fmt
                # 1. 是否具备"有序"特征：lvlText 包含动态占位符（如 %1, %2）或者 numFmt 不是 bullet
                is_ordered_feature = ('%' in lvl_text) or (num_fmt != 'bullet')
                
                # 2. 是否具备"无序"特征：numFmt 是 bullet，或者 (没有动态占位符且 numFmt 不具备增量意义)
                # 注：大部分情况下 bullet 已经是决定性特征
                is_unordered_feature = (num_fmt == 'bullet') or (not ('%' in lvl_text))

                # 按表现出来的特征进行收集，由于是倒序遍历，第一个遇到的有效定义即为该级别的"最新真理"
                if is_ordered_feature and level not in self._ordered_formats:
                    self._ordered_formats[level] = level_def
                    logger.debug(f"收集最新有效有序格式: level={level}, from abstractNumId={abstract_num_id}, fmt={num_fmt}")
                
                if is_unordered_feature and level not in self._unordered_formats:
                    # 对于无序格式，我们需要更谨慎，确保它看起来真的像是一个符号列表
                    # 如果一个级别既有 % 又被标记为 bullet（极罕见），我们优先把它当有序占位
                    self._unordered_formats[level] = level_def
                    logger.debug(f"收集最新有效无序格式: level={level}, from abstractNumId={abstract_num_id}")
        
        # 3. 对收集到的表进行二次验证：如果模板中某个级别被"误标"，比如有序表里混入了没有占位符的固定字符，
        # 在后面拼接时会有保护，这里主要是填充缺失
        self._fill_missing_with_presets()
        
        logger.info(f"格式表构建完成: 有序={len(self._ordered_formats)}级, 无序={len(self._unordered_formats)}级")
    
    def _extract_level_definition(self, lvl_elem) -> dict:
        """
        从 lvl 元素提取格式定义
        
        参数:
            lvl_elem: w:lvl 元素
            
        返回:
            dict: 格式定义字典
        """
        level_def = {}
        
        # start
        start_elem = lvl_elem.find('.//{%s}start' % WORD_NS)
        if start_elem is not None:
            level_def['start'] = start_elem.get('{%s}val' % WORD_NS, '1')
        
        # numFmt
        num_fmt_elem = lvl_elem.find('.//{%s}numFmt' % WORD_NS)
        if num_fmt_elem is not None:
            level_def['numFmt'] = num_fmt_elem.get('{%s}val' % WORD_NS, 'decimal')
        
        # lvlText
        lvl_text_elem = lvl_elem.find('.//{%s}lvlText' % WORD_NS)
        if lvl_text_elem is not None:
            level_def['lvlText'] = lvl_text_elem.get('{%s}val' % WORD_NS, '')
        
        # lvlJc
        lvl_jc_elem = lvl_elem.find('.//{%s}lvlJc' % WORD_NS)
        if lvl_jc_elem is not None:
            level_def['lvlJc'] = lvl_jc_elem.get('{%s}val' % WORD_NS, 'left')
        
        # pPr/ind（缩进）
        pPr = lvl_elem.find('.//{%s}pPr' % WORD_NS)
        if pPr is not None:
            ind = pPr.find('.//{%s}ind' % WORD_NS)
            if ind is not None:
                level_def['ind_left'] = ind.get('{%s}left' % WORD_NS)
                level_def['ind_hanging'] = ind.get('{%s}hanging' % WORD_NS)
        
        # rPr（字符格式，用于 bullet 的字体）
        rPr = lvl_elem.find('.//{%s}rPr' % WORD_NS)
        if rPr is not None:
            rFonts = rPr.find('.//{%s}rFonts' % WORD_NS)
            if rFonts is not None:
                level_def['rFonts_ascii'] = rFonts.get('{%s}ascii' % WORD_NS)
                level_def['rFonts_hAnsi'] = rFonts.get('{%s}hAnsi' % WORD_NS)
        
        return level_def
    
    def _fill_missing_with_presets(self):
        """用预设填充缺失的级别"""
        for level in range(9):
            # 有序
            if level not in self._ordered_formats:
                preset = ORDERED_PRESET.copy()
                preset['lvlText'] = f'%{level + 1}.'
                preset['ind_left'] = str(BASE_INDENT + INDENT_INCREMENT * level)
                preset['ind_hanging'] = str(INDENT_INCREMENT)
                self._ordered_formats[level] = preset
                logger.debug(f"预设填充有序格式: level={level}")
            
            # 无序
            if level not in self._unordered_formats:
                preset = UNORDERED_PRESET.copy()
                preset['ind_left'] = str(BASE_INDENT + INDENT_INCREMENT * level)
                preset['ind_hanging'] = str(INDENT_INCREMENT)
                self._unordered_formats[level] = preset
                logger.debug(f"预设填充无序格式: level={level}")
    
    def create_list_definition(self, level_types: Dict[int, str]) -> str:
        """
        根据需求拼接并创建列表定义
        
        参数:
            level_types: {0: 'ordered', 1: 'unordered', 2: 'unordered', ...}
            
        返回:
            str: numId（用于段落引用）
        """
        # 确保格式表已构建
        if self._ordered_formats is None:
            self.build_format_tables()
        
        if self._numbering_root is None:
            logger.error("无法创建列表定义：numbering.xml 不存在")
            return None
        
        # 拼接各级别格式
        levels = []
        for level in range(9):
            if level in level_types:
                type_ = level_types[level]
            else:
                type_ = 'ordered'  # 默认有序
            
            if type_ == 'unordered':
                levels.append(self._unordered_formats[level])
            else:
                levels.append(self._ordered_formats[level])
        
        # 创建 abstractNum
        abstract_num_id = self._next_abstract_num_id
        self._next_abstract_num_id += 1
        
        abstract_num_elem = self._create_abstract_num(abstract_num_id, levels)
        
        # 关键修复：abstractNum 必须插入到所有 num 元素之前
        # 根据 OOXML 规范，numbering.xml 的正确结构是：
        # 1. 所有 <w:abstractNum> 元素
        # 2. 所有 <w:num> 元素
        first_num = self._numbering_root.find('{%s}num' % WORD_NS)
        if first_num is not None:
            # 找到第一个 num 元素，将 abstractNum 插入到它之前
            index = list(self._numbering_root).index(first_num)
            self._numbering_root.insert(index, abstract_num_elem)
        else:
            # 没有 num 元素，直接追加
            self._numbering_root.append(abstract_num_elem)
        logger.debug(f"创建 abstractNum: abstractNumId={abstract_num_id}")
        
        # 创建 num（追加到末尾是正确的）
        num_id = self._next_num_id
        self._next_num_id += 1
        
        num_elem = self._create_num(num_id, abstract_num_id)
        self._numbering_root.append(num_elem)
        logger.debug(f"创建 num: numId={num_id} -> abstractNumId={abstract_num_id}")
        
        return str(num_id)
    
    def _create_abstract_num(self, abstract_num_id: int, levels: List[dict]):
        """
        创建 abstractNum 元素
        
        参数:
            abstract_num_id: abstractNumId 值
            levels: 9 个级别的格式列表
            
        返回:
            lxml.etree.Element: abstractNum 元素
        """
        abstract_num = etree.Element('{%s}abstractNum' % WORD_NS)
        abstract_num.set('{%s}abstractNumId' % WORD_NS, str(abstract_num_id))
        
        # nsid（必需，随机值）
        nsid = etree.SubElement(abstract_num, '{%s}nsid' % WORD_NS)
        nsid.set('{%s}val' % WORD_NS, format(hash(str(abstract_num_id)) & 0xFFFFFFFF, '08X'))
        
        # multiLevelType
        multi_level = etree.SubElement(abstract_num, '{%s}multiLevelType' % WORD_NS)
        multi_level.set('{%s}val' % WORD_NS, 'hybridMultilevel')
        
        # 添加 9 个级别
        for level, level_def in enumerate(levels):
            lvl = self._create_lvl(level, level_def)
            abstract_num.append(lvl)
        
        return abstract_num
    
    def _create_lvl(self, level: int, level_def: dict):
        """
        创建 lvl 元素
        
        参数:
            level: 级别 (0-8)
            level_def: 格式定义字典
            
        返回:
            lxml.etree.Element: lvl 元素
        """
        lvl = etree.Element('{%s}lvl' % WORD_NS)
        lvl.set('{%s}ilvl' % WORD_NS, str(level))
        
        # start
        start = etree.SubElement(lvl, '{%s}start' % WORD_NS)
        start.set('{%s}val' % WORD_NS, level_def.get('start', '1'))
        
        # numFmt
        num_fmt = etree.SubElement(lvl, '{%s}numFmt' % WORD_NS)
        num_fmt.set('{%s}val' % WORD_NS, level_def.get('numFmt', 'decimal'))
        
        # lvlText
        lvl_text = etree.SubElement(lvl, '{%s}lvlText' % WORD_NS)
        lvl_text.set('{%s}val' % WORD_NS, level_def.get('lvlText', f'%{level + 1}.'))
        
        # lvlJc
        lvl_jc = etree.SubElement(lvl, '{%s}lvlJc' % WORD_NS)
        lvl_jc.set('{%s}val' % WORD_NS, level_def.get('lvlJc', 'left'))
        
        # pPr（段落属性，包含缩进）
        pPr = etree.SubElement(lvl, '{%s}pPr' % WORD_NS)
        ind = etree.SubElement(pPr, '{%s}ind' % WORD_NS)
        ind.set('{%s}left' % WORD_NS, level_def.get('ind_left', str(BASE_INDENT + INDENT_INCREMENT * level)))
        ind.set('{%s}hanging' % WORD_NS, level_def.get('ind_hanging', str(INDENT_INCREMENT)))
        
        # rPr（字符属性，仅用于 bullet）
        if level_def.get('numFmt') == 'bullet':
            rPr = etree.SubElement(lvl, '{%s}rPr' % WORD_NS)
            rFonts = etree.SubElement(rPr, '{%s}rFonts' % WORD_NS)
            # 如果有字体设置则使用，否则不设置（使用 Unicode 字符）
            if level_def.get('rFonts_ascii'):
                rFonts.set('{%s}ascii' % WORD_NS, level_def['rFonts_ascii'])
                rFonts.set('{%s}hAnsi' % WORD_NS, level_def.get('rFonts_hAnsi', level_def['rFonts_ascii']))
                rFonts.set('{%s}hint' % WORD_NS, 'default')
        
        return lvl
    
    def _create_num(self, num_id: int, abstract_num_id: int):
        """
        创建 num 元素
        
        参数:
            num_id: numId 值
            abstract_num_id: 引用的 abstractNumId
            
        返回:
            lxml.etree.Element: num 元素
        """
        num = etree.Element('{%s}num' % WORD_NS)
        num.set('{%s}numId' % WORD_NS, str(num_id))
        
        abstract_num_id_elem = etree.SubElement(num, '{%s}abstractNumId' % WORD_NS)
        abstract_num_id_elem.set('{%s}val' % WORD_NS, str(abstract_num_id))
        
        return num


def apply_list_to_paragraph(paragraph, num_id: str, level: int):
    """
    将列表属性应用到段落
    
    参数:
        paragraph: python-docx Paragraph 对象
        num_id: 列表编号 ID
        level: 列表级别 (0-8)
    """
    # 获取或创建 pPr
    pPr = paragraph._element.get_or_add_pPr()
    
    # 移除现有的 numPr（如果有）
    existing_numPr = pPr.find('.//{%s}numPr' % WORD_NS)
    if existing_numPr is not None:
        pPr.remove(existing_numPr)
    
    # 创建 numPr
    numPr = OxmlElement('w:numPr')
    
    # ilvl
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), str(level))
    numPr.append(ilvl)
    
    # numId
    numId_elem = OxmlElement('w:numId')
    numId_elem.set(qn('w:val'), num_id)
    numPr.append(numId_elem)
    
    # 将 numPr 插入到 pPr 的开头
    pPr.insert(0, numPr)
    
    logger.debug(f"应用列表属性: numId={num_id}, level={level}")


def analyze_list_structure(list_items: List[dict]) -> Dict[int, str]:
    """
    分析列表结构，确定每级类型
    
    冲突处理：同级别有序和无序混合时，统一为有序
    
    参数:
        list_items: 列表项数据列表 [{'level': 0, 'list_type': 'ordered'}, ...]
        
    返回:
        dict: {level: 'ordered'/'unordered'}
    """
    level_types = {}
    conflicts = []
    
    for item in list_items:
        level = item.get('level', 0)
        type_ = item.get('list_type', 'ordered')
        
        if level in level_types:
            if level_types[level] != type_:
                # 检测到冲突
                conflicts.append(level)
                logger.warning(
                    f"列表级别 {level} 类型冲突: "
                    f"已有 {level_types[level]}, 新遇 {type_}, "
                    f"统一为 ordered"
                )
                level_types[level] = 'ordered'
        else:
            level_types[level] = type_
    
    if conflicts:
        logger.info(f"列表结构分析完成，冲突级别: {conflicts}，已统一为 ordered")
    
    return level_types


def group_consecutive_list_items(body_data: List[dict]) -> List[List[dict]]:
    """
    将连续的列表项和相关内容分组
    
    每个分组是一个独立的列表，需要创建独立的 numId。
    
    不会打断列表的元素：
    - list_item: 列表项
    - list_continuation: 列表续行文本
    - 带有 in_list_context 标记的元素（如代码块、表格、引用块）
    
    会打断列表的元素：
    - heading: 标题（开始新章节）
    - horizontal_rule: 分隔符
    - 普通 content（无缩进）
    - 其他不在列表上下文中的元素
    
    参数:
        body_data: 段落数据列表
        
    返回:
        list: [group1, group2, ...] 每个 group 是列表项和相关内容的列表
    """
    groups = []
    current_group = []
    
    for item in body_data:
        item_type = item.get('type')
        
        if item_type == 'list_item':
            current_group.append(item)
        elif item_type == 'list_continuation':
            # 续行内容不打断列表，加入当前分组
            if current_group:
                current_group.append(item)
            else:
                # 如果续行出现在列表开始之前，视为普通内容（不应该发生，但保护性处理）
                logger.warning("list_continuation 出现在列表项之前，忽略")
        elif item.get('in_list_context'):
            # 带有列表上下文标记的元素（如缩进的代码块、表格、引用块）
            # 不打断列表，加入当前分组
            if current_group:
                current_group.append(item)
                logger.debug(f"列表内 {item_type} 加入当前分组")
            else:
                # 异常情况：in_list_context 出现在列表开始之前
                logger.warning(f"{item_type} 带有 in_list_context 但出现在列表项之前，忽略")
        else:
            # 其他类型的段落会终止列表分组
            if current_group:
                groups.append(current_group)
                current_group = []
    
    if current_group:
        groups.append(current_group)
    
    logger.info(f"列表分组: 共 {len(groups)} 个独立列表")
    return groups
