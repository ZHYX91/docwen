"""
DOCX处理工具模块
包含所有与DOCX文件处理相关的工具函数
"""

import logging
import copy
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import lxml.etree as etree
from docx.shared import Pt
from gongwen_converter.utils.validation_utils import contains_chinese

# 配置日志
logger = logging.getLogger(__name__)

# 字号映射表（半磅值 -> 中文名称）
FONT_SIZE_MAP = {
    10: "八号",    # 5磅
    11: "七号",    # 5.5磅
    13: "小六",    # 6.5磅
    15: "六号",    # 7.5磅
    18: "小五",    # 9磅
    21: "五号",    # 10.5磅
    24: "小四",    # 12磅
    28: "四号",    # 14磅
    30: "小三",    # 15磅
    32: "三号",    # 16磅
    36: "小二",    # 18磅
    44: "二号",    # 22磅
    48: "小一",    # 24磅
    52: "一号",    # 26磅
    72: "小初",    # 36磅
    84: "初号",    # 42磅
    108: "特号",   # 54磅
    126: "大特号"  # 63磅
}

def get_font_size_name(sz_val):
    """
    获取字号的中文名称（使用国家标准GB/T 148-1997）
    :param sz_val: 字号值（半磅值）
    :return: 字号中文名称
    """
    if sz_val is None:
        return ""
    
    try:
        # 尝试转换为整数
        sz_val = int(float(sz_val))
        
        # 精确匹配
        if sz_val in FONT_SIZE_MAP:
            return FONT_SIZE_MAP[sz_val]
        
        # 查找最接近的映射
        differences = {k: abs(k - sz_val) for k in FONT_SIZE_MAP}
        closest = min(differences, key=differences.get)
        return FONT_SIZE_MAP[closest]
    
    except Exception as e:
        logger.error(f"无法获取字号名称: {e}", exc_info=True)
        return f"{Pt(sz_val/2):.1f}磅"  # 显示实际磅值作为备选

# 定义命名空间映射 - 统一管理
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'v': 'urn:schemas-microsoft-com:vml',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o': 'urn:schemas-microsoft-com:office:office',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

# 定义空字体字典模板
EMPTY_FONTS_DICT = {
    'eastAsia': None,     # 中文字体
    'ascii': None,        # 英文字体
    'hAnsi': None,        # 通用字体
    'sz': None,           # 字号
    'b': None,            # 加粗
    'i': None,            # 斜体
    'u': None,            # 下划线类型
    'uColor': None,       # 下划线颜色
    'color': None,        # 字体颜色
    'highlight': None,    # 背景高亮颜色
    'strike': None,       # 删除线
    'dStrike': None,      # 双删除线
    'vertAlign': None,    # 垂直对齐方式
    'align': None,        # 水平对齐方式
    'indent': None,       # 缩进
    'spacing': None,      # 行距
    'shd': None,          # 底纹
    'lang': None,         # 语言
    'position': None,     # 字符位置
    'kern': None,         # 字距调整
    'emboss': None,       # 浮雕效果
    'vanish': None,       # 隐藏文字
    'caps': None,         # 全部大写
    'smallCaps': None,    # 小型大写字母
    'shadow': None,       # 阴影
    'outline': None,      # 轮廓
    'rtl': None,          # 从右到左文本
    'cs': None,           # 复杂脚本
    'snapToGrid': None,   # 对齐网格
    'fitText': None       # 调整文本
}

# 硬编码默认字体设置
DEFAULT_FONTS = {
    "eastAsia": "宋体",
    "ascii": "Times New Roman", 
    "hAnsi": "Times New Roman",
    "sz": 24,
    "b": False,
    "i": False,
    "u": "none"
}

def extract_format_from_rpr(rPr, source_name=""):
    """
    从rPr元素中提取格式信息
    返回包含格式的字典，未提取到的字段为None
    """
    # 创建空字典副本
    fonts = EMPTY_FONTS_DICT.copy()
    
    if rPr is None:
        return fonts
    
    try:
        # 提取字体设置
        rFonts = rPr.find('.//w:rFonts', namespaces=NAMESPACES)
        if rFonts is not None:
            # 中文字体
            if qn('w:eastAsia') in rFonts.attrib:
                fonts['eastAsia'] = rFonts.get(qn('w:eastAsia'))
                if source_name and fonts['eastAsia']:
                    logger.debug(f"从{source_name}提取中文字体: {fonts['eastAsia']}")
            
            # 英文字体
            if qn('w:ascii') in rFonts.attrib:
                fonts['ascii'] = rFonts.get(qn('w:ascii'))
                if source_name and fonts['ascii']:
                    logger.debug(f"从{source_name}提取英文字体: {fonts['ascii']}")
            
            # 通用字体
            if qn('w:hAnsi') in rFonts.attrib:
                fonts['hAnsi'] = rFonts.get(qn('w:hAnsi'))
                if source_name and fonts['hAnsi']:
                    logger.debug(f"从{source_name}提取通用字体: {fonts['hAnsi']}")

        # 提取字号
        sz = rPr.find('.//w:sz', namespaces=NAMESPACES)
        if sz is not None and qn('w:val') in sz.attrib:
            fonts['sz'] = sz.get(qn('w:val'))
            if source_name and fonts['sz']:
                logger.debug(f"从{source_name}提取字号: {fonts['sz']}")
        
        # 提取加粗
        b = rPr.find('.//w:b', namespaces=NAMESPACES)
        fonts['b'] = b is not None
        if source_name and b:
            logger.debug(f"从{source_name}提取加粗属性")
        
        # 提取倾斜
        i = rPr.find('.//w:i', namespaces=NAMESPACES)
        fonts['i'] = i is not None
        if source_name and i:
            logger.debug(f"从{source_name}提取倾斜属性")
        
        # 提取下划线
        u = rPr.find('.//w:u', namespaces=NAMESPACES)
        if u is not None:
            fonts['u'] = u.get(qn('w:val'))
            if source_name and fonts['u']:
                logger.debug(f"从{source_name}提取下划线: {fonts['u']}")
            
            # 提取下划线颜色
            u_color = u.get(qn('w:color'))
            if u_color:
                fonts['uColor'] = u_color
                if source_name and fonts['uColor']:
                    logger.debug(f"从{source_name}提取下划线颜色: {fonts['uColor']}")
        
        # 提取字体颜色
        color = rPr.find('.//w:color', namespaces=NAMESPACES)
        if color is not None and qn('w:val') in color.attrib:
            fonts['color'] = color.get(qn('w:val'))
            if source_name and fonts['color']:
                logger.debug(f"从{source_name}提取字体颜色: {fonts['color']}")
        
        # 提取背景高亮颜色
        highlight = rPr.find('.//w:highlight', namespaces=NAMESPACES)
        if highlight is not None and qn('w:val') in highlight.attrib:
            fonts['highlight'] = highlight.get(qn('w:val'))
            if source_name and fonts['highlight']:
                logger.debug(f"从{source_name}提取高亮颜色: {fonts['highlight']}")
        
        # 提取删除线
        strike = rPr.find('.//w:strike', namespaces=NAMESPACES)
        if strike is not None:
            fonts['strike'] = strike.get(qn('w:val'), 'true') == 'true'
            if source_name and fonts['strike']:
                logger.debug(f"从{source_name}提取删除线")
        
        # 提取双删除线
        d_strike = rPr.find('.//w:dstrike', namespaces=NAMESPACES)
        if d_strike is not None:
            fonts['dStrike'] = d_strike.get(qn('w:val'), 'true') == 'true'
            if source_name and fonts['dStrike']:
                logger.debug(f"从{source_name}提取双删除线")
        
        # 提取垂直对齐方式
        vert_align = rPr.find('.//w:vertAlign', namespaces=NAMESPACES)
        if vert_align is not None and qn('w:val') in vert_align.attrib:
            fonts['vertAlign'] = vert_align.get(qn('w:val'))
            if source_name and fonts['vertAlign']:
                logger.debug(f"从{source_name}提取垂直对齐: {fonts['vertAlign']}")
        
        # 提取字符位置
        position = rPr.find('.//w:position', namespaces=NAMESPACES)
        if position is not None and qn('w:val') in position.attrib:
            fonts['position'] = position.get(qn('w:val'))
            if source_name and fonts['position']:
                logger.debug(f"从{source_name}提取字符位置: {fonts['position']}")
        
        # 提取字距调整
        kern = rPr.find('.//w:kern', namespaces=NAMESPACES)
        if kern is not None and qn('w:val') in kern.attrib:
            fonts['kern'] = kern.get(qn('w:val'))
            if source_name and fonts['kern']:
                logger.debug(f"从{source_name}提取字距: {fonts['kern']}")
    
    except Exception as e:
        logger.warning(f"从rPr提取格式失败: {str(e)}")
    
    return fonts

def extract_format_from_run(run):
    """
    从单个run中提取格式信息
    返回包含格式的字典，未提取到的字段为None
    """
    try:
        rPr = run._element.rPr
        return extract_format_from_rpr(rPr)
    except Exception as e:
        logger.error(f"从run提取格式失败: {str(e)}")
        return EMPTY_FONTS_DICT.copy()


def get_effective_run_format(paragraph, para_fonts):
    """
    获取段落中第一个有效Run的格式
    :param paragraph: 段落对象
    :param para_fonts: 段落基础格式
    :return: 最终格式字典
    """
    # 确保 para_fonts 不为 None
    if para_fonts is None:
        para_fonts = {}
        logger.warning("段落基础格式为空，使用空字典")
    
    # 查找第一个包含中文字符的Run
    for run in paragraph.runs:
        if contains_chinese(run.text):
            try:
                # 提取Run的格式
                run_fonts = extract_format_from_run(run)
                
                # 确保 run_fonts 不为 None
                if run_fonts is None:
                    run_fonts = {}
                    logger.warning("Run格式为空，使用空字典")
                
                # 用Run格式覆盖段落格式
                for key in ['eastAsia', 'ascii', 'sz', 'color']:
                    if run_fonts.get(key) is not None:
                        para_fonts[key] = run_fonts[key]
                        logger.debug(f"使用Run格式覆盖 {key}: {run_fonts[key]}")

                logger.info(f"找到有效Run: 中文字体={para_fonts.get('eastAsia')}, 英文字体={para_fonts.get('ascii')}, 字号={get_font_size_name(para_fonts.get('sz'))}")
                return para_fonts
            except Exception as e:
                logger.error(f"提取Run格式失败: {str(e)}")
                continue
    
    logger.info("未找到有效Run，使用段落基础格式")
    return para_fonts


def extract_format_from_paragraph_style(paragraph):
    """
    提取段落格式信息
    按照顺序提取格式信息：
    1. 从段落样式中提取
    2. 从段落样式的父样式中提取
    3. 从文档默认样式中提取
    4. 使用默认设置
    
    返回格式字典，所有字段都有值
    """
    logger.debug(f"开始提取段落格式信息")
    # 创建空字典副本
    fonts = EMPTY_FONTS_DICT.copy()
    
    # 1. 从段落样式中提取
    if paragraph.style is not None:
        logger.debug("== 从段落样式中提取格式信息 ==")
        try:
            if paragraph.style.element is not None:
                # 提取字符格式
                rPr = paragraph.style.element.find('.//w:rPr', namespaces=NAMESPACES)
                style_fonts = extract_format_from_rpr(rPr, "段落样式")
                
                # 更新fonts字典
                for key in fonts:
                    if fonts[key] is None and style_fonts.get(key) is not None:
                        fonts[key] = style_fonts[key]
        except Exception as e:
            logger.warning(f"从段落样式提取格式失败: {str(e)}")
            
        if all(fonts[key] is not None for key in ['eastAsia', 'ascii', 'sz']):
            return fonts

    # 2. 从段落样式的父样式中提取
    current_style = paragraph.style
    while current_style is not None and hasattr(current_style, 'base_style'):
        if current_style.base_style is not None:
            logger.debug(f"== 从父样式({current_style.base_style.name})中提取格式信息 ==")
            try:
                if current_style.base_style.element is not None:
                    rPr = current_style.base_style.element.find('.//w:rPr', namespaces=NAMESPACES)
                    parent_fonts = extract_format_from_rpr(rPr, "父样式")
                    
                    # 更新fonts字典
                    for key in fonts:
                        if fonts[key] is None and parent_fonts.get(key) is not None:
                            fonts[key] = parent_fonts[key]
            except Exception as e:
                logger.warning(f"从父样式提取格式失败: {str(e)}")
            
            if all(fonts[key] is not None for key in ['eastAsia', 'ascii', 'sz']):
                return fonts
            current_style = current_style.base_style
        else:
            break

    # 3. 从文档默认样式中提取
    try:
        doc = paragraph.part.document
        default_style = doc.styles['Normal']
        logger.debug("== 从默认样式中提取格式信息 ==")
        if default_style.element is not None:
            rPr = default_style.element.find('.//w:rPr', namespaces=NAMESPACES)
            default_fonts = extract_format_from_rpr(rPr, "默认样式")
            
            # 更新fonts字典
            for key in fonts:
                if fonts[key] is None and default_fonts.get(key) is not None:
                    fonts[key] = default_fonts[key]
    except Exception as e:
        logger.warning(f"获取默认样式失败: {str(e)}")

    # 4. 使用默认设置填充缺失值
    for key in fonts:
        if fonts[key] is None:
            fonts[key] = DEFAULT_FONTS.get(key, None)
            logger.debug(f"使用默认值填充 {key}: {DEFAULT_FONTS.get(key, '')}")

    logger.info(f"段落格式设置 - 中文字体: {fonts['eastAsia']}, 英文字体: {fonts['ascii']}, 字号: {fonts['sz']}")
    return fonts

def extract_format_from_paragraph(paragraph):
    """
    提取完整格式信息（字符格式+段落格式）
    按照顺序提取格式信息：
    1. 从run中提取
    2. 从段落格式中提取
    3. 使用默认设置
    
    返回格式字典，所有字段都有值
    """
    logger.debug(f"开始提取完整格式信息")
    # 创建空字典副本
    fonts = EMPTY_FONTS_DICT.copy()
    
    # 1. 首先尝试从run中提取
    if paragraph.runs:
        logger.debug("== 优先从run中提取格式信息 ==")
        for run in paragraph.runs:
            # 获取run的格式
            run_fonts = extract_format_from_run(run)
            
            # 更新fonts字典
            for key in fonts:
                if fonts[key] is None and run_fonts.get(key) is not None:
                    fonts[key] = run_fonts[key]
            
            # 如果已经找到所有需要的属性，提前返回
            if all(fonts[key] is not None for key in ['eastAsia', 'ascii', 'sz']):
                return fonts

    # 2. 从段落格式中提取
    logger.debug("== 从段落格式中提取格式信息 ==")
    para_fonts = extract_format_from_paragraph_style(paragraph)
    for key in fonts:
        if fonts[key] is None and para_fonts.get(key) is not None:
            fonts[key] = para_fonts[key]

    # 3. 使用默认设置填充缺失值
    for key in fonts:
        if fonts[key] is None:
            fonts[key] = DEFAULT_FONTS.get(key, None)
            logger.debug(f"使用默认值填充 {key}: {DEFAULT_FONTS.get(key, '')}")

    logger.info(f"最终格式设置 - 中文字体: {fonts['eastAsia']}, 英文字体: {fonts['ascii']}, 字号: {fonts['sz']}")
    return fonts



def apply_body_formatting(run, fonts):
    """应用格式到文本run"""
    logger.debug(f"开始应用格式到文本run: {run.text[:50]}")
    
    # 获取或创建rPr元素
    rPr = run._element.get_or_add_rPr()
    
    # 设置字体
    set_font_family(rPr, fonts)
    
    # 设置字号
    set_font_size(rPr, fonts)
    
    # 设置加粗
    set_font_bold(rPr, fonts)
    
    # 设置倾斜
    set_font_italic(rPr, fonts)
    
    # 设置下划线
    set_font_underline(rPr, fonts)
    
    logger.debug(f"文本格式设置完成")

def apply_run_style(run, source_rPr):
    """
    应用样式到run（安全方法）
    """
    if source_rPr is None:
        return
    
    # 确保目标run有rPr元素
    rPr = run._element.get_or_add_rPr()
    
    # 清除现有的所有属性
    for child in rPr.getchildren():
        rPr.remove(child)
    
    # 复制源样式的所有属性
    for child in source_rPr:
        new_child = copy.deepcopy(child)
        rPr.append(new_child)


def apply_paragraph_format(target_para, source_para):
    """
    应用段落格式（与 apply_run_style 配套）
    
    参数:
        target_para: 目标段落对象
        source_para: 源段落对象
        
    详细说明:
        复制段落样式、对齐、缩进、间距、分页控制等所有格式属性。
        与 apply_run_style 配套使用，确保段落和run格式都被完整复制。
    """
    try:
        # 1. 复制段落样式
        if source_para.style:
            target_para.style = source_para.style
            logger.debug(f"复制段落样式: {source_para.style.name if hasattr(source_para.style, 'name') else source_para.style}")
        
        # 2. 复制对齐方式
        if source_para.alignment is not None:
            target_para.alignment = source_para.alignment
        
        # 3. 复制段落格式属性
        pf_src = source_para.paragraph_format
        pf_tgt = target_para.paragraph_format
        
        # 缩进
        if pf_src.left_indent is not None:
            pf_tgt.left_indent = pf_src.left_indent
        if pf_src.right_indent is not None:
            pf_tgt.right_indent = pf_src.right_indent
        if pf_src.first_line_indent is not None:
            pf_tgt.first_line_indent = pf_src.first_line_indent
        
        # 间距
        if pf_src.space_before is not None:
            pf_tgt.space_before = pf_src.space_before
        if pf_src.space_after is not None:
            pf_tgt.space_after = pf_src.space_after
        if pf_src.line_spacing is not None:
            pf_tgt.line_spacing = pf_src.line_spacing
        if pf_src.line_spacing_rule is not None:
            pf_tgt.line_spacing_rule = pf_src.line_spacing_rule
        
        # 分页控制
        if pf_src.keep_together is not None:
            pf_tgt.keep_together = pf_src.keep_together
        if pf_src.keep_with_next is not None:
            pf_tgt.keep_with_next = pf_src.keep_with_next
        if pf_src.page_break_before is not None:
            pf_tgt.page_break_before = pf_src.page_break_before
        if pf_src.widow_control is not None:
            pf_tgt.widow_control = pf_src.widow_control
        
        logger.debug("成功应用段落格式")
    except Exception as e:
        logger.warning(f"应用段落格式时出现警告: {str(e)}")

def is_rPr_equal(rPr1, rPr2):
    """
    比较两个rPr元素是否相等（样式相同）
    """
    # 两个都为None则相等
    if rPr1 is None and rPr2 is None:
        return True
    
    # 一个为None一个不为None则不相等
    if rPr1 is None or rPr2 is None:
        return False
    
    # 序列化为字符串比较
    try:
        return etree.tostring(rPr1) == etree.tostring(rPr2)
    except:
        return False

def set_font_family(rPr, fonts):
    """设置字体族"""
    # 清除现有设置
    for elem in rPr.xpath('.//w:rFonts'):
        rPr.remove(elem)
    
    # 创建新设置
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), fonts['eastAsia'])
    rFonts.set(qn('w:ascii'), fonts['ascii'])
    rFonts.set(qn('w:hAnsi'), fonts['hAnsi'])
    rPr.append(rFonts)

def set_font_size(rPr, fonts):
    """设置字号"""
    # 清除现有设置
    for elem in rPr.xpath('.//w:sz'):
        rPr.remove(elem)
    
    # 创建新设置
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(fonts['sz']))
    rPr.append(sz)

def set_font_bold(rPr, fonts):
    """设置加粗"""
    # 清除现有设置
    for elem in rPr.xpath('.//w:b'):
        rPr.remove(elem)
    
    if fonts['b']:
        b = OxmlElement('w:b')
        rPr.append(b)

def set_font_italic(rPr, fonts):
    """设置倾斜"""
    # 清除现有设置
    for elem in rPr.xpath('.//w:i'):
        rPr.remove(elem)
    
    if fonts['i']:
        i = OxmlElement('w:i')
        rPr.append(i)

def set_font_underline(rPr, fonts):
    """设置下划线"""
    # 清除现有设置
    for elem in rPr.xpath('.//w:u'):
        rPr.remove(elem)
    
    if fonts['u']:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), fonts['u'])
        rPr.append(u)

def register_namespaces(root_element=None):
    """注册所有命名空间到XML处理"""
    # 默认注册所有命名空间
    for prefix, uri in NAMESPACES.items():
        etree.register_namespace(prefix, uri)
    
    # 如果提供了根元素，检查并添加缺失的命名空间声明
    if root_element is not None:
        for prefix, uri in NAMESPACES.items():
            # 检查是否已声明
            ns_declaration = f"xmlns:{prefix}"
            if ns_declaration not in root_element.attrib:
                root_element.attrib[ns_declaration] = uri

def create_element(tag, prefix='w'):
    """创建带命名空间的元素"""
    return etree.Element(f'{{{NAMESPACES[prefix]}}}{tag}')
