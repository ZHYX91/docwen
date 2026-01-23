"""
公式格式转换工具模块

提供LaTeX、MathML、OMML（Office Math XML）之间的转换功能。

主要功能：
- LaTeX → MathML：使用latex2mathml库
- MathML → OMML：使用自定义转换逻辑
- OMML → MathML：使用自定义转换逻辑
- MathML → LaTeX：使用自定义转换逻辑
"""

import re
import logging
from typing import Optional, Tuple, Dict
from lxml import etree

try:
    from latex2mathml.converter import convert as latex_to_mathml_convert
    LATEX2MATHML_AVAILABLE = True
except ImportError:
    LATEX2MATHML_AVAILABLE = False
    logging.warning("latex2mathml未安装，公式转换功能将不可用")

logger = logging.getLogger(__name__)

# Office Math命名空间
OMML_NS = {
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

# MathML命名空间
MATHML_NS = 'http://www.w3.org/1998/Math/MathML'

# 特殊符号映射表 (Unicode -> LaTeX)
SYMBOL_MAP = {
    # 希腊字母
    'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta', 'ε': r'\epsilon', 'ζ': r'\zeta',
    'η': r'\eta', 'θ': r'\theta', 'ι': r'\iota', 'κ': r'\kappa', 'λ': r'\lambda', 'μ': r'\mu',
    'ν': r'\nu', 'ξ': r'\xi', 'ο': r'o', 'π': r'\pi', 'ρ': r'\rho', 'σ': r'\sigma',
    'τ': r'\tau', 'υ': r'\upsilon', 'φ': r'\phi', 'χ': r'\chi', 'ψ': r'\psi', 'ω': r'\omega',
    'Α': r'A', 'Β': r'B', 'Γ': r'\Gamma', 'Δ': r'\Delta', 'Ε': r'E', 'Ζ': r'Z',
    'Η': r'H', 'Θ': r'\Theta', 'Ι': r'I', 'Κ': r'K', 'Λ': r'\Lambda', 'Μ': r'M',
    'Ν': r'N', 'Ξ': r'\Xi', 'Ο': r'O', 'Π': r'\Pi', 'Ρ': r'P', 'Σ': r'\Sigma',
    'Τ': r'T', 'Υ': r'\Upsilon', 'Φ': r'\Phi', 'Χ': r'X', 'Ψ': r'\Psi', 'Ω': r'\Omega',
    
    # 数学符号
    '∑': r'\sum', '∏': r'\prod', '∫': r'\int', '∬': r'\iint', '∮': r'\oint',
    '∂': r'\partial', '∇': r'\nabla', '∞': r'\infty', 'lim': r'\lim',
    '→': r'\to', '←': r'\gets', '↔': r'\leftrightarrow', '⇒': r'\Rightarrow',
    '≈': r'\approx', '≠': r'\neq', '≤': r'\leq', '≥': r'\geq', '±': r'\pm', '×': r'\times', '÷': r'\div',
    '∈': r'\in', '∉': r'\notin', '⊂': r'\subset', '⊃': r'\supset', '∪': r'\cup', '∩': r'\cap',
    '∀': r'\forall', '∃': r'\exists', '∅': r'\emptyset',
}

# 运算符映射表 (用于 n-ary 操作符)
NARY_OP_MAP = {
    '∑': r'\sum',
    '∏': r'\prod',
    '∫': r'\int',
    '⋃': r'\bigcup',
    '⋂': r'\bigcap',
}

def is_formula_supported() -> bool:
    """检查公式转换功能是否可用"""
    return LATEX2MATHML_AVAILABLE


# 矩阵环境到括号的映射表
MATRIX_BRACKET_MAP = {
    'bmatrix': ('[', ']'),      # 方括号
    'pmatrix': ('(', ')'),      # 圆括号
    'vmatrix': ('|', '|'),      # 单竖线（行列式）
    'Vmatrix': ('‖', '‖'),     # 双竖线（范数）- 使用Unicode双竖线
    'Bmatrix': ('{', '}'),      # 花括号
    'matrix': ('', ''),         # 无括号
}


def _convert_matrix_latex_to_mathml(latex_str: str) -> Optional[str]:
    """
    手动将LaTeX矩阵转换为MathML
    支持: bmatrix, pmatrix, matrix, vmatrix, Vmatrix, Bmatrix
    
    注意：只处理单个矩阵的情况。如果LaTeX包含多个矩阵或复杂表达式，
    应该让latex2mathml库处理。
    """
    try:
        # 检查是否包含多个矩阵
        matrix_count = len(re.findall(r'\\begin\{[bpBvV]?matrix\}', latex_str))
        if matrix_count > 1:
            logger.debug(f"检测到{matrix_count}个矩阵，不使用手动转换")
            return None  # 让latex2mathml处理
        
        # 简化处理：提取矩阵内容 - 支持所有矩阵类型
        pattern = r'\\begin\{([bpBvV]?matrix)\}(.+?)\\end\{\1\}'
        match = re.search(pattern, latex_str, re.DOTALL)
        
        if not match:
            return None
        
        matrix_type = match.group(1)
        matrix_content = match.group(2).strip()
        
        # 检查是否只包含这个矩阵（没有其他内容）
        if match.group(0).strip() != latex_str.strip():
            logger.debug("LaTeX包含矩阵外的其他内容，不使用手动转换")
            return None  # 让latex2mathml处理复杂情况
        
        # 解析矩阵行（用 \\ 分割）
        rows = [row.strip() for row in matrix_content.split(r'\\') if row.strip()]
        
        # 创建MathML
        math = etree.Element('math', xmlns=MATHML_NS)
        
        # 获取括号类型
        open_bracket, close_bracket = MATRIX_BRACKET_MAP.get(matrix_type, ('', ''))
        
        # 根据类型添加外围括号
        if open_bracket and close_bracket:
            # 使用mfenced包装
            mfenced = etree.SubElement(math, 'mfenced')
            mfenced.set('open', open_bracket)
            mfenced.set('close', close_bracket)
            table = etree.SubElement(mfenced, 'mtable')
        else:
            table = etree.SubElement(math, 'mtable')
        
        # 添加行
        for row_str in rows:
            mtr = etree.SubElement(table, 'mtr')
            # 分割单元格（用 & 分割）
            cells = [cell.strip() for cell in row_str.split('&')]
            for cell_str in cells:
                mtd = etree.SubElement(mtr, 'mtd')
                # 将单元格内容转为MathML（递归处理复杂内容）
                if cell_str:
                    _convert_cell_content_to_mathml(cell_str, mtd)
        
        logger.info(f"手动转换单个矩阵成功: {matrix_type}")
        return etree.tostring(math, encoding='unicode')
    except Exception as e:
        logger.error(f"手动矩阵转换失败: {e}")
        return None


def _convert_cell_content_to_mathml(cell_str: str, parent_elem) -> None:
    """
    将矩阵单元格内容转换为MathML节点
    
    支持：
    - 简单数字: 1, 2, 3
    - 简单变量: a, b, c
    - 下标: a_{11}, x_1
    - 上标: x^2
    - 希腊字母: \alpha, \beta
    """
    cell_str = cell_str.strip()
    
    if not cell_str:
        return
    
    # 尝试用 latex2mathml 处理复杂内容
    if LATEX2MATHML_AVAILABLE and ('_' in cell_str or '^' in cell_str or '\\' in cell_str):
        try:
            # 使用 latex2mathml 转换单元格内容
            mathml_str = latex_to_mathml_convert(cell_str)
            if mathml_str:
                # 解析生成的 MathML 并提取内容
                temp_tree = etree.fromstring(mathml_str.encode('utf-8'))
                # 将 math 元素的子节点复制到 mtd
                for child in temp_tree:
                    parent_elem.append(child)
                return
        except Exception as e:
            logger.debug(f"单元格 latex2mathml 转换失败，使用简单处理: {e}")
    
    # 简单处理：直接创建mi/mn
    if cell_str.isdigit() or (cell_str.startswith('-') and cell_str[1:].isdigit()):
        mn = etree.SubElement(parent_elem, 'mn')
        mn.text = cell_str
    else:
        mi = etree.SubElement(parent_elem, 'mi')
        mi.text = cell_str


def latex_to_mathml(latex_str: str) -> Optional[str]:
    """将LaTeX公式转换为MathML格式"""
    if not LATEX2MATHML_AVAILABLE:
        logger.error("latex2mathml库未安装")
        return None
    
    try:
        latex_str = latex_str.strip()
        if not latex_str:
            return None
        
        # 预处理：检测矩阵环境
        # 如果包含矩阵环境，尝试手动转换
        if r'\begin{' in latex_str and 'matrix' in latex_str:
            logger.debug(f"检测到矩阵环境，尝试手动转换")
            mathml = _convert_matrix_latex_to_mathml(latex_str)
            if mathml:
                return mathml
            # 如果手动转换失败，继续尝试latex2mathml
        
        mathml = latex_to_mathml_convert(latex_str)
        return mathml
    except Exception as e:
        logger.error(f"LaTeX转MathML失败: {e}, LaTeX: {latex_str}")
        # 如果latex2mathml失败，且是矩阵，尝试手动转换
        if r'\begin{' in latex_str and 'matrix' in latex_str:
            logger.debug(f"latex2mathml失败，尝试手动矩阵转换")
            return _convert_matrix_latex_to_mathml(latex_str)
        return None


def mathml_to_omml(mathml_str: str) -> Optional[str]:
    """将MathML转换为OMML格式"""
    try:
        mathml_tree = etree.fromstring(mathml_str.encode('utf-8'))
        omml_math = etree.Element('{%s}oMath' % OMML_NS['m'], nsmap={'m': OMML_NS['m']})
        _convert_mathml_to_omml_node(mathml_tree, omml_math)
        return etree.tostring(omml_math, encoding='unicode', pretty_print=True)
    except Exception as e:
        logger.error(f"MathML转OMML失败: {e}")
        return None


# 括号配对映射 (用于识别括号-内容-括号模式)
BRACKET_PAIRS = {
    '[': ']',
    '(': ')',
    '{': '}',
    '|': '|',
    '‖': '‖',  # 双竖线
    '||': '||',
}


def _process_children_to_omml(children: list, omml_parent) -> None:
    """
    处理MathML子节点列表，识别括号-内容-括号模式并转换为OMML
    
    主要功能：
    1. 扫描子节点列表，识别 mo(左括号) + 内容 + mo(右括号) 模式
    2. 将此模式转换为 OMML 的 <m:d> (delimiter) 结构
    3. 非括号模式的节点直接递归转换
    
    这个函数用于处理 latex2mathml 生成的非标准结构，如：
    - 矩阵带括号：[mtable] 变成 mo([) + mtable + mo(])
    - 矩阵求逆：[矩阵]^{-1} 在 msup 中有多个子节点
    """
    i = 0
    while i < len(children):
        child = children[i]
        child_tag = etree.QName(child).localname
        
        # 检查是否是括号模式的开始
        if child_tag == 'mo':
            open_char = (child.text or '').strip()
            
            # 检查是否是左括号
            if open_char in BRACKET_PAIRS:
                expected_close = BRACKET_PAIRS[open_char]
                
                # 查找匹配的右括号
                close_idx = -1
                for j in range(i + 2, len(children)):
                    check_child = children[j]
                    if etree.QName(check_child).localname == 'mo':
                        close_text = (check_child.text or '').strip()
                        if close_text == expected_close:
                            close_idx = j
                            break
                
                if close_idx > i + 1:
                    # 找到匹配的括号对，创建 OMML delimiter 结构
                    logger.debug(f"检测到括号模式: {open_char}...{expected_close}，内容节点数: {close_idx - i - 1}")
                    
                    d_elem = etree.SubElement(omml_parent, '{%s}d' % OMML_NS['m'])
                    
                    # 设置括号属性
                    dPr = etree.SubElement(d_elem, '{%s}dPr' % OMML_NS['m'])
                    begChr = etree.SubElement(dPr, '{%s}begChr' % OMML_NS['m'])
                    begChr.set('{%s}val' % OMML_NS['m'], open_char)
                    endChr = etree.SubElement(dPr, '{%s}endChr' % OMML_NS['m'])
                    endChr.set('{%s}val' % OMML_NS['m'], expected_close)
                    
                    # 转换括号内的内容
                    e_elem = etree.SubElement(d_elem, '{%s}e' % OMML_NS['m'])
                    for content_idx in range(i + 1, close_idx):
                        _convert_mathml_to_omml_node(children[content_idx], e_elem)
                    
                    # 跳过已处理的元素
                    i = close_idx + 1
                    continue
        
        # 非括号模式，直接转换
        _convert_mathml_to_omml_node(child, omml_parent)
        i += 1


def _convert_mathml_to_omml_node(mathml_node, omml_parent):
    """递归转换MathML节点为OMML节点"""
    tag = etree.QName(mathml_node).localname
    
    if tag == 'math':
        children = list(mathml_node)
        # 使用增强的子节点处理函数
        _process_children_to_omml(children, omml_parent)
    
    elif tag in ['mi', 'mn', 'mo']:
        run = etree.SubElement(omml_parent, '{%s}r' % OMML_NS['m'])
        text_elem = etree.SubElement(run, '{%s}t' % OMML_NS['m'])
        text_elem.text = mathml_node.text or ''
    
    elif tag == 'mfrac':
        frac = etree.SubElement(omml_parent, '{%s}f' % OMML_NS['m'])
        num = etree.SubElement(frac, '{%s}num' % OMML_NS['m'])
        den = etree.SubElement(frac, '{%s}den' % OMML_NS['m'])
        children = list(mathml_node)
        if len(children) >= 2:
            _convert_mathml_to_omml_node(children[0], num)
            _convert_mathml_to_omml_node(children[1], den)
            
    # 上标 - 适配非标准结构（latex2mathml 可能生成 >2 个子节点）
    elif tag == 'msup':
        ssup = etree.SubElement(omml_parent, '{%s}sSup' % OMML_NS['m'])
        base = etree.SubElement(ssup, '{%s}e' % OMML_NS['m'])
        sup = etree.SubElement(ssup, '{%s}sup' % OMML_NS['m'])
        children = list(mathml_node)
        
        if len(children) > 2:
            # 非标准结构：最后一个是上标，前面所有是底数
            # 例如：[矩阵]^{-1} -> mo([) + mtable + mo(]) + mrow(-1)
            logger.debug(f"msup 非标准结构：{len(children)} 个子节点，使用增强处理")
            _process_children_to_omml(children[:-1], base)
            _convert_mathml_to_omml_node(children[-1], sup)
        elif len(children) == 2:
            # 标准结构
            _convert_mathml_to_omml_node(children[0], base)
            _convert_mathml_to_omml_node(children[1], sup)
            
    # 下标 - 适配非标准结构
    elif tag == 'msub':
        ssub = etree.SubElement(omml_parent, '{%s}sSub' % OMML_NS['m'])
        base = etree.SubElement(ssub, '{%s}e' % OMML_NS['m'])
        sub = etree.SubElement(ssub, '{%s}sub' % OMML_NS['m'])
        children = list(mathml_node)
        
        if len(children) > 2:
            # 非标准结构：最后一个是下标，前面所有是底数
            logger.debug(f"msub 非标准结构：{len(children)} 个子节点，使用增强处理")
            _process_children_to_omml(children[:-1], base)
            _convert_mathml_to_omml_node(children[-1], sub)
        elif len(children) == 2:
            # 标准结构
            _convert_mathml_to_omml_node(children[0], base)
            _convert_mathml_to_omml_node(children[1], sub)
            
    # 同时上下标 - 适配非标准结构
    elif tag == 'msubsup':
        subsup = etree.SubElement(omml_parent, '{%s}sSubSup' % OMML_NS['m'])
        base = etree.SubElement(subsup, '{%s}e' % OMML_NS['m'])
        sub = etree.SubElement(subsup, '{%s}sub' % OMML_NS['m'])
        sup = etree.SubElement(subsup, '{%s}sup' % OMML_NS['m'])
        children = list(mathml_node)
        
        if len(children) > 3:
            # 非标准结构：最后两个分别是下标和上标，前面所有是底数
            logger.debug(f"msubsup 非标准结构：{len(children)} 个子节点，使用增强处理")
            _process_children_to_omml(children[:-2], base)
            _convert_mathml_to_omml_node(children[-2], sub)
            _convert_mathml_to_omml_node(children[-1], sup)
        elif len(children) >= 3:
            # 标准结构
            _convert_mathml_to_omml_node(children[0], base)
            _convert_mathml_to_omml_node(children[1], sub)
            _convert_mathml_to_omml_node(children[2], sup)

    # 根式
    elif tag == 'msqrt':
        rad = etree.SubElement(omml_parent, '{%s}rad' % OMML_NS['m'])
        deg = etree.SubElement(rad, '{%s}deg' % OMML_NS['m']) # 默认为空（平方根）
        base = etree.SubElement(rad, '{%s}e' % OMML_NS['m'])
        for child in mathml_node:
            _convert_mathml_to_omml_node(child, base)
            
    # 根式 (n次根)
    elif tag == 'mroot':
        rad = etree.SubElement(omml_parent, '{%s}rad' % OMML_NS['m'])
        deg = etree.SubElement(rad, '{%s}deg' % OMML_NS['m'])
        base = etree.SubElement(rad, '{%s}e' % OMML_NS['m'])
        children = list(mathml_node)
        if len(children) >= 2:
            _convert_mathml_to_omml_node(children[0], base) # 底数
            _convert_mathml_to_omml_node(children[1], deg)  # 次数

    # 行
    elif tag == 'mrow':
        for child in mathml_node:
            _convert_mathml_to_omml_node(child, omml_parent)
    
    # 矩阵 (MathML mtable → OMML m)
    elif tag == 'mtable':
        # 创建OMML矩阵
        m_elem = etree.SubElement(omml_parent, '{%s}m' % OMML_NS['m'])
        
        # 遍历所有行
        for mtr in mathml_node:
            if etree.QName(mtr).localname == 'mtr':
                mr = etree.SubElement(m_elem, '{%s}mr' % OMML_NS['m'])
                
                # 遍历行中的所有单元格
                for mtd in mtr:
                    if etree.QName(mtd).localname == 'mtd':
                        e_elem = etree.SubElement(mr, '{%s}e' % OMML_NS['m'])
                        
                        # 转换单元格内容
                        for child in mtd:
                            _convert_mathml_to_omml_node(child, e_elem)
    
    # mfenced（括号包围的内容，如矩阵外围括号）
    elif tag == 'mfenced':
        # 提取open和close属性
        open_char = mathml_node.get('open', '(')
        close_char = mathml_node.get('close', ')')
        
        # 创建带括号的矩阵结构
        d_elem = etree.SubElement(omml_parent, '{%s}d' % OMML_NS['m'])
        
        # 设置括号属性
        dPr = etree.SubElement(d_elem, '{%s}dPr' % OMML_NS['m'])
        begChr = etree.SubElement(dPr, '{%s}begChr' % OMML_NS['m'])
        begChr.set('{%s}val' % OMML_NS['m'], open_char)
        endChr = etree.SubElement(dPr, '{%s}endChr' % OMML_NS['m'])
        endChr.set('{%s}val' % OMML_NS['m'], close_char)
        
        # 添加内容（通常是mtable）
        e_elem = etree.SubElement(d_elem, '{%s}e' % OMML_NS['m'])
        for child in mathml_node:
            _convert_mathml_to_omml_node(child, e_elem)
            
    # 默认处理
    else:
        for child in mathml_node:
            _convert_mathml_to_omml_node(child, omml_parent)


def omml_to_mathml(omml_str: str) -> Optional[str]:
    """将OMML转换为MathML格式"""
    try:
        omml_tree = etree.fromstring(omml_str.encode('utf-8'))
        mathml_math = etree.Element('math', xmlns=MATHML_NS)
        _convert_omml_to_mathml_node(omml_tree, mathml_math)
        return etree.tostring(mathml_math, encoding='unicode', pretty_print=True)
    except Exception as e:
        logger.error(f"OMML转MathML失败: {e}")
        return None


def _get_child_by_tag(element, tag_name):
    """辅助函数：按标签名查找直接子元素"""
    for child in element:
        if etree.QName(child).localname == tag_name:
            return child
    return None

def _convert_omml_to_mathml_node(omml_node, mathml_parent):
    """递归转换OMML节点为MathML节点"""
    tag = etree.QName(omml_node).localname
    
    if tag == 'oMath':
        for child in omml_node:
            _convert_omml_to_mathml_node(child, mathml_parent)
    
    elif tag == 'r':
        # Run → mi/mn/mo
        t_elem = _get_child_by_tag(omml_node, 't')
        
        if t_elem is not None and t_elem.text:
            text = t_elem.text.strip()
            # 区分运算符、数字、标识符
            if text in SYMBOL_MAP or text in NARY_OP_MAP or not text.isalnum():
                elem = etree.SubElement(mathml_parent, 'mo')
            elif text.isdigit():
                elem = etree.SubElement(mathml_parent, 'mn')
            else:
                elem = etree.SubElement(mathml_parent, 'mi')
            elem.text = text
    
    elif tag == 'f':
        # 分式：将分子分母包装在mrow容器中保持MathML结构完整
        frac = etree.SubElement(mathml_parent, 'mfrac')
        num_elem = _get_child_by_tag(omml_node, 'num')
        den_elem = _get_child_by_tag(omml_node, 'den')
        
        if num_elem is not None:
            # 为分子创建mrow容器
            num_row = etree.SubElement(frac, 'mrow')
            for child in num_elem:
                _convert_omml_to_mathml_node(child, num_row)
        else:
            # 如果没有分子，添加空的mrow
            etree.SubElement(frac, 'mrow')
                
        if den_elem is not None:
            # 为分母创建mrow容器
            den_row = etree.SubElement(frac, 'mrow')
            for child in den_elem:
                _convert_omml_to_mathml_node(child, den_row)
        else:
            # 如果没有分母，添加空的mrow
            etree.SubElement(frac, 'mrow')
    
    elif tag == 'sSup':
        # 上标：将基础和上标部分包装在mrow容器中
        ssup = etree.SubElement(mathml_parent, 'msup')
        base_elem = _get_child_by_tag(omml_node, 'e')
        sup_elem = _get_child_by_tag(omml_node, 'sup')
        
        # 基础部分
        if base_elem is not None:
            base_row = etree.SubElement(ssup, 'mrow')
            for child in base_elem: 
                _convert_omml_to_mathml_node(child, base_row)
        else:
            etree.SubElement(ssup, 'mrow')
        
        # 上标部分
        if sup_elem is not None:
            sup_row = etree.SubElement(ssup, 'mrow')
            for child in sup_elem: 
                _convert_omml_to_mathml_node(child, sup_row)
        else:
            etree.SubElement(ssup, 'mrow')
            
    elif tag == 'sSub':
        # 下标：将基础和下标部分包装在mrow容器中
        ssub = etree.SubElement(mathml_parent, 'msub')
        base_elem = _get_child_by_tag(omml_node, 'e')
        sub_elem = _get_child_by_tag(omml_node, 'sub')
        
        # 基础部分
        if base_elem is not None:
            base_row = etree.SubElement(ssub, 'mrow')
            for child in base_elem: 
                _convert_omml_to_mathml_node(child, base_row)
        else:
            etree.SubElement(ssub, 'mrow')
        
        # 下标部分
        if sub_elem is not None:
            sub_row = etree.SubElement(ssub, 'mrow')
            for child in sub_elem: 
                _convert_omml_to_mathml_node(child, sub_row)
        else:
            etree.SubElement(ssub, 'mrow')
            
    elif tag == 'sSubSup':
        # 同时上下标：将基础、下标和上标部分分别包装在mrow容器中
        subsup = etree.SubElement(mathml_parent, 'msubsup')
        base_elem = _get_child_by_tag(omml_node, 'e')
        sub_elem = _get_child_by_tag(omml_node, 'sub')
        sup_elem = _get_child_by_tag(omml_node, 'sup')
        
        # 基础部分
        if base_elem is not None:
            base_row = etree.SubElement(subsup, 'mrow')
            for child in base_elem: 
                _convert_omml_to_mathml_node(child, base_row)
        else:
            etree.SubElement(subsup, 'mrow')
        
        # 下标部分
        if sub_elem is not None:
            sub_row = etree.SubElement(subsup, 'mrow')
            for child in sub_elem: 
                _convert_omml_to_mathml_node(child, sub_row)
        else:
            etree.SubElement(subsup, 'mrow')
        
        # 上标部分
        if sup_elem is not None:
            sup_row = etree.SubElement(subsup, 'mrow')
            for child in sup_elem: 
                _convert_omml_to_mathml_node(child, sup_row)
        else:
            etree.SubElement(subsup, 'mrow')

    elif tag == 'nary':
        # N元运算符：处理求和、积分等运算符及其上下标
        naryPr = _get_child_by_tag(omml_node, 'naryPr')
        op_char = '∫'
        lim_loc = 'subSup'
        
        if naryPr is not None:
            chr_elem = _get_child_by_tag(naryPr, 'chr')
            if chr_elem is not None:
                op_char = chr_elem.attrib.get('{%s}val' % OMML_NS['m'], '∫')
            
            lim_loc_elem = _get_child_by_tag(naryPr, 'limLoc')
            if lim_loc_elem is not None:
                lim_loc = lim_loc_elem.attrib.get('{%s}val' % OMML_NS['m'], 'subSup')
        
        # 先处理操作符及其上下标/界限
        mathml_tag = 'munderover' if lim_loc == 'undOvr' else 'msubsup'
        nary_elem = etree.SubElement(mathml_parent, mathml_tag)
        
        # 1. 操作符
        op = etree.SubElement(nary_elem, 'mo')
        op.text = op_char
        
        # 2. 下限/下标
        sub_elem = _get_child_by_tag(omml_node, 'sub')
        sub_row = etree.SubElement(nary_elem, 'mrow')
        if sub_elem is not None:
            for child in sub_elem: 
                _convert_omml_to_mathml_node(child, sub_row)

        # 3. 上限/上标
        sup_elem = _get_child_by_tag(omml_node, 'sup')
        sup_row = etree.SubElement(nary_elem, 'mrow')
        if sup_elem is not None:
            for child in sup_elem: 
                _convert_omml_to_mathml_node(child, sup_row)

        # 4. 基础内容（被积/求和表达式）- 这应该紧跟在nary_elem之后
        # 作为mathml_parent的下一个子元素，而不是nary_elem的子元素
        base_elem = _get_child_by_tag(omml_node, 'e')
        if base_elem is not None:
            for child in base_elem:
                _convert_omml_to_mathml_node(child, mathml_parent)

    elif tag == 'rad':
        # 根式
        deg_elem = _get_child_by_tag(omml_node, 'deg')
        base_elem = _get_child_by_tag(omml_node, 'e')
        
        has_deg = False
        if deg_elem is not None:
            # 检查是否有非空文本
            text = "".join([t for t in deg_elem.itertext()]).strip()
            if text: has_deg = True
            
        if has_deg:
            root = etree.SubElement(mathml_parent, 'mroot')
            if base_elem is not None:
                for child in base_elem: _convert_omml_to_mathml_node(child, root)
            for child in deg_elem: _convert_omml_to_mathml_node(child, root)
        else:
            sqrt = etree.SubElement(mathml_parent, 'msqrt')
            if base_elem is not None:
                for child in base_elem: _convert_omml_to_mathml_node(child, sqrt)
                
    elif tag == 'limLow':
        # 极限：处理lim运算符及其下标条件
        munder = etree.SubElement(mathml_parent, 'munder')
        
        base_elem = _get_child_by_tag(omml_node, 'e') # 通常是 'lim'
        lim_elem = _get_child_by_tag(omml_node, 'lim') # 下面的内容
        
        if base_elem is not None:
            base_row = etree.SubElement(munder, 'mrow')
            for child in base_elem: 
                _convert_omml_to_mathml_node(child, base_row)
        else:
            # 默认添加lim操作符
            mo = etree.SubElement(munder, 'mo')
            mo.text = 'lim'
            
        if lim_elem is not None:
            lim_row = etree.SubElement(munder, 'mrow')
            for child in lim_elem: 
                _convert_omml_to_mathml_node(child, lim_row)
        else:
            # 空的下标
            etree.SubElement(munder, 'mrow')

    elif tag == 'm':
        # OMML矩阵 → MathML mtable
        mtable = etree.SubElement(mathml_parent, 'mtable')
        
        # 遍历所有行
        for mr in omml_node:
            if etree.QName(mr).localname == 'mr':
                mtr = etree.SubElement(mtable, 'mtr')
                
                # 遍历行中的所有单元格
                for e_elem in mr:
                    if etree.QName(e_elem).localname == 'e':
                        mtd = etree.SubElement(mtr, 'mtd')
                        
                        # 转换单元格内容
                        for child in e_elem:
                            _convert_omml_to_mathml_node(child, mtd)
    
    elif tag == 'd':
        # OMML带括号结构 → MathML mfenced
        # 提取括号属性
        dPr = _get_child_by_tag(omml_node, 'dPr')
        open_char = '('
        close_char = ')'
        
        if dPr is not None:
            begChr = _get_child_by_tag(dPr, 'begChr')
            if begChr is not None:
                open_char = begChr.get('{%s}val' % OMML_NS['m'], '(')
            endChr = _get_child_by_tag(dPr, 'endChr')
            if endChr is not None:
                close_char = endChr.get('{%s}val' % OMML_NS['m'], ')')
        
        # 创建mfenced
        mfenced = etree.SubElement(mathml_parent, 'mfenced')
        mfenced.set('open', open_char)
        mfenced.set('close', close_char)
        
        # 转换内容
        e_elem = _get_child_by_tag(omml_node, 'e')
        if e_elem is not None:
            for child in e_elem:
                _convert_omml_to_mathml_node(child, mfenced)

    elif tag == 'e':
        # 容器元素 m:e，直接处理子元素
        for child in omml_node:
            _convert_omml_to_mathml_node(child, mathml_parent)
            
    else:
        # 默认处理所有子节点
        for child in omml_node:
            _convert_omml_to_mathml_node(child, mathml_parent)


def mathml_to_latex(mathml_str: str) -> Optional[str]:
    """将MathML转换为LaTeX格式"""
    try:
        mathml_tree = etree.fromstring(mathml_str.encode('utf-8'))
        result = _mathml_node_to_latex(mathml_tree)
        
        # 后处理：清理矩阵外的重复括号
        if result:
            result = _clean_matrix_brackets(result)
        
        return result
    except Exception as e:
        logger.error(f"MathML转LaTeX失败: {e}")
        return None


def _clean_matrix_brackets(latex_str: str) -> str:
    r"""
    清理矩阵环境外的重复括号
    
    处理情况（由WPS等软件生成的文档可能出现）：
    - [\begin{bmatrix} -> \begin{bmatrix}
    - \end{bmatrix}] -> \end{bmatrix}
    - (\begin{pmatrix} -> \begin{pmatrix}  
    - \end{pmatrix}) -> \end{pmatrix}
    等
    """
    # 清理 bmatrix 外的方括号
    latex_str = re.sub(r'\[\\begin\{bmatrix\}', r'\\begin{bmatrix}', latex_str)
    latex_str = re.sub(r'\\end\{bmatrix\}\]', r'\\end{bmatrix}', latex_str)
    
    # 清理 pmatrix 外的圆括号
    latex_str = re.sub(r'\(\\begin\{pmatrix\}', r'\\begin{pmatrix}', latex_str)
    latex_str = re.sub(r'\\end\{pmatrix\}\)', r'\\end{pmatrix}', latex_str)
    
    # 清理 vmatrix 外的竖线
    latex_str = re.sub(r'\|\\begin\{vmatrix\}', r'\\begin{vmatrix}', latex_str)
    latex_str = re.sub(r'\\end\{vmatrix\}\|', r'\\end{vmatrix}', latex_str)
    
    # 清理 Vmatrix 外的双竖线
    latex_str = re.sub(r'\|\|\\begin\{Vmatrix\}', r'\\begin{Vmatrix}', latex_str)
    latex_str = re.sub(r'\\end\{Vmatrix\}\|\|', r'\\end{Vmatrix}', latex_str)
    
    # 清理 Bmatrix 外的花括号
    latex_str = re.sub(r'\{\\begin\{Bmatrix\}', r'\\begin{Bmatrix}', latex_str)
    latex_str = re.sub(r'\\end\{Bmatrix\}\}', r'\\end{Bmatrix}', latex_str)
    
    return latex_str


def _mathml_node_to_latex(node) -> str:
    """递归将MathML节点转换为LaTeX"""
    tag = etree.QName(node).localname
    
    if tag == 'math':
        # 在math层级处理括号-矩阵模式（逐个处理，不提前返回）
        children = list(node)
        
        # 递归查找 mtable 元素
        def find_mtable(elem):
            """递归查找第一个 mtable 元素"""
            elem_tag = etree.QName(elem).localname
            if elem_tag == 'mtable':
                return elem
            for child in elem:
                result = find_mtable(child)
                if result is not None:
                    return result
            return None
        
        # 逐个处理子元素，检测括号-矩阵模式
        parts = []
        i = 0
        while i < len(children):
            # 检查当前位置是否是"括号-矩阵-括号"模式的开始
            if i + 2 < len(children):
                curr_tag = etree.QName(children[i]).localname
                next_tag = etree.QName(children[i+2]).localname
                
                if curr_tag == 'mo' and next_tag == 'mo':
                    curr_text = (children[i].text or '').strip()
                    next_text = (children[i+2].text or '').strip()
                    
                    # 检查中间元素是否包含矩阵
                    middle_child = children[i+1]
                    matrix_elem = find_mtable(middle_child)
                    
                    if matrix_elem is not None:
                        # 根据括号类型确定矩阵环境
                        # 注意：使用 startswith 处理WPS合并文本节点的情况
                        # WPS可能将 `|` 和后续字符合并为 `|=ad-bc` 等
                        matrix_type = None
                        if curr_text == '[' and (next_text == ']' or next_text.startswith(']')):
                            matrix_type = 'bmatrix'
                            logger.debug(f"math层级检测到 [矩阵] 模式，使用 bmatrix")
                        elif curr_text == '(' and (next_text == ')' or next_text.startswith(')')):
                            matrix_type = 'pmatrix'
                            logger.debug(f"math层级检测到 (矩阵) 模式，使用 pmatrix")
                        elif curr_text == '|' and (next_text == '|' or next_text.startswith('|')):
                            matrix_type = 'vmatrix'
                            logger.debug(f"math层级检测到 |矩阵| 模式，使用 vmatrix")
                        
                        if matrix_type:
                            # 转换矩阵
                            rows = []
                            for mtr in matrix_elem:
                                if etree.QName(mtr).localname == 'mtr':
                                    cells = []
                                    for mtd in mtr:
                                        if etree.QName(mtd).localname == 'mtd':
                                            cell_content = "".join([_mathml_node_to_latex(c) for c in mtd]).strip()
                                            cells.append(cell_content)
                                    if cells:
                                        rows.append(" & ".join(cells))
                            
                            if rows:
                                matrix_content = " \\\\\n".join(rows)
                                result = f"\\begin{{{matrix_type}}}\n{matrix_content}\n\\end{{{matrix_type}}}"
                                parts.append(result)
                                logger.debug(f"math层级成功转换矩阵: {matrix_type}，{len(rows)}行")
                                # 跳过这3个元素（左括号、矩阵内容、右括号）
                                i += 3
                                continue
            
            # 常规处理当前元素
            parts.append(_mathml_node_to_latex(children[i]))
            i += 1
        
        return "".join(parts)
    
    elif tag in ['mi', 'mn', 'mo']:
        text = node.text or ''
        return SYMBOL_MAP.get(text, text)
        
    elif tag == 'mfrac':
        children = list(node)
        if len(children) >= 2:
            num = _mathml_node_to_latex(children[0]).strip()
            den = _mathml_node_to_latex(children[1]).strip()
            return f"\\frac{{{num}}}{{{den}}}"
        return ""
            
    elif tag == 'msup':
        children = list(node)
        if len(children) >= 2:
            base = _mathml_node_to_latex(children[0]).strip()
            sup = _mathml_node_to_latex(children[1]).strip()
            # 如果base是单个字符或操作符，不需要额外的花括号
            if len(base) == 1 or base in NARY_OP_MAP.values():
                return f"{base}^{{{sup}}}"
            return f"{{{base}}}^{{{sup}}}"
        return ""
            
    elif tag == 'msub':
        children = list(node)
        if len(children) >= 2:
            base = _mathml_node_to_latex(children[0]).strip()
            sub = _mathml_node_to_latex(children[1]).strip()
            # 如果base是单个字符或操作符，不需要额外的花括号
            if len(base) == 1 or base in NARY_OP_MAP.values():
                return f"{base}_{{{sub}}}"
            return f"{{{base}}}_{{{sub}}}"
        return ""
            
    elif tag == 'msubsup':
        children = list(node)
        if len(children) >= 3:
            base = _mathml_node_to_latex(children[0]).strip()
            sub = _mathml_node_to_latex(children[1]).strip()
            sup = _mathml_node_to_latex(children[2]).strip()
            
            # 如果base是求和/积分等符号或单个字符，使用 \sum_{...}^{...} 格式
            if base in NARY_OP_MAP.values() or len(base) == 1:
                return f"{base}_{{{sub}}}^{{{sup}}}"
            else:
                return f"{{{base}}}_{{{sub}}}^{{{sup}}}"
        return ""
    
    elif tag == 'munderover':
        children = list(node)
        if len(children) >= 3:
            base = _mathml_node_to_latex(children[0]).strip()
            under = _mathml_node_to_latex(children[1]).strip()
            over = _mathml_node_to_latex(children[2]).strip()
            
            # 对于求和/积分，LaTeX通常用 sub/sup 语法
            return f"{base}_{{{under}}}^{{{over}}}"
        return ""
            
    elif tag == 'munder':
        # 极限等
        children = list(node)
        if len(children) >= 2:
            base = _mathml_node_to_latex(children[0]).strip()
            under = _mathml_node_to_latex(children[1]).strip()
            # 对于lim等操作符，使用下标格式
            return f"{base}_{{{under}}}"
        return ""

    elif tag == 'msqrt':
        content = "".join([_mathml_node_to_latex(child) for child in node]).strip()
        return f"\\sqrt{{{content}}}"
        
    elif tag == 'mroot':
        children = list(node)
        if len(children) >= 2:
            base = _mathml_node_to_latex(children[0]).strip()
            deg = _mathml_node_to_latex(children[1]).strip()
            return f"\\sqrt[{deg}]{{{base}}}"
        return ""

    elif tag == 'mrow':
        """
        处理 MathML 的 mrow 标签（数学行）
        
        关键：正确处理元素间的空格和括号-矩阵模式
        - 运算符（+、-、×、= 等）前后需要空格
        - 变量和数字之间通常不需要空格
        - 函数名和括号之间不需要空格
        - 检测"括号-矩阵-括号"模式，避免重复括号
        - 逐个处理元素，不提前返回
        """
        children = list(node)
        
        # 递归查找 mtable 元素
        def find_mtable(elem):
            """递归查找第一个 mtable 元素"""
            elem_tag = etree.QName(elem).localname
            if elem_tag == 'mtable':
                return elem
            for child in elem:
                result = find_mtable(child)
                if result is not None:
                    return result
            return None
        
        # 逐个处理子元素，检测括号-矩阵模式
        parts = []
        i = 0
        
        while i < len(children):
            # 检查当前位置是否是"括号-矩阵-括号"模式的开始
            if i + 2 < len(children):
                curr_tag = etree.QName(children[i]).localname
                next_tag = etree.QName(children[i+2]).localname
                
                if curr_tag == 'mo' and next_tag == 'mo':
                    curr_text = (children[i].text or '').strip()
                    next_text = (children[i+2].text or '').strip()
                    
                    # 检查中间元素是否包含矩阵
                    middle_child = children[i+1]
                    matrix_elem = find_mtable(middle_child)
                    
                    if matrix_elem is not None:
                        # 根据括号类型确定矩阵环境
                        # 注意：使用 startswith 处理WPS合并文本节点的情况
                        # WPS可能将 `|` 和后续字符合并为 `|=ad-bc` 等
                        matrix_type = None
                        if curr_text == '[' and (next_text == ']' or next_text.startswith(']')):
                            matrix_type = 'bmatrix'
                            logger.debug(f"mrow层级检测到 [矩阵] 模式，使用 bmatrix")
                        elif curr_text == '(' and (next_text == ')' or next_text.startswith(')')):
                            matrix_type = 'pmatrix'
                            logger.debug(f"mrow层级检测到 (矩阵) 模式，使用 pmatrix")
                        elif curr_text == '|' and (next_text == '|' or next_text.startswith('|')):
                            matrix_type = 'vmatrix'
                            logger.debug(f"mrow层级检测到 |矩阵| 模式，使用 vmatrix")
                        
                        if matrix_type:
                            # 转换矩阵
                            rows = []
                            for mtr in matrix_elem:
                                if etree.QName(mtr).localname == 'mtr':
                                    cells = []
                                    for mtd in mtr:
                                        if etree.QName(mtd).localname == 'mtd':
                                            cell_content = "".join([_mathml_node_to_latex(c) for c in mtd]).strip()
                                            cells.append(cell_content)
                                    if cells:
                                        rows.append(" & ".join(cells))
                            
                            if rows:
                                matrix_content = " \\\\\n".join(rows)
                                result = f"\\begin{{{matrix_type}}}\n{matrix_content}\n\\end{{{matrix_type}}}"
                                parts.append(result)
                                logger.debug(f"mrow层级成功转换矩阵: {matrix_type}，{len(rows)}行")
                                # 跳过这3个元素（左括号、矩阵内容、右括号）
                                i += 3
                                continue
            
            # 常规处理当前元素
            child = children[i]
            child_tag = etree.QName(child).localname
            part = _mathml_node_to_latex(child)
            
            if part:
                # 检查是否需要在前面添加空格
                need_space_before = False
                if i > 0 and parts:  # 不是第一个元素
                    prev_part = parts[-1].strip() if parts else ""
                    current_part = part.strip()
                    
                    # 当前元素是运算符，且前面不是空格
                    if child_tag == 'mo':
                        # 二元运算符前需要空格（排除括号、逗号等）
                        if current_part in ['+', '-', '=', '<', '>', '\\times', '\\div', '\\pm', '\\leq', '\\geq', '\\neq', '\\approx']:
                            if prev_part and not prev_part.endswith(' '):
                                need_space_before = True
                    
                    # 前一个元素是运算符，当前元素不是运算符
                    if i > 0:
                        prev_child = children[i-1]
                        prev_tag = etree.QName(prev_child).localname
                        if prev_tag == 'mo' and child_tag != 'mo':
                            prev_text = prev_child.text or ''
                            # 二元运算符后需要空格
                            if prev_text.strip() in ['+', '-', '=', '<', '>', '×', '÷', '±', '≤', '≥', '≠', '≈']:
                                if prev_part and not prev_part.endswith(' '):
                                    need_space_before = True
                
                if need_space_before:
                    parts.append(' ')
                    logger.debug(f"在运算符前后添加空格")
                
                parts.append(part)
            
            i += 1
        
        result = "".join(parts)
        return result
    
    elif tag == 'mtable':
        # MathML矩阵 → LaTeX matrix
        # 默认使用bmatrix（方括号矩阵）
        rows = []
        for mtr in node:
            if etree.QName(mtr).localname == 'mtr':
                cells = []
                for mtd in mtr:
                    if etree.QName(mtd).localname == 'mtd':
                        cell_content = "".join([_mathml_node_to_latex(child) for child in mtd]).strip()
                        cells.append(cell_content)
                if cells:
                    rows.append(" & ".join(cells))
        
        if rows:
            matrix_content = " \\\\\n".join(rows)
            return f"\\begin{{bmatrix}}\n{matrix_content}\n\\end{{bmatrix}}"
        return ""
    
    elif tag == 'mfenced':
        """
        处理 MathML 的 mfenced 标签（带括号的内容）
        
        关键问题：避免在矩阵外添加重复括号
        - bmatrix/pmatrix 等环境已经包含括号
        - 需要递归检测嵌套在 mrow 等容器中的 mtable
        """
        open_char = node.get('open', '(')
        close_char = node.get('close', ')')
        
        logger.debug(f"处理 mfenced 标签，括号类型: {open_char}...{close_char}")
        
        # 递归检测是否包含矩阵表格（不只检查直接子元素）
        def contains_mtable(elem):
            """递归检查元素树中是否包含 mtable"""
            if etree.QName(elem).localname == 'mtable':
                return True
            for child in elem:
                if contains_mtable(child):
                    return True
            return False
        
        # 递归查找 mtable 元素
        def find_mtable(elem):
            """递归查找第一个 mtable 元素"""
            if etree.QName(elem).localname == 'mtable':
                return elem
            for child in elem:
                result = find_mtable(child)
                if result is not None:
                    return result
            return None
        
        has_mtable = contains_mtable(node)
        
        if has_mtable:
            logger.debug(f"mfenced 中检测到矩阵，括号: {open_char}...{close_char}")
            
            # 根据括号类型选择合适的 LaTeX 矩阵环境
            if open_char == '[' and close_char == ']':
                matrix_type = 'bmatrix'  # 方括号矩阵
            elif open_char == '(' and close_char == ')':
                matrix_type = 'pmatrix'  # 圆括号矩阵
            elif open_char == '|' and close_char == '|':
                matrix_type = 'vmatrix'  # 单竖线矩阵（行列式）
            elif open_char in ('||', '‖') and close_char in ('||', '‖'):
                matrix_type = 'Vmatrix'  # 双竖线矩阵（范数）
            elif open_char == '{' and close_char == '}':
                matrix_type = 'Bmatrix'  # 花括号矩阵
            else:
                matrix_type = 'matrix'   # 无括号矩阵
            
            logger.debug(f"选择矩阵类型: {matrix_type}")
            
            # 查找 mtable 并转换
            mtable = find_mtable(node)
            if mtable is not None:
                rows = []
                for mtr in mtable:
                    if etree.QName(mtr).localname == 'mtr':
                        cells = []
                        for mtd in mtr:
                            if etree.QName(mtd).localname == 'mtd':
                                cell_content = "".join([_mathml_node_to_latex(c) for c in mtd]).strip()
                                cells.append(cell_content)
                        if cells:
                            rows.append(" & ".join(cells))
                
                if rows:
                    matrix_content = " \\\\\n".join(rows)
                    result = f"\\begin{{{matrix_type}}}\n{matrix_content}\n\\end{{{matrix_type}}}"
                    logger.debug(f"矩阵转换成功: {matrix_type}，{len(rows)}行")
                    return result
                else:
                    # 即使 rows 为空，也要返回，避免 fallback 添加额外括号
                    logger.warning(f"矩阵行提取失败，返回空矩阵")
                    return f"\\begin{{{matrix_type}}}\n\\end{{{matrix_type}}}"
            else:
                # mtable 为 None，返回空矩阵
                logger.warning(f"未找到 mtable 元素，返回空矩阵")
                return f"\\begin{{{matrix_type}}}\n\\end{{{matrix_type}}}"
        
        # 非矩阵的普通括号内容
        logger.debug("mfenced 不包含矩阵，作为普通括号处理")
        content = "".join([_mathml_node_to_latex(child) for child in node])
        return f"{open_char}{content}{close_char}"
        
    else:
        # 未知标签，尝试处理子节点
        return "".join([_mathml_node_to_latex(child) for child in node])


def parse_latex_from_markdown(md_text: str) -> list:
    """从Markdown文本中解析LaTeX公式"""
    formulas = []
    
    # 匹配块公式 $$...$$
    block_pattern = r'\$\$\s*(.+?)\s*\$\$'
    for match in re.finditer(block_pattern, md_text, re.DOTALL):
        latex = match.group(1).strip()
        # 简单规范化：移除多余的换行符，除非是矩阵环境
        if r'\begin' not in latex:
             latex = latex.replace('\n', ' ')
             
        formulas.append({
            'latex': latex,
            'is_inline': False,
            'start': match.start(),
            'end': match.end()
        })
    
    # 匹配行内公式 $...$
    inline_pattern = r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)'
    for match in re.finditer(inline_pattern, md_text):
        is_in_block = any(f['start'] <= match.start() < f['end'] for f in formulas if not f['is_inline'])
        if not is_in_block:
            formulas.append({
                'latex': match.group(1).strip(),
                'is_inline': True,
                'start': match.start(),
                'end': match.end()
            })
    
    formulas.sort(key=lambda x: x['start'])
    return formulas
