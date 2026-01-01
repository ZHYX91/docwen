"""
样式模板定义模块

本模块包含 Markdown 转 Word 过程中需要注入的所有样式 XML 模板。
这些模板定义了各种 Word 样式的格式属性，用于在模板中缺少对应样式时自动注入。

样式类型说明：
- 段落样式 (w:type="paragraph"): 应用到整个段落，如 Heading 1, Code Block
- 字符样式 (w:type="character"): 应用到选中的文字，如 Inline Code
- 表格样式 (w:type="table"): 应用到表格，如 Three Line Table

模块结构：
- HEADING_STYLE_TEMPLATES: 标题样式 (heading 1~9) - Word内置，不国际化
- 自定义样式（需要国际化）：
  - Quote 1~9: 引用样式
  - Code Block / Inline Code: 代码样式
  - Formula Block / Inline Formula: 公式样式
  - List Block: 列表样式
  - Horizontal Rule 1/2/3: 分隔线样式
  - Table Content / Three Line Table: 表格样式
- NOTE_*_STYLE_TEMPLATE: 脚注/尾注样式 - Word内置，不国际化

国际化说明：
自定义样式名称通过 i18n/style_resolver.py 获取，支持多语言：
- 中文环境：注入 "代码块"、"引用 1" 等中文样式名
- 英文环境：注入 "Code Block"、"Quote 1" 等英文样式名
"""

import logging

# 配置日志
logger = logging.getLogger(__name__)


# ==============================================================================
# 标题样式模板 (heading 1~9)
# 
# 基于公文通用模板的格式规范：
# - Heading 1: 黑体，三号（16pt）
# - Heading 2: 楷体_GB2312，三号
# - Heading 3-9: 仿宋_GB2312，三号
# 所有标题都带有首行缩进（两字符）
# ==============================================================================

HEADING_STYLE_TEMPLATES = {
    1: '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
            <w:ind w:firstLineChars="200" w:firstLine="632"/>
            <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/>
            <w:sz w:val="32"/>
            <w:szCs w:val="32"/>
        </w:rPr>
    </w:style>''',
    
    2: '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="heading 2"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
            <w:ind w:firstLineChars="200" w:firstLine="632"/>
            <w:outlineLvl w:val="1"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:eastAsia="楷体_GB2312" w:hAnsi="Times New Roman"/>
            <w:sz w:val="32"/>
            <w:szCs w:val="32"/>
        </w:rPr>
    </w:style>''',
}

# Heading 3-9 使用相同的仿宋样式模板
_HEADING_3_9_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="Heading{level}">
    <w:name w:val="heading {level}"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:firstLineChars="200" w:firstLine="632"/>
        <w:outlineLvl w:val="{outline_level}"/>
    </w:pPr>
    <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋_GB2312" w:hAnsi="Times New Roman"/>
        <w:sz w:val="32"/>
        <w:szCs w:val="32"/>
    </w:rPr>
</w:style>'''

# 生成 Heading 3-9 的模板
for _level in range(3, 10):
    HEADING_STYLE_TEMPLATES[_level] = _HEADING_3_9_TEMPLATE.format(
        level=_level, 
        outline_level=_level - 1
    )


# ==============================================================================
# 引用样式模板 (quote 1~9)
#
# 引用块样式特点：
# - 每级增加 240 twips 的左缩进
# - 左侧边框颜色逐级加深（从 #CCCCCC 到 #666666）
# - 统一的灰色背景 (#F5F5F5)
# - 灰色文字 (#666666)
# ==============================================================================

def _generate_quote_template(level: int) -> str:
    """
    生成指定级别的引用样式 XML 模板
    
    参数:
        level: 引用级别 (1-9)
        
    返回:
        str: 引用样式的 XML 模板字符串
    """
    # 计算缩进：基础 480 twips + 每级增加 240 twips
    left_indent = 480 + (level - 1) * 240
    right_indent = 480
    
    # 计算边框颜色：从 #CCCCCC (级别1) 逐级加深到 #666666 (级别9)
    color_value = max(0x66, 0xCC - (level - 1) * 0x0B)
    border_color = f"{color_value:02X}{color_value:02X}{color_value:02X}"
    
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="Quote{level}">
    <w:name w:val="Quote {level}"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:left="{left_indent}" w:right="{right_indent}" w:firstLine="0"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>
        <w:pBdr>
            <w:left w:val="single" w:sz="24" w:space="12" w:color="{border_color}"/>
        </w:pBdr>
    </w:pPr>
    <w:rPr>
        <w:color w:val="666666"/>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
    </w:rPr>
</w:style>'''


# 生成 Quote 1-9 的模板
QUOTE_STYLE_TEMPLATES = {}
for _level in range(1, 10):
    QUOTE_STYLE_TEMPLATES[_level] = _generate_quote_template(_level)


# ==============================================================================
# 代码样式模板
#
# Code Block: 段落样式，用于代码块
# - 等宽字体 (Consolas/等线)
# - 灰色背景 (#F5F5F5)
# - 无首行缩进
#
# Inline Code: 字符样式，用于行内代码
# - 等宽字体
# - 浅灰色背景 (#F0F0F0)
# ==============================================================================

CODE_BLOCK_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="CodeBlock">
    <w:name w:val="Code Block"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:firstLine="0"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
        <w:rFonts w:ascii="Consolas" w:eastAsia="等线" w:hAnsi="Consolas" w:cs="Courier New"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
    </w:rPr>
</w:style>'''

INLINE_CODE_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="InlineCode">
    <w:name w:val="Inline Code"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:rPr>
        <w:rFonts w:ascii="Consolas" w:eastAsia="等线" w:hAnsi="Consolas" w:cs="Courier New"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F0F0F0"/>
    </w:rPr>
</w:style>'''


# ==============================================================================
# 公式样式模板
#
# Formula Block: 段落样式，用于公式块
# - 居中对齐
# - 无首行缩进
# - 上下留有间距
#
# Inline Formula: 字符样式，用于行内公式（预留，目前无特殊格式）
# ==============================================================================

FORMULA_BLOCK_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="FormulaBlock">
    <w:name w:val="Formula Block"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:firstLine="0"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:jc w:val="center"/>
    </w:pPr>
</w:style>'''

INLINE_FORMULA_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="InlineFormula">
    <w:name w:val="Inline Formula"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
</w:style>'''


# ==============================================================================
# 脚注/尾注样式模板
#
# 脚注引用 (footnote reference): 字符样式，上标格式
# 脚注文本 (footnote text): 段落样式，小字号，无首行缩进
# 尾注引用 (endnote reference): 字符样式，上标格式
# 尾注文本 (endnote text): 段落样式，小字号，无首行缩进
# ==============================================================================

FOOTNOTE_REF_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
        <w:vertAlign w:val="superscript"/>
    </w:rPr>
</w:style>'''

FOOTNOTE_TEXT_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:pPr>
        <w:ind w:firstLine="0"/>
        <w:snapToGrid w:val="0"/>
    </w:pPr>
    <w:rPr>
        <w:sz w:val="18"/>
        <w:szCs w:val="18"/>
    </w:rPr>
</w:style>'''

ENDNOTE_REF_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="EndnoteReference">
    <w:name w:val="endnote reference"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
        <w:vertAlign w:val="superscript"/>
    </w:rPr>
</w:style>'''

ENDNOTE_TEXT_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="EndnoteText">
    <w:name w:val="endnote text"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="99"/>
    <w:unhideWhenUsed/>
    <w:pPr>
        <w:ind w:firstLine="0"/>
        <w:snapToGrid w:val="0"/>
    </w:pPr>
    <w:rPr>
        <w:sz w:val="18"/>
        <w:szCs w:val="18"/>
    </w:rPr>
</w:style>'''


# ==============================================================================
# 列表样式模板
#
# List Block: 段落样式，用于列表项
# - 左缩进 720 twips
# - 无首行缩进
# - 使用上下文间距
# ==============================================================================

LIST_BLOCK_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="ListBlock">
    <w:name w:val="List Block"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="34"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:left="720" w:firstLine="0"/>
        <w:contextualSpacing/>
    </w:pPr>
</w:style>'''


# ==============================================================================
# 分隔线样式模板 (Horizontal Rule 1/2/3)
#
# 分隔线通过空段落 + 底部边框实现视觉分隔效果
# - Horizontal Rule 1: 细实线 (sz="4", 约 0.5pt)
# - Horizontal Rule 2: 中等实线 (sz="8", 约 1pt)
# - Horizontal Rule 3: 粗实线 (sz="12", 约 1.5pt)
# ==============================================================================

HORIZONTAL_RULE_1_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="HorizontalRule1">
    <w:name w:val="Horizontal Rule 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="99"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>
        <w:ind w:firstLine="0"/>
        <w:pBdr>
            <w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>
        </w:pBdr>
    </w:pPr>
</w:style>'''

HORIZONTAL_RULE_2_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="HorizontalRule2">
    <w:name w:val="Horizontal Rule 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="99"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>
        <w:ind w:firstLine="0"/>
        <w:pBdr>
            <w:bottom w:val="single" w:sz="8" w:space="1" w:color="auto"/>
        </w:pBdr>
    </w:pPr>
</w:style>'''

HORIZONTAL_RULE_3_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="HorizontalRule3">
    <w:name w:val="Horizontal Rule 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="99"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>
        <w:ind w:firstLine="0"/>
        <w:pBdr>
            <w:bottom w:val="single" w:sz="12" w:space="1" w:color="auto"/>
        </w:pBdr>
    </w:pPr>
</w:style>'''


# ==============================================================================
# 表格样式模板
#
# Three Line Table: 三线表样式（学术论文常用）
# - 上下粗线 (sz="12", 约 1.5pt)
# - 表头分隔细线 (sz="4", 约 0.5pt)
# 注意：WPS 对 tblStylePr 条件格式支持不完整，表头底线需在表格生成时显式设置
#
# Table Content: 表格内容段落样式
# - 五号字 (10.5pt, sz="21")
# - 居中对齐
# - 无首行缩进
# ==============================================================================

THREE_LINE_TABLE_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="table" w:styleId="ThreeLineTable">
    <w:name w:val="Three Line Table"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="59"/>
    <w:qFormat/>
    <w:tblPr>
        <w:tblBorders>
            <w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/>
        </w:tblBorders>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow">
        <w:tcPr>
            <w:tcBorders>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            </w:tcBorders>
        </w:tcPr>
    </w:tblStylePr>
</w:style>'''

TABLE_CONTENT_STYLE_TEMPLATE = '''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="TableContent">
    <w:name w:val="Table Content"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="39"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:firstLine="0"/>
        <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
        <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
    </w:rPr>
</w:style>'''


# ==============================================================================
# 样式配置表
#
# 为了简化 style_injector.py 中的注入逻辑，这里提供统一的配置表
# 每个配置项包含：(样式名称, 样式类型, 默认 styleId, 模板)
# ==============================================================================

# 代码样式配置
CODE_STYLE_CONFIGS = [
    ('Code Block', 'paragraph', 'CodeBlock', CODE_BLOCK_STYLE_TEMPLATE),
    ('Inline Code', 'character', 'InlineCode', INLINE_CODE_STYLE_TEMPLATE),
]

# 公式样式配置
FORMULA_STYLE_CONFIGS = [
    ('Formula Block', 'paragraph', 'FormulaBlock', FORMULA_BLOCK_STYLE_TEMPLATE),
    ('Inline Formula', 'character', 'InlineFormula', INLINE_FORMULA_STYLE_TEMPLATE),
]

# 脚注/尾注样式配置
NOTE_STYLE_CONFIGS = [
    ('footnote reference', 'character', 'FootnoteReference', FOOTNOTE_REF_STYLE_TEMPLATE),
    ('footnote text', 'paragraph', 'FootnoteText', FOOTNOTE_TEXT_STYLE_TEMPLATE),
    ('endnote reference', 'character', 'EndnoteReference', ENDNOTE_REF_STYLE_TEMPLATE),
    ('endnote text', 'paragraph', 'EndnoteText', ENDNOTE_TEXT_STYLE_TEMPLATE),
]

# 分隔线样式配置
HORIZONTAL_RULE_STYLE_CONFIGS = [
    ('Horizontal Rule 1', 'paragraph', 'HorizontalRule1', HORIZONTAL_RULE_1_STYLE_TEMPLATE),
    ('Horizontal Rule 2', 'paragraph', 'HorizontalRule2', HORIZONTAL_RULE_2_STYLE_TEMPLATE),
    ('Horizontal Rule 3', 'paragraph', 'HorizontalRule3', HORIZONTAL_RULE_3_STYLE_TEMPLATE),
]


# ==============================================================================
# 国际化样式模板生成函数
#
# 这些函数根据当前语言设置动态生成带有本地化样式名的 XML 模板。
# 样式名通过 i18n/style_resolver.py 获取，支持多语言：
# - 中文环境：使用 "代码块"、"引用 1" 等中文样式名
# - 英文环境：使用 "Code Block"、"Quote 1" 等英文样式名
# ==============================================================================

def get_localized_quote_template(level: int, style_name: str) -> str:
    """
    生成带有国际化样式名的引用样式 XML 模板
    
    参数:
        level: 引用级别 (1-9)
        style_name: 国际化的样式名（如 "引用 1" 或 "Quote 1"）
        
    返回:
        str: 引用样式的 XML 模板字符串
    """
    # 计算缩进：基础 480 twips + 每级增加 240 twips
    left_indent = 480 + (level - 1) * 240
    right_indent = 480
    
    # 计算边框颜色：从 #CCCCCC (级别1) 逐级加深到 #666666 (级别9)
    color_value = max(0x66, 0xCC - (level - 1) * 0x0B)
    border_color = f"{color_value:02X}{color_value:02X}{color_value:02X}"
    
    # styleId 保持英文（用于内部引用）
    style_id = f"Quote{level}"
    
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="{style_id}">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:left="{left_indent}" w:right="{right_indent}" w:firstLine="0"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>
        <w:pBdr>
            <w:left w:val="single" w:sz="24" w:space="12" w:color="{border_color}"/>
        </w:pBdr>
    </w:pPr>
    <w:rPr>
        <w:color w:val="666666"/>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
    </w:rPr>
</w:style>'''


def get_localized_code_block_template(style_name: str) -> str:
    """
    生成带有国际化样式名的代码块样式 XML 模板
    
    参数:
        style_name: 国际化的样式名（如 "代码块" 或 "Code Block"）
        
    返回:
        str: 代码块样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="CodeBlock">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:firstLine="0"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
        <w:rFonts w:ascii="Consolas" w:eastAsia="等线" w:hAnsi="Consolas" w:cs="Courier New"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
    </w:rPr>
</w:style>'''


def get_localized_inline_code_template(style_name: str) -> str:
    """
    生成带有国际化样式名的行内代码样式 XML 模板
    
    参数:
        style_name: 国际化的样式名（如 "行内代码" 或 "Inline Code"）
        
    返回:
        str: 行内代码样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="InlineCode">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:rPr>
        <w:rFonts w:ascii="Consolas" w:eastAsia="等线" w:hAnsi="Consolas" w:cs="Courier New"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F0F0F0"/>
    </w:rPr>
</w:style>'''


def get_localized_formula_block_template(style_name: str) -> str:
    """
    生成带有国际化样式名的公式块样式 XML 模板
    
    参数:
        style_name: 国际化的样式名（如 "公式块" 或 "Formula Block"）
        
    返回:
        str: 公式块样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="FormulaBlock">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:firstLine="0"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:jc w:val="center"/>
    </w:pPr>
</w:style>'''


def get_localized_inline_formula_template(style_name: str) -> str:
    """
    生成带有国际化样式名的行内公式样式 XML 模板
    
    参数:
        style_name: 国际化的样式名（如 "行内公式" 或 "Inline Formula"）
        
    返回:
        str: 行内公式样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="InlineFormula">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
</w:style>'''


def get_localized_list_block_template(style_name: str) -> str:
    """
    生成带有国际化样式名的列表块样式 XML 模板
    
    参数:
        style_name: 国际化的样式名（如 "列表块" 或 "List Block"）
        
    返回:
        str: 列表块样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="ListBlock">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="34"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:left="720" w:firstLine="0"/>
        <w:contextualSpacing/>
    </w:pPr>
</w:style>'''


def get_localized_horizontal_rule_template(level: int, style_name: str) -> str:
    """
    生成带有国际化样式名的分隔线样式 XML 模板
    
    参数:
        level: 分隔线级别 (1-3)
        style_name: 国际化的样式名（如 "分隔线 1" 或 "Horizontal Rule 1"）
        
    返回:
        str: 分隔线样式的 XML 模板字符串
    """
    # 分隔线粗细：1->4, 2->8, 3->12
    line_size = 4 * level
    style_id = f"HorizontalRule{level}"
    
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="{style_id}">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="99"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>
        <w:ind w:firstLine="0"/>
        <w:pBdr>
            <w:bottom w:val="single" w:sz="{line_size}" w:space="1" w:color="auto"/>
        </w:pBdr>
    </w:pPr>
</w:style>'''


def get_localized_table_content_template(style_name: str) -> str:
    """
    生成带有国际化样式名的表格内容样式 XML 模板
    
    参数:
        style_name: 国际化的样式名（如 "表格内容" 或 "Table Content"）
        
    返回:
        str: 表格内容样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="paragraph" w:styleId="TableContent">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="39"/>
    <w:qFormat/>
    <w:pPr>
        <w:ind w:firstLine="0"/>
        <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
        <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
        <w:sz w:val="21"/>
        <w:szCs w:val="21"/>
    </w:rPr>
</w:style>'''


def get_localized_three_line_table_template(style_name: str) -> str:
    """
    生成带有国际化样式名的三线表样式 XML 模板
    
    参数:
        style_name: 国际化的样式名（如 "三线表" 或 "Three Line Table"）
        
    返回:
        str: 三线表样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="table" w:styleId="ThreeLineTable">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="59"/>
    <w:qFormat/>
    <w:tblPr>
        <w:tblBorders>
            <w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/>
        </w:tblBorders>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow">
        <w:tcPr>
            <w:tcBorders>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            </w:tcBorders>
        </w:tcPr>
    </w:tblStylePr>
</w:style>'''


def get_localized_table_grid_template(style_name: str) -> str:
    """
    生成带有国际化样式名的网格表样式 XML 模板
    
    网格表是最常用的办公文档表格样式：
    - 四周边框 + 内部网格线
    - 所有边框等粗（0.5pt）
    - Word/WPS 内置 Table Grid 样式的等效实现
    
    参数:
        style_name: 国际化的样式名（如 "网格表" 或 "Table Grid"）
        
    返回:
        str: 网格表样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="table" w:styleId="TableGrid">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="39"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:tblPr>
        <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>
    </w:tblPr>
</w:style>'''


def get_custom_table_grid_template(style_name: str) -> str:
    """
    生成用户自定义表格样式的 XML 模板（网格表格式）
    
    当用户输入自定义样式名但模板中不存在时，使用此模板创建一个网格表格式的样式。
    使用 "UserCustomTable" 作为 styleId 以避免与内置 TableGrid 冲突。
    
    参数:
        style_name: 用户自定义的样式名称
        
    返回:
        str: 自定义表格样式的 XML 模板字符串
    """
    return f'''<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="table" w:styleId="UserCustomTable">
    <w:name w:val="{style_name}"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="99"/>
    <w:qFormat/>
    <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:tblPr>
        <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>
    </w:tblPr>
</w:style>'''


# ==============================================================================
# 国际化样式配置生成函数
#
# 根据 StyleNameResolver 提供的国际化样式名，动态生成样式配置表。
# 用于替换静态的 CODE_STYLE_CONFIGS、FORMULA_STYLE_CONFIGS 等配置。
# ==============================================================================

def get_localized_style_configs(style_resolver):
    """
    获取国际化的样式配置表
    
    根据当前语言设置，返回带有本地化样式名的配置表。
    用于替换静态配置（CODE_STYLE_CONFIGS 等）。
    
    参数:
        style_resolver: StyleNameResolver 实例
        
    返回:
        dict: 包含各类样式配置的字典
            - 'code': 代码样式配置列表
            - 'formula': 公式样式配置列表
            - 'horizontal_rule': 分隔线样式配置列表
            - 'list': 列表样式配置
            - 'table': 表格样式配置列表
            - 'quote': 引用样式配置字典 {level: (name, type, id, template)}
    """
    configs = {}
    
    # 代码样式配置
    code_block_name = style_resolver.get_injection_name("code_block")
    inline_code_name = style_resolver.get_injection_name("inline_code")
    configs['code'] = [
        (code_block_name, 'paragraph', 'CodeBlock', get_localized_code_block_template(code_block_name)),
        (inline_code_name, 'character', 'InlineCode', get_localized_inline_code_template(inline_code_name)),
    ]
    
    # 公式样式配置
    formula_block_name = style_resolver.get_injection_name("formula_block")
    inline_formula_name = style_resolver.get_injection_name("inline_formula")
    configs['formula'] = [
        (formula_block_name, 'paragraph', 'FormulaBlock', get_localized_formula_block_template(formula_block_name)),
        (inline_formula_name, 'character', 'InlineFormula', get_localized_inline_formula_template(inline_formula_name)),
    ]
    
    # 分隔线样式配置
    configs['horizontal_rule'] = []
    for level in range(1, 4):
        hr_name = style_resolver.get_injection_name(f"horizontal_rule_{level}")
        configs['horizontal_rule'].append(
            (hr_name, 'paragraph', f'HorizontalRule{level}', get_localized_horizontal_rule_template(level, hr_name))
        )
    
    # 列表样式配置
    list_block_name = style_resolver.get_injection_name("list_block")
    configs['list'] = (list_block_name, 'paragraph', 'ListBlock', get_localized_list_block_template(list_block_name))
    
    # 表格样式配置
    table_content_name = style_resolver.get_injection_name("table_content")
    three_line_table_name = style_resolver.get_injection_name("three_line_table")
    configs['table'] = [
        (table_content_name, 'paragraph', 'TableContent', get_localized_table_content_template(table_content_name)),
        (three_line_table_name, 'table', 'ThreeLineTable', get_localized_three_line_table_template(three_line_table_name)),
    ]
    
    # 引用样式配置
    configs['quote'] = {}
    for level in range(1, 10):
        quote_name = style_resolver.get_injection_name(f"quote_{level}")
        configs['quote'][level] = (
            quote_name, 'paragraph', f'Quote{level}', 
            get_localized_quote_template(level, quote_name)
        )
    
    return configs
