"""
样式注入模块

负责检测并注入缺失的样式到Word文档模板中，包括：
- 标题样式 (heading 1~9) - Word内置样式名，不国际化
- 引用样式 (quote 1~9) - 国际化：中文"引用 1~9"，英文"Quote 1~9"
- 代码样式 (Code Block, Inline Code) - 国际化：中文"代码块/行内代码"
- 公式样式 (Formula Block, Inline Formula) - 国际化：中文"公式块/行内公式"
- 脚注/尾注样式 - Word内置样式名，不国际化
- 列表样式 (List Block) - 国际化：中文"列表块"
- 分隔线样式 (Horizontal Rule 1/2/3) - 国际化：中文"分隔线 1/2/3"

【Word/WPS 样式类型说明】

Word 中存在三种样式类型（在 styles.xml 中由 w:type 属性定义）：

1. 段落样式 (Paragraph Style, w:type="paragraph")
   - 应用到整个段落，如 Heading 1, Normal, Code Block
   - 本模块注入的 Code Block 和 quote 1~9 都是段落样式

2. 字符样式 (Character Style, w:type="character")
   - 应用到选中的文字，如 Inline Code, Strong
   - 本模块注入的 Inline Code 是字符样式

3. 链接样式 (Linked Style)
   - 段落样式 + 关联字符样式，通过 <w:link> 标记关联
   - Word 内置的 Heading 1 等是链接样式
   - 本模块注入的是纯段落/字符样式，不带 <w:link>

【WPS 自动转换机制】

即使我们注入的是纯段落样式（如 Code Block），WPS 打开文档后会自动：
1. 将其转换为链接样式
2. 创建关联字符样式 "Code Block Char"
3. 用户选中部分文字点击 Code Block 时，WPS 应用 "Code Block Char"

【样式检测配合】

由于 WPS 会创建关联字符样式，docx_utils.py 中的 detect_run_style_type()
使用 startswith 匹配，确保能识别：
- "Code Block Char" / "代码块 Char" → 代码样式
- "quote 1 Char" / "引用 1 Char" → 引用样式

【国际化说明】

自定义样式（引用、代码、公式、列表、分隔线）的名称通过 i18n/style_resolver.py 获取，
根据当前语言设置注入对应语言的样式名：
- 中文环境：注入 "代码块"、"引用 1" 等中文样式名
- 英文环境：注入 "Code Block"、"Quote 1" 等英文样式名
"""

import logging
import tempfile
import zipfile
import os
from lxml import etree

# 从 templates 导入标题样式和脚注/尾注样式（Word内置，不国际化）
from .templates import (
    HEADING_STYLE_TEMPLATES,
    NOTE_STYLE_CONFIGS,
    # 国际化样式配置生成函数
    get_localized_style_configs,
)

# 导入 StyleNameResolver 用于获取国际化样式名
from docwen.i18n import style_resolver, t

# 使用docx_utils中的命名空间（统一管理）
from docwen.utils.docx_utils import NAMESPACES

# 配置日志
logger = logging.getLogger(__name__)


# ==============================================================================
# 工具函数
# ==============================================================================

def _get_all_styles_with_type(root):
    """
    获取 styles.xml 中所有样式的 (name, type) 对
    
    Word 允许不同类型的样式同名（如段落样式和字符样式可以都叫 "Code Block"），
    因此检测时必须同时考虑名称和类型。
    
    返回:
        set: {(name_lower, style_type), ...}
              style_type: 'paragraph', 'character', 'table', 'numbering'
    """
    styles = set()
    for style_elem in root.findall('.//w:style', namespaces=NAMESPACES):
        # 获取样式类型
        style_type = style_elem.get(f'{{{NAMESPACES["w"]}}}type', '')
        # 获取样式名称
        name_elem = style_elem.find('w:name', namespaces=NAMESPACES)
        if name_elem is not None:
            name_val = name_elem.get(f'{{{NAMESPACES["w"]}}}val', '')
            if name_val and style_type:
                styles.add((name_val.lower(), style_type))
    return styles


def _find_style_by_name(root, style_name: str):
    """
    通过样式名称查找样式元素（不考虑类型）
    
    参数:
        root: styles.xml 的根元素
        style_name: 要查找的样式名称
    
    返回:
        etree._Element: 找到的样式元素，未找到返回 None
    """
    for style_elem in root.findall('.//w:style', namespaces=NAMESPACES):
        name_elem = style_elem.find('w:name', namespaces=NAMESPACES)
        if name_elem is not None:
            name_val = name_elem.get(f'{{{NAMESPACES["w"]}}}val', '')
            if name_val.lower() == style_name.lower():
                return style_elem
    return None


def _get_all_style_ids(root):
    """获取 styles.xml 中所有样式的 styleId"""
    ids = set()
    for style_elem in root.findall('.//w:style', namespaces=NAMESPACES):
        style_id = style_elem.get(f'{{{NAMESPACES["w"]}}}styleId', '')
        if style_id:
            ids.add(style_id)
    return ids


def _find_base_style_ids(root):
    """
    查找基础样式的实际 styleId
    
    不同模板中基础样式的 styleId 可能不同：
    - 中文版 Word/WPS: Normal 的 styleId 可能是 "a"
    - 英文版 Word: Normal 的 styleId 通常是 "Normal"
    
    返回:
        dict: {
            'Normal': 实际的 Normal styleId,
            'DefaultParagraphFont': 实际的默认字符样式 styleId
        }
    """
    base_ids = {
        'Normal': 'Normal',  # 默认值
        'DefaultParagraphFont': 'DefaultParagraphFont'  # 默认值
    }
    
    for style_elem in root.findall('.//w:style', namespaces=NAMESPACES):
        name_elem = style_elem.find('w:name', namespaces=NAMESPACES)
        if name_elem is None:
            continue
        
        name_val = name_elem.get(f'{{{NAMESPACES["w"]}}}val', '')
        style_id = style_elem.get(f'{{{NAMESPACES["w"]}}}styleId', '')
        
        if not name_val or not style_id:
            continue
        
        # Normal 样式（段落）
        if name_val.lower() == 'normal':
            style_type = style_elem.get(f'{{{NAMESPACES["w"]}}}type', '')
            if style_type == 'paragraph':
                base_ids['Normal'] = style_id
                logger.debug(f"找到 Normal 样式的实际 styleId: {style_id}")
        
        # Default Paragraph Font 样式（字符）
        elif name_val.lower() == 'default paragraph font':
            base_ids['DefaultParagraphFont'] = style_id
            logger.debug(f"找到 DefaultParagraphFont 的实际 styleId: {style_id}")
    
    return base_ids


def _replace_base_style_refs(template: str, base_ids: dict) -> str:
    """
    替换样式模板中的基础样式引用
    
    将模板中硬编码的 "Normal" 和 "DefaultParagraphFont" 替换为实际的 styleId
    
    参数:
        template: 样式 XML 模板字符串
        base_ids: 基础样式 ID 映射字典
    
    返回:
        替换后的模板字符串
    """
    result = template
    
    # 替换 basedOn 中的 Normal 引用
    if base_ids.get('Normal') != 'Normal':
        result = result.replace(
            'w:val="Normal"',
            f'w:val="{base_ids["Normal"]}"'
        )
    
    # 替换 basedOn 中的 DefaultParagraphFont 引用
    if base_ids.get('DefaultParagraphFont') != 'DefaultParagraphFont':
        result = result.replace(
            'w:val="DefaultParagraphFont"',
            f'w:val="{base_ids["DefaultParagraphFont"]}"'
        )
    
    return result


def _get_unique_style_id(existing_ids, preferred_id):
    """
    获取唯一的 styleId
    如果 preferred_id 已存在，则添加数字后缀
    """
    if preferred_id not in existing_ids:
        return preferred_id
    
    # 冲突：尝试 _1, _2, _3...
    counter = 1
    while f"{preferred_id}_{counter}" in existing_ids:
        counter += 1
    
    return f"{preferred_id}_{counter}"


def _inject_style_if_needed(root, existing_styles, existing_ids, target_name, expected_type, default_style_id, style_template, progress_callback=None):
    """
    安全注入样式（考虑样式类型，处理冲突）
    
    Word 允许不同类型的样式同名（段落样式和字符样式可以都叫 "Code Block"），
    因此检测时必须同时考虑名称和类型。如果发现同名但类型不同的冲突样式，
    会删除冲突样式后重新注入正确类型的样式。
    
    参数:
        root: styles.xml 的根元素
        existing_styles: 已存在的样式集合 {(name_lower, type), ...}
        existing_ids: 已存在的 styleId 集合
        target_name: 目标样式名称
        expected_type: 期望的样式类型 ('paragraph' 或 'character')
        default_style_id: 默认 styleId
        style_template: 样式 XML 模板
        progress_callback: 进度回调函数（可选），用于向UI报告状态
        
    返回:
        (injected: bool, final_style_id: str, conflict_resolved: bool)
        - injected: 是否注入了新样式
        - final_style_id: 最终的 styleId（如果注入了）
        - conflict_resolved: 是否解决了类型冲突
    """
    conflict_resolved = False
    
    # 1. 检查是否已存在**同名同类型**的样式
    if (target_name.lower(), expected_type) in existing_styles:
        logger.debug(f"样式 '{target_name}' ({expected_type}) 已存在，跳过注入")
        return False, None, False
    
    # 2. 检查是否存在**同名但类型不同**的冲突样式
    conflicting_style = _find_style_by_name(root, target_name)
    if conflicting_style is not None:
        conflicting_type = conflicting_style.get(f'{{{NAMESPACES["w"]}}}type', '')
        if conflicting_type != expected_type:
            # 发现冲突：同名但类型不同
            conflicting_id = conflicting_style.get(f'{{{NAMESPACES["w"]}}}styleId', '')
            logger.warning(f"发现冲突样式 '{target_name}'（类型={conflicting_type}，期望={expected_type}），将删除并重新注入")
            
            # 通知UI
            if progress_callback:
                progress_callback(t('conversion.progress.fixing_conflict_style', name=target_name, from_type=conflicting_type, to_type=expected_type))
            
            # 删除冲突样式
            root.remove(conflicting_style)
            
            # 从集合中移除旧记录
            existing_styles.discard((target_name.lower(), conflicting_type))
            if conflicting_id:
                existing_ids.discard(conflicting_id)
            
            conflict_resolved = True
    
    # 3. 确定唯一的 styleId
    final_style_id = _get_unique_style_id(existing_ids, default_style_id)
    
    # 4. 如果 styleId 需要修改，替换模板中的 ID
    final_template = style_template
    if final_style_id != default_style_id:
        final_template = style_template.replace(
            f'w:styleId="{default_style_id}"',
            f'w:styleId="{final_style_id}"'
        )
        logger.debug(f"styleId '{default_style_id}' 已被占用，使用 '{final_style_id}'")
    
    # 5. 注入
    try:
        style_elem = etree.fromstring(final_template.encode('utf-8'))
        root.append(style_elem)
        
        # 更新已存在集合
        existing_styles.add((target_name.lower(), expected_type))
        existing_ids.add(final_style_id)
        
        logger.debug(f"注入样式: {target_name} ({expected_type}, styleId={final_style_id})")
        return True, final_style_id, conflict_resolved
    except Exception as e:
        logger.warning(f"注入样式 '{target_name}' 失败: {e}")
        return False, None, conflict_resolved


# ==============================================================================
# 主入口函数
# ==============================================================================

def ensure_styles(template_path, progress_callback=None, cleanup_callback=None):
    """
    确保模板文件包含所有必需的样式（标题、引用、代码等）
    如果缺少样式，自动注入预定义的样式定义
    如果发现同名但类型不同的冲突样式，会删除并重新注入正确类型
    
    参数:
        template_path: 模板文件路径
        progress_callback: 进度回调函数（可选），用于向UI报告状态
        cleanup_callback: 清理回调函数（可选），用于注册临时目录以便后续统一清理
        
    返回:
        处理后的临时文件路径（如果有修改）或原始路径（如果无需修改）
    """
    logger.info("检查模板样式...")
    
    # 解压docx文件读取styles.xml
    try:
        with zipfile.ZipFile(template_path, 'r') as zf:
            styles_xml = zf.read('word/styles.xml')
    except Exception as e:
        logger.warning(f"无法读取模板样式文件: {e}")
        return template_path
    
    # 解析XML
    try:
        root = etree.fromstring(styles_xml)
    except Exception as e:
        logger.warning(f"无法解析样式XML: {e}")
        return template_path
    
    # 获取现有样式信息（同时包含名称和类型）
    existing_styles = _get_all_styles_with_type(root)
    existing_ids = _get_all_style_ids(root)
    
    # 查找基础样式的实际 styleId（Normal 可能是 "a"，DefaultParagraphFont 可能是 "a0"）
    base_ids = _find_base_style_ids(root)
    if base_ids['Normal'] != 'Normal':
        logger.info(f"模板中 Normal 样式的 styleId 为 '{base_ids['Normal']}'，将调整样式引用")
    
    injected_count = 0
    conflict_count = 0
    
    # ==== 获取国际化样式配置 ====
    # 根据当前语言设置获取本地化的样式名
    localized_configs = get_localized_style_configs(style_resolver)
    logger.debug(f"已获取国际化样式配置，当前语言: {style_resolver.i18n.get_current_locale()}")
    
    # ==== 1. 检查并注入标题样式（Heading 1-9，段落样式）====
    # 标题样式使用 Word 内置名称，不国际化
    # 现在使用语言相关的格式配置动态生成模板
    heading_configs = localized_configs.get('heading', {})
    for level in range(1, 10):
        config = heading_configs.get(level)
        if config:
            target_name, style_type, default_style_id, style_template = config
            adjusted_template = _replace_base_style_refs(style_template, base_ids)
            injected, _, conflict = _inject_style_if_needed(
                root, existing_styles, existing_ids,
                target_name, style_type, default_style_id, adjusted_template,
                progress_callback
            )
            if injected:
                injected_count += 1
            if conflict:
                conflict_count += 1
        else:
            # 兜底：使用默认模板
            target_name = f'heading {level}'
            default_style_id = f'Heading{level}'
            style_template = HEADING_STYLE_TEMPLATES.get(level)
            
            if style_template:
                adjusted_template = _replace_base_style_refs(style_template, base_ids)
                injected, _, conflict = _inject_style_if_needed(
                    root, existing_styles, existing_ids,
                    target_name, 'paragraph', default_style_id, adjusted_template,
                    progress_callback
                )
                if injected:
                    injected_count += 1
                if conflict:
                    conflict_count += 1
    
    # ==== 2. 检查并注入引用样式（Quote 1-9，段落样式）====
    # 引用样式使用国际化名称
    quote_configs = localized_configs.get('quote', {})
    for level in range(1, 10):
        config = quote_configs.get(level)
        if config:
            target_name, style_type, default_style_id, style_template = config
            adjusted_template = _replace_base_style_refs(style_template, base_ids)
            injected, _, conflict = _inject_style_if_needed(
                root, existing_styles, existing_ids,
                target_name, style_type, default_style_id, adjusted_template,
                progress_callback
            )
            if injected:
                injected_count += 1
                logger.debug(f"注入引用样式: {target_name}")
            if conflict:
                conflict_count += 1
    
    # ==== 3. 检查并注入代码样式 ====
    # 代码样式使用国际化名称
    code_configs = localized_configs.get('code', [])
    for name, style_type, style_id, template in code_configs:
        adjusted_template = _replace_base_style_refs(template, base_ids)
        injected, _, conflict = _inject_style_if_needed(
            root, existing_styles, existing_ids,
            name, style_type, style_id, adjusted_template,
            progress_callback
        )
        if injected:
            injected_count += 1
            logger.debug(f"注入代码样式: {name}")
        if conflict:
            conflict_count += 1
    
    # ==== 4. 检查并注入脚注/尾注样式 ====
    # 脚注/尾注样式使用 Word 内置名称，不国际化
    for name, style_type, style_id, template in NOTE_STYLE_CONFIGS:
        adjusted_template = _replace_base_style_refs(template, base_ids)
        injected, _, conflict = _inject_style_if_needed(
            root, existing_styles, existing_ids,
            name, style_type, style_id, adjusted_template,
            progress_callback
        )
        if injected:
            injected_count += 1
        if conflict:
            conflict_count += 1
    
    # ==== 5. 检查并注入列表样式（List Block，段落样式）====
    # 列表样式使用国际化名称
    list_config = localized_configs.get('list')
    if list_config:
        target_name, style_type, default_style_id, style_template = list_config
        adjusted_template = _replace_base_style_refs(style_template, base_ids)
        injected, _, conflict = _inject_style_if_needed(
            root, existing_styles, existing_ids,
            target_name, style_type, default_style_id, adjusted_template,
            progress_callback
        )
        if injected:
            injected_count += 1
            logger.debug(f"注入列表样式: {target_name}")
        if conflict:
            conflict_count += 1
    
    # ==== 6. 检查并注入公式样式 ====
    # 公式样式使用国际化名称
    formula_configs = localized_configs.get('formula', [])
    for name, style_type, style_id, template in formula_configs:
        adjusted_template = _replace_base_style_refs(template, base_ids)
        injected, _, conflict = _inject_style_if_needed(
            root, existing_styles, existing_ids,
            name, style_type, style_id, adjusted_template,
            progress_callback
        )
        if injected:
            injected_count += 1
            logger.debug(f"注入公式样式: {name}")
        if conflict:
            conflict_count += 1
    
    # ==== 7. 检查并注入分隔线样式 ====
    # 分隔线样式使用国际化名称
    hr_configs = localized_configs.get('horizontal_rule', [])
    for name, style_type, style_id, template in hr_configs:
        adjusted_template = _replace_base_style_refs(template, base_ids)
        injected, _, conflict = _inject_style_if_needed(
            root, existing_styles, existing_ids,
            name, style_type, style_id, adjusted_template,
            progress_callback
        )
        if injected:
            injected_count += 1
            logger.info(f"注入分隔线样式: {name}")
        if conflict:
            conflict_count += 1
    
    # ==== 8. 检查并注入正文段落样式（Body Paragraph，段落样式）====
    # 正文段落样式使用国际化名称和语言相关格式
    body_paragraph_config = localized_configs.get('body_paragraph')
    if body_paragraph_config:
        target_name, style_type, default_style_id, style_template = body_paragraph_config
        adjusted_template = _replace_base_style_refs(style_template, base_ids)
        injected, _, conflict = _inject_style_if_needed(
            root, existing_styles, existing_ids,
            target_name, style_type, default_style_id, adjusted_template,
            progress_callback
        )
        if injected:
            injected_count += 1
            logger.info(f"注入正文段落样式: {target_name}")
        if conflict:
            conflict_count += 1
    
    # ==== 9. 检查并注入表格样式 ====
    # 表格样式使用国际化名称
    table_configs = localized_configs.get('table', [])
    for name, style_type, style_id, template in table_configs:
        adjusted_template = _replace_base_style_refs(template, base_ids)
        injected, _, conflict = _inject_style_if_needed(
            root, existing_styles, existing_ids,
            name, style_type, style_id, adjusted_template,
            progress_callback
        )
        if injected:
            injected_count += 1
            logger.info(f"注入表格样式: {name}")
        if conflict:
            conflict_count += 1
    
    # ==== 10. 检查并注入网格表样式（根据配置决定是否需要） ====
    # 网格表作为备用表格样式，始终注入
    table_grid_name = style_resolver.get_injection_name("table_grid")
    from .templates import get_localized_table_grid_template
    table_grid_template = get_localized_table_grid_template(table_grid_name)
    adjusted_template = _replace_base_style_refs(table_grid_template, base_ids)
    injected, _, conflict = _inject_style_if_needed(
        root, existing_styles, existing_ids,
        table_grid_name, 'table', 'TableGrid', adjusted_template,
        progress_callback
    )
    if injected:
        injected_count += 1
        logger.info(f"注入网格表样式: {table_grid_name}")
    if conflict:
        conflict_count += 1
    
    # 如果没有注入任何样式，直接返回原路径
    if injected_count == 0:
        logger.info("模板已包含所有必需样式")
        return template_path
    
    # 通知UI注入情况
    if progress_callback and injected_count > 0:
        if conflict_count > 0:
            progress_callback(t('conversion.progress.styles_injected_with_conflicts', count=injected_count, conflicts=conflict_count))
        else:
            progress_callback(t('conversion.progress.styles_injected', count=injected_count))
    
    logger.info(f"共注入 {injected_count} 个缺失样式，修正 {conflict_count} 个冲突")
    
    # 保存修改后的文件到临时位置
    try:
        # 创建临时文件
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, os.path.basename(template_path))
        
        # 如果提供了清理回调，注册临时目录以便后续统一清理
        if cleanup_callback:
            cleanup_callback(temp_dir)
        
        # 复制原始docx并替换styles.xml
        with zipfile.ZipFile(template_path, 'r') as zf_in:
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    if item.filename == 'word/styles.xml':
                        # 写入修改后的styles.xml
                        modified_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
                        zf_out.writestr(item, modified_xml)
                    else:
                        # 复制其他文件
                        zf_out.writestr(item, zf_in.read(item.filename))
        
        logger.info(f"已创建包含完整样式的临时模板: {temp_path}")
        return temp_path
        
    except Exception as e:
        logger.error(f"保存修改后的模板失败: {e}")
        return template_path
