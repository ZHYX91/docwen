"""
策略模块初始化文件

定义了策略注册表和相关的注册、获取函数。
采用基于元数据（Source->Target）的注册机制，支持精确匹配和类别匹配。
"""
import logging
from typing import Dict, Type, Any, Optional, Tuple, Union

logger = logging.getLogger(__name__)

# ==================== 类别常量定义 ====================
CATEGORY_DOCUMENT = 'document'      # docx, doc, rtf, odt, wps
CATEGORY_SPREADSHEET = 'spreadsheet' # xlsx, xls, et, ods, csv
CATEGORY_IMAGE = 'image'            # jpg, png, tiff, etc.
CATEGORY_LAYOUT = 'layout'          # pdf, ofd
CATEGORY_MARKDOWN = 'markdown'      # md, txt (source/target)

# ==================== 策略注册表 ====================
# 转换策略注册表: (source_format, target_format) -> StrategyClass
_conversion_registry: Dict[Tuple[str, str], Type] = {}

# 命名动作注册表: action_name -> StrategyClass
_action_registry: Dict[str, Type] = {}

def register_conversion(source_format: str, target_format: str):
    """
    注册一个特定格式转换的策略。
    
    参数:
        source_format: 源格式 (如 'docx', 'image', 'document')
        target_format: 目标格式 (如 'pdf', 'image')
        
    说明:
        支持注册具体格式 (如 'docx') 或 通用类别 (如 CATEGORY_DOCUMENT)。
        查找时优先精确匹配，其次尝试类别匹配。
    """
    def decorator(cls):
        key = (source_format.lower(), target_format.lower())
        _conversion_registry[key] = cls
        # logger.debug(f"注册转换策略: {source_format} -> {target_format} : {cls.__name__}")
        return cls
    return decorator

def register_action(action_name: str):
    """
    注册一个命名动作 (如 'validate', 'split_pdf').
    
    参数:
        action_name: 动作标识符
    """
    def decorator(cls):
        _action_registry[action_name] = cls
        # logger.debug(f"注册动作策略: {action_name} : {cls.__name__}")
        return cls
    return decorator

# ==================== 辅助函数 ====================

def _get_category(fmt: str) -> Optional[str]:
    """根据文件格式获取所属类别"""
    fmt = fmt.lower()
    if fmt in ['docx', 'doc', 'rtf', 'odt', 'wps']:
        return CATEGORY_DOCUMENT
    if fmt in ['xlsx', 'xls', 'et', 'ods', 'csv']:
        return CATEGORY_SPREADSHEET
    if fmt in ['pdf', 'ofd']:
        return CATEGORY_LAYOUT
    if fmt in ['md', 'markdown', 'txt']:
        return CATEGORY_MARKDOWN
    if fmt in ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'tif', 'webp', 'heic', 'heif']:
        return CATEGORY_IMAGE
    return None

def get_strategy(
    action_type: Optional[str] = None, 
    source_format: Optional[str] = None, 
    target_format: Optional[str] = None
) -> Type:
    """
    智能查找策略类。
    
    查找逻辑 (优先级从高到低):
    1. 命名动作匹配 (如果提供了 action_type)
    2. 精确格式匹配 (source_format -> target_format)
    3. 源文件类别通用匹配 (Category -> target_format) (例如: document -> pdf)
    4. 纯类别匹配 (Category -> Category) (例如: image -> image)
    
    参数:
        action_type: 命名动作 (可选, 如 'validate')
        source_format: 源文件格式 (可选, 如 'docx')
        target_format: 目标文件格式 (可选, 如 'pdf')
        
    返回:
        StrategyClass: 匹配到的策略类
        
    异常:
        ValueError: 如果未找到匹配的策略
    """
    # 1. 尝试查找命名动作
    if action_type and action_type in _action_registry:
        # logger.debug(f"策略查找: 匹配到命名动作 '{action_type}'")
        return _action_registry[action_type]
    
    if source_format and target_format:
        src = source_format.lower()
        tgt = target_format.lower()
        src_cat = _get_category(src)
        tgt_cat = _get_category(tgt)
        
        # 2. 精确格式匹配 (如: docx -> doc)
        if (src, tgt) in _conversion_registry:
            # logger.debug(f"策略查找: 精确匹配 {src} -> {tgt}")
            return _conversion_registry[(src, tgt)]
            
        # 3. 源类别 -> 具体目标 匹配 (如: document -> pdf)
        if src_cat and (src_cat, tgt) in _conversion_registry:
            # logger.debug(f"策略查找: 类别匹配 {src_cat} -> {tgt}")
            return _conversion_registry[(src_cat, tgt)]
            
        # 4. 源类别 -> 目标类别 匹配 (如: image -> image)
        # 注意: 这里假设如果目标是 image，则使用通用图片策略
        if src_cat == CATEGORY_IMAGE and tgt_cat == CATEGORY_IMAGE:
             if (CATEGORY_IMAGE, CATEGORY_IMAGE) in _conversion_registry:
                 # logger.debug(f"策略查找: 通用图片匹配 {CATEGORY_IMAGE} -> {CATEGORY_IMAGE}")
                 return _conversion_registry[(CATEGORY_IMAGE, CATEGORY_IMAGE)]

    # 构建错误信息
    error_msg = "没有找到策略: "
    if action_type:
        error_msg += f"action='{action_type}' "
    if source_format and target_format:
        error_msg += f"conversion='{source_format}->{target_format}'"
    
    logger.error(error_msg)
    raise ValueError(error_msg)

# 导入所有策略模块以触发注册
# 子包形式的策略模块（按文件类型分类）
from . import markdown    # MD → 文档/表格
from . import layout      # 版式文件策略（PDF/OFD/XPS/CAJ）
from . import image       # 图片文件策略
from . import spreadsheet # 表格文件策略
from . import document    # 文档文件策略

# 操作类策略子包（非格式转换）
from . import operations  # 表格汇总、MD序号处理等
