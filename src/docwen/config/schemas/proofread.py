"""
校对配置模块

对应配置文件：
    - proofread_config.toml: 校对主配置（开关和跳过规则）
    - proofread_symbols.toml: 符号配对和映射表
    - proofread_typos.toml: 错别字映射表
    - proofread_sensitive.toml: 敏感词映射表

详细说明：
    包含校对功能的默认配置和访问方法。
    校对配置控制错别字、符号配对、符号校正、敏感词等校对引擎的行为。

包含：
    - DEFAULT_PROOFREAD_CONFIG: 校对主配置
    - DEFAULT_PROOFREAD_SYMBOLS: 符号校对配置
    - DEFAULT_PROOFREAD_TYPOS: 错别字配置
    - DEFAULT_PROOFREAD_SENSITIVE: 敏感词配置
    - ProofreadConfigMixin: 校对配置获取方法

依赖：
    - safe_logger: 安全日志记录（用于配置访问时的调试日志）

使用方式：
    # 通过 ConfigManager 访问（推荐）
    from docwen.config import config_manager
    
    typos = config_manager.get_typos()
    symbol_pairs = config_manager.get_symbol_pairs()
    
    # 直接导入默认配置
    from docwen.config.schemas.proofread import DEFAULT_PROOFREAD_CONFIG
"""

from typing import Dict, Any, List, Tuple

from ..safe_logger import safe_log


# ==============================================================================
#                              默认配置
# ==============================================================================

# 默认校对主配置（开关和跳过规则）
DEFAULT_PROOFREAD_CONFIG = {
    "proofread_config": {
        "engine": {
            "enable_typos_rule": True,          # 启用错别字校对
            "enable_symbol_pairing": True,       # 启用符号配对检查
            "enable_symbol_correction": True,    # 启用符号校正
            "enable_sensitive_word": True,       # 启用敏感词检测
            "skip_code_blocks": True,            # 跳过代码块
            "skip_quote_blocks": False,          # 跳过引用块
        }
    }
}

# 默认符号校对配置（符号配对和映射表）
DEFAULT_PROOFREAD_SYMBOLS = {
    "proofread_symbols": {
        "symbol_pairing": {
            "pairs": []  # 符号配对列表，如 [["（", "）"], ["《", "》"]]
        },
        "symbol_map": {},  # 符号映射表，如 {"1": ["'", "'"], "2": [""", """]}
    }
}

# 默认错别字配置（错别字映射表）
DEFAULT_PROOFREAD_TYPOS = {
    "proofread_typos": {
        "typos": {}  # 错别字映射表，如 {"错字": ["正确写法1", "正确写法2"]}
    }
}

# 默认敏感词配置（敏感词映射表）
DEFAULT_PROOFREAD_SENSITIVE = {
    "proofread_sensitive": {
        "sensitive_words": {}  # 敏感词映射表，如 {"敏感词": ["替换词1", "替换词2"]}
    }
}

# 配置文件名
CONFIG_FILES = {
    "proofread_config": "proofread_config.toml",
    "proofread_symbols": "proofread_symbols.toml",
    "proofread_typos": "proofread_typos.toml",
    "proofread_sensitive": "proofread_sensitive.toml",
}


# ==============================================================================
#                              Mixin 类
# ==============================================================================

class ProofreadConfigMixin:
    """
    校对配置获取方法 Mixin
    
    提供校对相关配置的访问方法，包括错别字、符号配对、符号校正、敏感词等。
    
    注意：
        此类设计为 Mixin，需要与 ConfigManager 一起使用。
        假定宿主类具有 _configs 属性（配置字典）。
    
    配置结构：
        proofread_config:
            engine: 校对引擎开关和跳过规则
        proofread_symbols:
            symbol_pairing: 符号配对配置
            symbol_map: 符号映射配置
        proofread_typos:
            typos: 错别字映射
        proofread_sensitive:
            sensitive_words: 敏感词映射
    """
    
    # 类型提示：声明 _configs 属性（由 ConfigManager 提供）
    _configs: Dict[str, Dict[str, Any]]
    
    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------
    
    def get_proofread_config_block(self) -> Dict[str, Any]:
        """
        获取校对主配置块
        
        返回：
            Dict[str, Any]: 校对主配置字典，包含 engine 子表
        """
        return self._configs.get("proofread_config", {})
    
    def get_proofread_symbols_block(self) -> Dict[str, Any]:
        """
        获取符号校对配置块
        
        返回：
            Dict[str, Any]: 符号校对配置字典，包含 symbol_pairing 和 symbol_map 子表
        """
        return self._configs.get("proofread_symbols", {})

    def get_proofread_typos_block(self) -> Dict[str, Any]:
        """
        获取错别字配置块
        
        返回：
            Dict[str, Any]: 错别字配置字典，包含 typos 子表
        """
        return self._configs.get("proofread_typos", {})

    def get_proofread_sensitive_block(self) -> Dict[str, Any]:
        """
        获取敏感词配置块
        
        返回：
            Dict[str, Any]: 敏感词配置字典，包含 sensitive_words 子表
        """
        return self._configs.get("proofread_sensitive", {})
    
    # --------------------------------------------------------------------------
    # 第二层：子表
    # --------------------------------------------------------------------------
    
    def get_proofread_engine_config(self) -> Dict[str, Any]:
        """
        获取校对引擎配置（开关和跳过规则）
        
        返回：
            Dict[str, Any]: 引擎配置字典
        """
        return self.get_proofread_config_block().get("engine", {})

    def get_symbol_pairing_config(self) -> Dict[str, Any]:
        """
        获取符号配对子表
        
        返回：
            Dict[str, Any]: 符号配对配置字典
        """
        return self.get_proofread_symbols_block().get("symbol_pairing", {})
    
    def get_symbol_map_config(self) -> Dict[str, Any]:
        """
        获取符号映射子表
        
        返回：
            Dict[str, Any]: 符号映射配置字典
        """
        return self.get_proofread_symbols_block().get("symbol_map", {})
    
    def get_typos_config(self) -> Dict[str, Any]:
        """
        获取错别字子表
        
        返回：
            Dict[str, Any]: 错别字配置字典
        """
        return self.get_proofread_typos_block().get("typos", {})

    def get_sensitive_words_config(self) -> Dict[str, Any]:
        """
        获取敏感词子表
        
        返回：
            Dict[str, Any]: 敏感词配置字典
        """
        return self.get_proofread_sensitive_block().get("sensitive_words", {})
    
    # --------------------------------------------------------------------------
    # 第三层：校对跳过规则
    # --------------------------------------------------------------------------
    
    def is_skip_code_blocks_enabled(self) -> bool:
        """
        是否跳过代码块段落
        
        返回：
            bool: 是否跳过代码块
        """
        config = self.get_proofread_engine_config()
        enabled = config.get("skip_code_blocks", True)
        safe_log.debug("校对跳过代码块: %s", enabled)
        return enabled
    
    def is_skip_quote_blocks_enabled(self) -> bool:
        """
        是否跳过引用段落
        
        返回：
            bool: 是否跳过引用块
        """
        config = self.get_proofread_engine_config()
        enabled = config.get("skip_quote_blocks", False)
        safe_log.debug("校对跳过引用块: %s", enabled)
        return enabled
    
    # --------------------------------------------------------------------------
    # 第三层：具体配置值
    # --------------------------------------------------------------------------
    
    def get_typos(self) -> Dict[str, list]:
        """
        获取错别字映射表（带默认值）
        
        返回：
            Dict[str, list]: 错别字映射表，键为错字，值为正确写法列表
        """
        typos = self.get_typos_config()
        return typos if isinstance(typos, dict) else {}
    
    def get_sensitive_words(self) -> Dict[str, list]:
        """
        获取敏感词映射表（带默认值）
        
        返回：
            Dict[str, list]: 敏感词映射表，键为敏感词，值为替换词列表
        """
        words = self.get_sensitive_words_config()
        return words if isinstance(words, dict) else {}
    
    def get_symbol_map(self) -> Dict[str, list]:
        """
        获取标点符号映射（带默认值，键统一为字符串）
        
        注意：
            统一将所有键转为字符串，兼容整数键和字符串键。
            这确保 GUI 编辑后的字符串键与配置文件中的数字键行为一致。
        
        返回：
            Dict[str, list]: 符号映射表，键为映射ID，值为符号列表
        """
        symbol_map = self.get_symbol_map_config()
        if isinstance(symbol_map, dict):
            return {str(k): v for k, v in symbol_map.items()}
        return {}
    
    def get_symbol_pairs(self) -> List[Tuple[str, str]]:
        """
        获取符号对列表（带默认值）
        
        返回：
            List[Tuple[str, str]]: 符号对列表，每个元素为 (左符号, 右符号)
        """
        pairing_config = self.get_symbol_pairing_config()
        pairs = pairing_config.get("pairs", [])
        
        # 转换为元组列表
        if isinstance(pairs, list):
            return [tuple(pair) for pair in pairs if isinstance(pair, list) and len(pair) == 2]
        return []
