"""
YAML处理器模块
处理MD文档的YAML内容
"""

import yaml
import logging
from gongwen_converter.utils.validation_utils import is_value_empty
from gongwen_converter.utils.text_utils import clean_text_in_data, clean_text

# 配置日志
logger = logging.getLogger(__name__)


def parse_yaml(yaml_content: str) -> dict:
    """
    解析YAML内容，并进行通用清理
    
    参数:
        yaml_content: YAML格式的字符串内容
        
    返回:
        dict: 处理后的YAML数据字典
    """
    logger.info("开始解析YAML内容...")
    logger.debug(f"YAML原始内容: {yaml_content[:100]}...")
    
    if not yaml_content:
        logger.warning("YAML内容为空")
        return {}
    
    try:
        # 1. 解析YAML原始内容
        data = yaml.safe_load(yaml_content) or {}
        logger.info(f"成功解析YAML，包含 {len(data)} 个字段")
        
        # 注意：不在这里清理文本！
        # - clean_text 不处理链接（已移除链接处理功能）
        # - 链接处理由 md2docx/xlsx 中的 process_yaml_links() 负责
        # - HTML/零宽字符清理不需要在这里做（不影响YAML数据）

        return data
        
    except Exception as e:
        logger.error(f"YAML解析错误: {str(e)}", exc_info=True)
        return {}

def process_list_field(value) -> str:
    """
    通用字段值处理
    如果是空、None等，返回空字符串
    如果是数字或字符串，直接返回字符串
    如果是列表，将列表的每个元素去除链接后，用顿号拼接，返回字符串
    如果是字符串形式的列表（如"['a','b']"），会尝试解析为列表后处理
    
    参数:
        value: 原始值（可能是空值、数字、字符串、列表）
        
    返回:
        str: 处理后的字符串
    """
    logger.debug(f"处理字段值: {value} (类型: {type(value).__name__})")
    
    # 空值处理
    if is_value_empty(value):
        logger.debug("空值，返回空字符串")
        return ""
    
    # 处理字符串形式的列表（如"['张三','李四']"）
    if isinstance(value, str) and value.startswith('[') and value.endswith(']'):
        try:
            import ast
            parsed_value = ast.literal_eval(value)
            if isinstance(parsed_value, list):
                logger.debug(f"检测到字符串形式的列表，解析为: {parsed_value}")
                # 递归调用自己处理解析后的列表
                return process_list_field(parsed_value)
        except (ValueError, SyntaxError) as e:
            logger.debug(f"字符串解析为列表失败: {value}, 错误: {str(e)}，按普通字符串处理")
            # 解析失败，继续按普通字符串处理
    
    # 数字类型处理
    if isinstance(value, (int, float)):
        logger.debug(f"数字类型，转换为字符串: {str(value)}")
        return str(value)
    
    # 字符串类型处理
    if isinstance(value, str):
        logger.debug(f"字符串类型，直接返回: {value}")
        return value
    
    # 列表类型处理
    if isinstance(value, (list, tuple)):
        logger.debug(f"列表类型，处理 {len(value)} 个元素")
        processed_items = []
        
        for item in value:
            try:
                # 跳过空值
                if is_value_empty(item):
                    continue
                
                # 转换为字符串并清理链接
                item_str = str(item).strip()
                cleaned_item = clean_text(item_str)
                
                if cleaned_item:
                    processed_items.append(cleaned_item)
                    logger.debug(f"添加处理项: {item_str} -> {cleaned_item}")
                    
            except Exception as e:
                logger.error(f"处理列表项失败: {item}, 错误: {str(e)}")
                # 出错时使用原始字符串
                processed_items.append(str(item).strip())
        
        # 用顿号连接非空项
        if processed_items:
            result = "、".join(processed_items)
            logger.debug(f"列表处理结果: {result}")
            return result
        else:
            logger.debug("列表无有效项，返回空字符串")
            return ""
    
    # 其他类型转换为字符串
    logger.debug(f"其他类型 {type(value).__name__}，转换为字符串: {str(value)}")
    return str(value)

# 模块测试
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(level=logging.DEBUG)
    logger.info("YAML处理器模块测试")
    
    # 测试数据
    test_yaml = """
标题: 关于[[2023年]]重点工作安排
负责人:
  - "[[张三]]"
  - "[[李四]]"
金额: 1234.56
"""
    
    # 解析并处理YAML
    logger.info("测试YAML解析和处理")
    parsed_data = parse_yaml(test_yaml)
    
    # 打印处理结果
    print("\n--- 解析后原始数据 ---")
    for key, value in parsed_data.items():
        print(f"{key}: {value} (类型: {type(value).__name__})")

    # 测试 process_list_field
    logger.info("\n--- 测试 process_list_field ---")
    list_value = parsed_data.get("负责人")
    processed_list = process_list_field(list_value)
    print(f"原始列表: {list_value}")
    print(f"处理结果: {processed_list}")
    
    logger.info("\n模块测试完成!")
