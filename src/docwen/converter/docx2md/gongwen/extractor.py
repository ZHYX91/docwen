"""
公文元素提取模块

负责从文本中提取公文特定元素，如发文字号、签发人等。

主要功能：
- extract_doc_number_and_signers(): 从组合格式中提取发文字号和签发人
- extract_signers_from_text(): 从纯人名文本中提取签发人列表
- extract_doc_number_and_name(): 从组合格式中提取发文字号和人名

使用示例：
    from .extractor import extract_doc_number_and_signers

    text = "国办发〔2024〕1号   签发人：张三"
    doc_num, signers = extract_doc_number_and_signers(text)
    # doc_num = "国办发〔2024〕1号"
    # signers = ["张三"]
"""

import logging
import re

logger = logging.getLogger(__name__)


def extract_doc_number_and_signers(text: str) -> tuple:
    """
    从"发文字号+签发人"组合格式中提取信息

    支持的格式：
    - "XX〔2024〕1号   签发人：张三"
    - "国办发〔2024〕123号  签发人：张三、李四"

    参数:
        text: str - 包含发文字号和签发人的文本

    返回:
        tuple: (发文字号, 签发人列表)
            - 发文字号: str - 提取的发文字号，无法提取则为空字符串
            - 签发人列表: list[str] - 签发人列表，无法提取则为空列表

    示例:
        >>> extract_doc_number_and_signers("国办发〔2024〕1号   签发人：张三")
        ("国办发〔2024〕1号", ["张三"])

        >>> extract_doc_number_and_signers("京政办〔2024〕5号  签发人：李四、王五")
        ("京政办〔2024〕5号", ["李四", "王五"])
    """
    text = text.strip()

    # 分离"签发人："及其之前的内容
    # 使用[\s\u3000\t]+匹配各种空白字符（普通空格、全角空格、制表符）
    match = re.match(r"^(.+?)[\s\u3000\t]+签发人[：:](.+)$", text)
    if match:
        doc_number = match.group(1).strip()
        signer_text = match.group(2).strip()

        # 提取签发人（可能多个，用顿号或空格分隔）
        signers = re.split(r"[、\s\u3000]+", signer_text)
        signers = [s.strip() for s in signers if s.strip()]

        logger.debug(f"提取发文字号和签发人: 发文字号='{doc_number}', 签发人={signers}")
        return doc_number, signers

    logger.debug(f"无法提取发文字号和签发人: '{text}'")
    return "", []


def extract_signers_from_text(text: str) -> list:
    """
    从纯人名文本中提取多个签发人

    支持的格式：
    - "李四、王五" - 顿号分隔
    - "李四 王五" - 空格分隔
    - "张三　李四" - 全角空格分隔

    参数:
        text: str - 包含签发人的文本

    返回:
        list[str]: 签发人列表，只包含符合人名格式（2-4个中文字符）的项

    示例:
        >>> extract_signers_from_text("李四、王五")
        ["李四", "王五"]

        >>> extract_signers_from_text("张三 李四 王五")
        ["张三", "李四", "王五"]

    注意:
        - 只返回符合中文人名格式的项（2-4个中文字符）
        - 不符合格式的项会被过滤并记录警告日志
    """
    text = text.strip()

    # 按顿号或空格分割
    signers = re.split(r"[、\s\u3000]+", text)
    signers = [s.strip() for s in signers if s.strip()]

    # 验证每个都是人名格式（2-4个中文字符）
    valid_signers = []
    for signer in signers:
        if re.match(r"^[\u4e00-\u9fa5]{2,4}$", signer):
            valid_signers.append(signer)
        else:
            logger.warning(f"签发人格式不符合预期: '{signer}'")

    logger.debug(f"提取签发人: {valid_signers}")
    return valid_signers


def extract_doc_number_and_name(text: str) -> tuple:
    """
    从"发文字号+人名"格式中提取信息（无"签发人："前缀）

    支持的格式：
    - "XX〔2024〕2号   李四" - 六角括号
    - "XX[2024]2号   李四" - 方括号
    - "XX(2024)2号   李四" - 圆括号
    - "XX2024-2号   李四" - 连字符

    参数:
        text: str - 包含发文字号和人名的文本

    返回:
        tuple: (发文字号, 签发人)
            - 发文字号: str - 提取的发文字号，无法提取则为空字符串
            - 签发人: str - 提取的签发人，无法提取则为空字符串

    示例:
        >>> extract_doc_number_and_name("国办发〔2024〕2号   李四")
        ("国办发〔2024〕2号", "李四")

        >>> extract_doc_number_and_name("京政[2024]5号  王五")
        ("京政[2024]5号", "王五")

    注意:
        - 人名必须是2-4个中文字符
        - 发文字号和人名之间必须有空白字符分隔
    """
    text = text.strip()

    # 使用正则提取：发文字号 + 空白 + 人名（2-4个中文字符）
    # 需要匹配多种发文字号格式
    patterns = [
        # 六角括号格式：国办发〔2024〕1号
        r"^([\u4e00-\u9fa5A-Za-z0-9]+〔\d{4}〕\s*\d*\s*号)[\s\u3000\t]+([\u4e00-\u9fa5]{2,4})$",
        # 方括号格式：国办发[2024]1号
        r"^([\u4e00-\u9fa5A-Za-z0-9]+\[\d{4}\]\s*\d*\s*号)[\s\u3000\t]+([\u4e00-\u9fa5]{2,4})$",
        # 圆括号格式：国办发(2024)1号
        r"^([\u4e00-\u9fa5A-Za-z0-9]+\(\d{4}\)\s*\d*\s*号)[\s\u3000\t]+([\u4e00-\u9fa5]{2,4})$",
        # 连字符格式：国办发2024-1号
        r"^([\u4e00-\u9fa5A-Za-z0-9]+\d{4}-\s*\d*\s*号)[\s\u3000\t]+([\u4e00-\u9fa5]{2,4})$",
    ]

    for pattern in patterns:
        match = re.match(pattern, text)
        if match:
            doc_number = match.group(1).strip()
            name = match.group(2).strip()
            logger.debug(f"提取发文字号和人名: 发文字号='{doc_number}', 人名='{name}'")
            return doc_number, name

    logger.debug(f"无法提取发文字号和人名: '{text}'")
    return "", ""
