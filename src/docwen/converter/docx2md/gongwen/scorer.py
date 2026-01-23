"""
文档元素验证器模块
实现公文元素的识别评分和合规性验证
"""

import re
import logging
from abc import ABC
from docwen.utils import docx_utils

# 配置日志
logger = logging.getLogger(__name__)

class BaseValidator(ABC):
    """验证器抽象基类，提供统一接口"""
    
    def __init__(self):
        """初始化基础验证器"""
        logger.debug(f"初始化验证器: {self.__class__.__name__}")
       
    def create_result(self, is_valid: bool, message: str = "", details: dict = None) -> dict:
        """
        创建标准化的验证结果
        
        参数:
            is_valid: 验证是否通过
            message: 验证结果消息
            details: 详细验证数据
            
        返回:
            dict: 标准验证结果 {
                "valid": bool,
                "message": str,
                "details": dict
            }
        """
        result = {
            "valid": is_valid,
            "message": message,
            "details": details or {}
        }
        logger.debug(f"创建验证结果: {result}")
        return result
    
    def log_validation_error(self, error_type: str, context: str) -> dict:
        """
        记录验证错误并返回结果
        
        参数:
            error_type: 错误类型
            context: 错误上下文
            
        返回:
            dict: 错误验证结果
        """
        logger.error(f"{error_type}验证失败: {context}")
        return self.create_result(False, f"{error_type}验证失败", {"context": context})

class ElementScorer(BaseValidator):
    """
    文档元素验证器
    实现公文元素评分系统和三轮识别架构
    
    核心功能:
    1. 三轮元素识别系统
    2. 公文完整性验证 (validate_document)
    3. 上下文管理 (reset_context)
    
    三轮识别架构:
    - 第1轮: 识别大部分唯一性元素
    - 第2轮: 识别发文机关标志和印发机关
    - 第3轮: 识别非唯一性元素
    """

    # 共享模式常量
    DOCUMENT_NUMBER_PATTERNS = [
        # 标准格式：机关代字 + 〔年份〕 + 序号 + 号
        # 要求必须有机关代字和年份，序号可以为空
        r'[\u4e00-\u9fa5A-Za-z0-9]+〔\d{4}〕\s*\d*\s*号',
        r'[\u4e00-\u9fa5A-Za-z0-9]+\[\d{4}\]\s*\d*\s*号',
        r'[\u4e00-\u9fa5A-Za-z0-9]+\(\d{4}\)\s*\d*\s*号',
        r'[\u4e00-\u9fa5A-Za-z0-9]+\d{4}-\s*\d*\s*号',
        
        # 支持全角空格和特殊空格字符
        r'[\u4e00-\u9fa5A-Za-z0-9]+〔\d{4}〕[\s\u3000]*[\d]*[\s\u3000]*号',
        r'[\u4e00-\u9fa5A-Za-z0-9]+\[\d{4}\][\s\u3000]*[\d]*[\s\u3000]*号',
    ]
    
    # 创建完整匹配模式（添加^和$）
    DOCUMENT_NUMBER_FULL_PATTERNS = [
        f'^{pattern}\\s*$' for pattern in DOCUMENT_NUMBER_PATTERNS
    ]

    def __init__(self):
        """初始化元素验证器"""
        super().__init__()
        logger.info("初始化文档元素验证器")
        
        # 元素类型定义 - 按标准顺序排列
        self.ELEMENT_TYPES = {
            'combined_id': '份号+发文字号',
            'copy_id': '份号',
            'security': '密级和保密期限',
            'urgency': '紧急程度',
            'issuing_authority_mark': '发文机关标志',
            'doc_number': '发文字号',
            'combined_doc_number_signer': '发文字号+签发人',
            'signer': '签发人',
            'title': '标题',
            'title_following': '标题后续部分',
            'recipient': '主送机关',
            'body': '正文区域',
            'attachment': '第1个附件',
            'attachment_following': '后续附件',
            'issuing_authority_signature': '发文机关署名',
            'issue_date': '成文日期',
            'notes': '附注',
            'disclosure': '公开方式',
            'copy_to': '抄送机关',
            'printing_authority': '印发机关',
            'printing_date': '印发日期',
            'signer_following': '签发人后续部分',
            'combined_doc_number_signer_following': '发文字号+签发人后续部分',
            'attachment_content': '附件内容'
        }
        
        # 三轮识别元素分组
        self.ROUND1_ELEMENTS = [
            'combined_id', 'copy_id', 'security', 'urgency', 'doc_number',
            'combined_doc_number_signer', 'signer',
            'title', 'recipient', 'attachment', 'issuing_authority_signature',
            'issue_date', 'notes', 'disclosure', 'copy_to', 'printing_date'
        ]
        
        self.ROUND2_ELEMENTS = [
            'issuing_authority_mark', 'printing_authority'
        ]
        
        self.ROUND3_ELEMENTS = ['title_following', 'body', 'attachment_following', 
                                'signer_following', 'combined_doc_number_signer_following', 
                                'attachment_content']
        
        # 保留原有分组用于兼容性
        self.UNIQUE_ELEMENTS = self.ROUND1_ELEMENTS + self.ROUND2_ELEMENTS
        self.NON_UNIQUE_ELEMENTS = self.ROUND3_ELEMENTS
        
        # 评分规则配置 (完整保留原规则)
        self.SCORING_RULES = {
            'combined_id': [ # 份号+发文字号
                {'condition': self.has_combined_ids, 'score': 100}
            ],
            'copy_id': [ # 份号
                {'condition': self.is_numeric_sequence, 'score': 60},
                {'condition': self.is_in_expected_position('copy_id'), 'score': 40}
            ],
            'security': [ # 密级和保密期限
                {'condition': self.starts_with(['绝密', '机密', '秘密']), 'score': 60},
                {'condition': self.is_in_expected_position('security'), 'score': 40}
            ],
            'urgency': [ # 紧急程度
                {'condition': self.starts_with_urgency_keywords, 'score': 60},
                {'condition': self.is_in_expected_position('urgency'), 'score': 40}
            ],
            'issuing_authority_mark': [ # 发文机关标志
                {'condition': self.ends_with(['人民政府', '委员会', '办公室', '局', '厅', '部', '院', '校', '中心', '办', '组', '室', '部', '司', '署', '厅', '局', '处', '科']), 'score': 50},
                {'condition': self.is_in_expected_position('issuing_authority_mark'), 'score': 30},
                {'condition': self.within_document_range(within_first=3), 'score': 30}
            ],
            'doc_number': [ # 发文字号
                {'condition': self.is_document_number_format, 'score': 60},
                {'condition': self.is_in_expected_position('doc_number'), 'score': 40}
            ],
            'combined_doc_number_signer': [ # 发文字号+签发人
                {'condition': self.has_doc_number_and_signer, 'score': 80},
                {'condition': self.follows_element('issuing_authority_mark'), 'score': 20}
            ],
            'signer': [ # 签发人
                {'condition': self.starts_with(['签发人：', '签发人:']), 'score': 60},
                {'condition': self.follows_element('issuing_authority_mark'), 'score': 20},
                {'condition': self.is_in_expected_position('signer'), 'score': 20}
            ],
            'title': [ # 标题（第1段）
                {'condition': self.is_official_title_font, 'score': 40},
                {'condition': self.is_official_title_size, 'score': 40},
                {'condition': self.is_in_expected_position('title'), 'score': 20},
                {'condition': self.matches_title_pattern, 'score': 60} 
            ],
            'recipient': [ # 主送机关
                {'condition': self.is_official_title_font, 'score': -10},
                {'condition': self.is_official_title_size, 'score': -10},
                {'condition': self.ends_with(['：', ':']), 'score': 60},
                {'condition': self.contains_any(['各', '委', '局', '办', '厅', '部', '院', '校']), 'score': 20},
                {'condition': self.is_in_expected_position('recipient'), 'score': 20}
            ],
            'attachment': [ # 附件说明（第1个）
                {'condition': self.is_official_title_font, 'score': -10},
                {'condition': self.is_official_title_size, 'score': -10},
                {'condition': self.starts_with(['附件：', '附件:', '附件1：', '附件1:']), 'score': 60},
                {'condition': self.is_in_expected_position('attachment'), 'score': 20},
                {'condition': self.is_first_attachment, 'score': 20},
                {'condition': self.is_too_close_to_recipient_or_title, 'score': -20}
            ],
            'issuing_authority_signature': [ # 发文机关署名
                {'condition': self.is_official_title_font, 'score': -10},
                {'condition': self.is_official_title_size, 'score': -10},
                {'condition': self.ends_with(['人民政府', '委员会', '办公室', '局', '厅', '部', '院', '校', '中心', '办', '组', '室', '部', '司', '署', '厅', '局', '处', '科']), 'score': 60},
                {'condition': self.follows_element('attachment'), 'score': 20},
                {'condition': self.is_in_expected_position('issuing_authority_signature'), 'score': 20},
                {'condition': self.is_too_close_to_recipient_or_title, 'score': -20},
                {'condition': self.within_document_range(within_first=5), 'score': -20}
            ],
            'issue_date': [ # 成文日期
                {'condition': self.is_official_title_font, 'score': -10},
                {'condition': self.is_official_title_size, 'score': -10},
                {'condition': self.is_standalone_date, 'score': 60},
                {'condition': self.follows_element('issuing_authority_signature'), 'score': 20},
                {'condition': self.is_in_expected_position('issue_date'), 'score': 20},
                {'condition': self.is_too_close_to_recipient_or_title, 'score': -20}
            ],
            'notes': [ # 附注
                {'condition': self.is_official_title_font, 'score': -10},
                {'condition': self.is_official_title_size, 'score': -10},
                {'condition': self.is_wrapped_in_brackets, 'score': 50},
                {'condition': self.follows_element('issue_date'), 'score': 30},
                {'condition': self.is_in_expected_position('notes'), 'score': 20},
                {'condition': self.is_too_close_to_recipient_or_title, 'score': -20}
            ],
            'disclosure': [ # 公开方式
                {'condition': self.is_official_title_font, 'score': -10},
                {'condition': self.is_official_title_size, 'score': -10},
                {'condition': self.starts_with(['公开方式：', '公开方式:', '公开属性：', '公开属性:']), 'score': 60},
                {'condition': self.is_in_expected_position('disclosure'), 'score': 40},
                {'condition': self.is_too_close_to_recipient_or_title, 'score': -20}
            ],
            'copy_to': [ # 抄送
                {'condition': self.is_official_title_font, 'score': -10},
                {'condition': self.is_official_title_size, 'score': -10},
                {'condition': self.starts_with(['抄送：', '抄送:', '报送：', '报送:', '分送：', '分送:']), 'score': 60},
                {'condition': self.is_in_expected_position('copy_to'), 'score': 40},
                {'condition': self.is_too_close_to_recipient_or_title, 'score': -20}
            ],
            'printing_authority': [ # 印发机关
                {'condition': self.follows_element('printing_date', max_empty_paras=float('inf'), reverse=True), 'score': 40},
                {'condition': self.is_in_expected_position('printing_authority'), 'score': 40},
                {'condition': self.ends_with(['人民政府', '委员会', '办公室', '局', '厅', '部', '院', '校', '中心', '办', '组', '室', '部', '司', '署', '厅', '局', '处', '科']), 'score': 40}
            ],
            'printing_date': [ # 印发日期
                {'condition': self.is_printing_date_format, 'score': 80}
            ],
            'title_following': [  # 标题后续部分
                {'condition': self.is_official_title_font, 'score': 40},
                {'condition': self.is_official_title_size, 'score': 40},
                {'condition': self.follows_element('title', max_empty_paras=0), 'score': 40},  # 紧接标题之后
            ],
            'body': [  # 正文区域（包含小标题和正文文本）
                {'condition': self.is_body_position, 'score': 100}
            ],
            'attachment_following': [  # 附件说明后续部分
                {'condition': self.is_following_attachment, 'score': 100}
            ],
            'signer_following': [  # 签发人后续部分
                {'condition': self.follows_element('signer', max_empty_paras=0), 'score': 60},
                {'condition': self.is_person_name_format, 'score': 40}
            ],
            'combined_doc_number_signer_following': [  # 发文字号+签发人后续部分
                {'condition': self.has_doc_number_and_name, 'score': 60},
                {'condition': self.follows_element('combined_doc_number_signer', max_empty_paras=0), 'score': 40}
            ],
            'attachment_content': [  # 附件内容评分规则
                {'condition': self.is_after_last_known_element, 'score': 100}
            ]
        }
        
        # 初始化上下文
        self.reset_context()
    
    def reset_context(self):
        """重置验证上下文"""
        logger.debug("重置元素验证上下文")
        
        # 使用更简洁的字典推导式初始化所有位置
        self.element_positions = {key: -1 for key in self.ELEMENT_TYPES}
        
        # 其他状态
        self.last_detected_element = None
        self.identified_unique_elements = set()
        self.last_known_element_index = -1


    def validate_document(self, detected_elements: set) -> dict:
        """
        验证公文文档完整性
        检查必要元素是否存在
        
        必要元素:
        - 标题 (title)
        - 发文机关署名 (issuing_authority_signature)
        - 成文日期 (issue_date)
        
        参数:
            detected_elements: 已识别的元素集合
            
        返回:
            dict: 验证结果
        """
        logger.info("开始公文文档完整性验证")
        
        required_elements = {'title', 'issuing_authority_signature', 'issue_date'}
        missing = required_elements - detected_elements
        
        if missing:
            missing_names = [self.ELEMENT_TYPES.get(e, e) for e in missing]
            msg = f"缺少必要公文元素: {', '.join(missing_names)}"
            logger.error(msg)
            return self.create_result(False, msg, {"missing_elements": list(missing)})
        
        logger.info("公文文档完整性验证通过")
        return self.create_result(True, "所有必要元素存在")

    def update_context(self, element_type, para_index):
        """更新元素位置上下文"""
        if element_type in self.element_positions:
            # 更新元素位置
            if self.element_positions[element_type] == -1:
                self.element_positions[element_type] = para_index
                self.last_detected_element = element_type
                element_name = self.ELEMENT_TYPES.get(element_type, element_type)
                logger.info(f"记录{element_name}位置: 段落{para_index+1}")
                logger.debug(f"更新元素位置: {element_type} -> 段落{para_index}")

            # 如果是唯一性元素，更新最后一个已知元素位置
            if element_type in self.UNIQUE_ELEMENTS:
                self.last_known_element_index = max(self.last_known_element_index, para_index)
                logger.debug(f"更新最后一个已知元素位置: {self.last_known_element_index}")
    
    def score_round1_unique_elements(self, run, context):
        """
        第1轮：识别大部分唯一性元素
        返回: (最佳元素类型, 最高分)
        """
        # 段落级别日志（更显眼）
        para_index = context.get('para_index', '未知')
        text = run['text'].strip()
        logger.info(f"======= 第1轮处理段落 {para_index+1}: {text} =======")
        
        max_score = 0
        best_type = None
        
        for element_type in self.ROUND1_ELEMENTS:
            # 如果该元素类型已被识别，跳过
            if element_type in self.identified_unique_elements:
                continue
            
            # 元素级别日志（较显眼，比段落低一级）
            logger.debug(f"--- 元素评分: {element_type} ---")
                
            element_score = 0
            for rule in self.SCORING_RULES[element_type]:
                if rule['condition'](run, context):
                    element_score += rule['score']
            
            logger.debug(f"第1轮元素 '{element_type}' 得分: {element_score}")
            
            # 达到阈值80分，确定为该元素类型
            if element_score >= 80:
                best_type = element_type
                max_score = element_score
                self.identified_unique_elements.add(element_type)
                logger.info(f"第1轮识别到元素: {element_type} (得分: {element_score})")
                break  # 找到一个就停止
        
        return best_type, max_score

    def score_round2_unique_elements(self, run, context):
        """
        第2轮：识别发文机关标志和印发机关
        返回: (最佳元素类型, 最高分)
        """
        # 段落级别日志（更显眼）
        para_index = context.get('para_index', '未知')
        text = run['text'].strip()
        logger.info(f"======= 第2轮处理段落 {para_index+1}: {text} =======")
        
        max_score = 0
        best_type = None
        
        for element_type in self.ROUND2_ELEMENTS:
            # 如果该元素类型已被识别，跳过
            if element_type in self.identified_unique_elements:
                continue
            
            # 元素级别日志（较显眼，比段落低一级）
            logger.debug(f"--- 元素评分: {element_type} ---")
                
            element_score = 0
            for rule in self.SCORING_RULES[element_type]:
                if rule['condition'](run, context):
                    element_score += rule['score']
            
            logger.debug(f"第2轮元素 '{element_type}' 得分: {element_score}")
            
            # 达到阈值80分，确定为该元素类型
            if element_score >= 80:
                best_type = element_type
                max_score = element_score
                self.identified_unique_elements.add(element_type)
                logger.info(f"第2轮识别到元素: {element_type} (得分: {element_score})")
                break  # 找到一个就停止
        
        return best_type, max_score

    def score_round3_non_unique_elements(self, run, context):
        """
        第3轮：识别非唯一性元素
        返回: (最佳元素类型, 最高分)
        """
        # 段落级别日志（更显眼）
        para_index = context.get('para_index', '未知')
        text = run['text'].strip()
        logger.info(f"======= 第3轮处理段落 {para_index+1}: {text} =======")
        
        max_score = 0
        best_type = None
        
        for element_type in self.ROUND3_ELEMENTS:
            # 元素级别日志（较显眼，比段落低一级）
            logger.debug(f"--- 元素评分: {element_type} ---")
            
            element_score = 0
            for rule in self.SCORING_RULES[element_type]:
                if rule['condition'](run, context):
                    element_score += rule['score']
            
            logger.debug(f"第3轮元素 '{element_type}' 得分: {element_score}")
            
            if element_score > max_score:
                max_score = element_score
                best_type = element_type
        
        return best_type, max_score
    
    def follows_element(self, preceding_element, max_empty_paras=None, reverse=False):
        """
        检查当前元素是否紧跟某个前置元素
        
        "紧跟"含义：两个元素之间只能有空段落，不能有内容段落
        
        参数:
            preceding_element: 前置元素类型（被紧跟的元素），如 'title', 'attachment'
            max_empty_paras: 中间允许的最大空段数
                - None: 不限制距离（元素在之后即可，无论多远）
                - 0: 必须直接相邻，中间不允许有空段
                - n: 中间最多允许n个空段
            reverse: 
                - False: 检查当前元素是否紧跟在前置元素之后
                - True: 检查前置元素是否紧跟在当前元素之后
        
        返回:
            function: 条件检查函数 checker(run, context)
            
        示例:
            # 紧贴title，不允许空段
            self.follows_element('title', max_empty_paras=0)
            
            # 在attachment之后，不限距离
            self.follows_element('attachment')
            
            # printing_date在当前之后，允许中间有任意空段
            self.follows_element('printing_date', max_empty_paras=float('inf'), reverse=True)
        """
        def checker(run, context):
            # 获取前置元素位置
            preceding_pos = self.element_positions.get(preceding_element, -1)
            if preceding_pos == -1:
                logger.debug(f"前置元素 '{preceding_element}' 未找到")
                return False
            
            current_pos = context['para_index']
            
            # 确定检查方向
            if reverse:
                # 检查前置元素是否在当前之后
                if preceding_pos <= current_pos:
                    logger.debug(f"反向检查失败: {preceding_element}位置({preceding_pos}) <= 当前位置({current_pos})")
                    return False
                start, end = current_pos, preceding_pos
            else:
                # 检查当前元素是否在前置元素之后
                if current_pos <= preceding_pos:
                    logger.debug(f"正向检查失败: 当前位置({current_pos}) <= {preceding_element}位置({preceding_pos})")
                    return False
                start, end = preceding_pos, current_pos
            
            # 如果不限制距离，直接返回True
            if max_empty_paras is None:
                logger.debug(f"位置关系检查通过（不限距离）: {preceding_element}({preceding_pos}) <-> 当前({current_pos})")
                return True
            
            # 检查中间段落
            empty_count = 0
            for i in range(start + 1, end):
                if i < context.get('total_paras', 0):
                    para_text = context.get('doc').paragraphs[i].text.strip() if context.get('doc') else ""
                    if para_text:  # 非空段落
                        logger.debug(f"中间存在非空段落(索引{i}): '{para_text[:30]}...'")
                        return False
                    empty_count += 1
            
            # 检查空段数量是否在允许范围内
            result = empty_count <= max_empty_paras
            logger.debug(f"空段数检查: {empty_count} <= {max_empty_paras} -> {result}")
            return result
        
        return checker

    def is_after_last_known_element(self, run, context):
        """
        检查是否在最后一个已知元素之后
        
        参数:
            run: 文档元素数据
            context: 上下文信息
            
        返回:
            bool: 是否在最后一个已知元素之后
        """
        # 如果没有记录最后一个已知元素，返回False
        if self.last_known_element_index == -1:
            return False
        
        # 检查当前段落索引是否大于最后一个已知元素的位置
        is_after = context['para_index'] > self.last_known_element_index
        logger.debug(f"检查是否在最后一个已知元素之后: 当前={context['para_index']}, 最后已知={self.last_known_element_index}, 结果={is_after}")
        
        return is_after

    def within_document_range(self, within_first=None, within_last=None):
        """
        检查当前元素是否在文档的指定范围内
        
        参数:
            within_first: 必须在文档前N段内
                - None: 不限制从开头的距离
                - n: 必须在前n段内（段落索引 < n）
            within_last: 必须在文档后N段内
                - None: 不限制从末尾的距离  
                - n: 必须在后n段内
        
        返回:
            function: 条件检查函数 checker(run, context)
            
        示例:
            # 检查是否在文档前3段内
            self.within_document_range(within_first=3)
            
            # 检查是否在文档后5段内
            self.within_document_range(within_last=5)
            
            # 检查是否在前3段或后5段内
            self.within_document_range(within_first=3, within_last=5)
        """
        def checker(run, context):
            current_pos = context['para_index']
            total_paras = context.get('total_paras', 0)
            
            # 检查是否在前N段内
            if within_first is not None:
                in_first = current_pos < within_first
                logger.debug(f"前段范围检查: 当前位置{current_pos} < {within_first} -> {in_first}")
                if within_last is None:
                    return in_first
                # 如果也设置了within_last，需要满足其中之一
                if in_first:
                    return True
            
            # 检查是否在后N段内
            if within_last is not None:
                in_last = current_pos >= (total_paras - within_last)
                logger.debug(f"后段范围检查: 当前位置{current_pos} >= {total_paras - within_last} -> {in_last}")
                return in_last
            
            # 如果两个参数都是None，返回True（不限制）
            return True
        
        return checker

    def starts_with(self, keywords):
        """
        检查文本是否以指定关键词开头（任意一个）
        
        参数:
            keywords: 关键词列表
            
        返回:
            function: 条件检查函数 checker(run, context)
            
        示例:
            # 检查是否以密级关键词开头
            self.starts_with(['绝密', '机密', '秘密'])
        """
        def checker(run, context):
            text = run['text'].strip()
            matched = any(text.startswith(kw) for kw in keywords)
            logger.debug(f"开头检查: '{text}' 匹配 {keywords} -> {matched}")
            return matched
        return checker

    def ends_with(self, keywords):
        """
        检查文本是否以指定关键词结尾（任意一个）
        
        参数:
            keywords: 关键词列表
            
        返回:
            function: 条件检查函数 checker(run, context)
            
        示例:
            # 检查是否以政府机关关键词结尾
            self.ends_with(['人民政府', '委员会', '办公室'])
        """
        def checker(run, context):
            text = run['text'].strip()
            matched = any(text.endswith(kw) for kw in keywords)
            logger.debug(f"结尾检查: '{text}' 匹配 {keywords} -> {matched}")
            return matched
        return checker

    def contains_any(self, keywords):
        """
        检查文本是否包含任意一个关键词（OR逻辑）
        
        参数:
            keywords: 关键词列表
            
        返回:
            function: 条件检查函数 checker(run, context)
            
        示例:
            # 检查是否包含任意一个机关关键词
            self.contains_any(['各', '委', '局', '办'])
        """
        def checker(run, context):
            text = run['text']
            matched = any(kw in text for kw in keywords)
            logger.debug(f"包含任意检查: '{text}' 匹配 {keywords} -> {matched}")
            return matched
        return checker

    def contains_all(self, keywords):
        """
        检查文本是否包含所有关键词（AND逻辑）
        
        参数:
            keywords: 关键词列表
            
        返回:
            function: 条件检查函数 checker(run, context)
            
        示例:
            # 检查是否包含所有关键词
            self.contains_all(['政府', '办公室'])
        """
        def checker(run, context):
            text = run['text']
            matched = all(kw in text for kw in keywords)
            logger.debug(f"包含全部检查: '{text}' 匹配 {keywords} -> {matched}")
            return matched
        return checker

    def is_in_expected_position(self, element_type):
        """检查元素是否在预期位置"""
        def checker(run, context):
            # 获取当前元素的标准顺序索引
            try:
                expected_idx = self.UNIQUE_ELEMENTS.index(element_type)
            except ValueError:
                return False
            
            # 获取最后一个已识别元素的索引
            last_idx = -1
            if self.last_detected_element:
                try:
                    last_idx = self.UNIQUE_ELEMENTS.index(self.last_detected_element)
                except ValueError:
                    pass
            
            # 当前元素应在最后一个元素之后
            matched = expected_idx > last_idx
            logger.debug(f"顺序检查: {element_type} 应在{self.last_detected_element}之后 -> {matched}")
            return matched
        return checker

    def is_too_close_to_recipient_or_title(self, run, context):
        """检查元素是否离主送机关或标题太近（少于2段距离）"""
        # 获取参考位置（主送机关或标题）
        reference_position = -1
        
        # 优先使用主送机关位置
        if self.element_positions['recipient'] != -1:
            reference_position = self.element_positions['recipient']
        # 如果没有主送机关，则使用标题位置
        elif self.element_positions['title'] != -1:
            reference_position = self.element_positions['title']
        else:
            # 既没有主送机关也没有标题，无法判断距离
            return False
        
        # 计算当前段落与参考位置的间隔
        distance = context['para_index'] - reference_position
        too_close = distance < 2  # 少于2段距离视为太近
        
        logger.debug(f"距离检查: 当前段落{context['para_index']}, 参考位置{reference_position}, 距离{distance}, 太近{too_close}")
        return too_close

    # ================== 条件函数 (完整保留) ==================

    def extract_document_number_parts(self, text):
        """提取发文字号的组成部分"""
        text = text.strip()
        
        # 首先检查是否是完整的发文字号
        for pattern in self.DOCUMENT_NUMBER_FULL_PATTERNS:
            if re.match(pattern, text):
                return "", text  # 份号为空，整个文本为发文字号
        
        # 如果不是完整的发文字号，尝试匹配组合格式
        copy_id_patterns = [
            r'\d+[\s\u3000\t]*',        # 标准份号
            r'\d+[\s\u3000\t]+\d+[\s\u3000\t]*',  # 份号含空格
        ]
        
        # 尝试匹配组合格式
        for copy_pattern in copy_id_patterns:
            for doc_pattern in self.DOCUMENT_NUMBER_PATTERNS:
                full_pattern = f'^({copy_pattern})({doc_pattern})\\s*$'
                
                match = re.match(full_pattern, text)
                if match:
                    copy_id = match.group(1).strip()
                    doc_number = match.group(2).strip()
                    return copy_id, doc_number
        
        return None, None

    def has_combined_ids(self, run, context):
        """检查是否包含份号+发文字号组合"""
        text = run['text'].strip()
        copy_id, doc_number = self.extract_document_number_parts(text)
        
        # 只有当份号和发文字号都有效时才认为是组合格式
        if doc_number is not None and copy_id:
            logger.debug(f"组合号格式匹配: '{text}' -> 份号: '{copy_id}', 发文字号: '{doc_number}'")
            return True
        
        logger.debug(f"组合号格式不匹配: '{text}'")
        return False

    def is_numeric_sequence(self, run, context):
        """检查是否为纯数字序列（份号）"""
        text = run['text'].strip()
        matched = text.isdigit()
        logger.debug(f"份号检查 -> {matched}")
        return matched
    
    def is_document_number_format(self, run, context):
        """检查是否符合发文字号格式）"""
        text = run['text'].strip()
        
        # 使用完整模式检查
        for pattern in self.DOCUMENT_NUMBER_FULL_PATTERNS:
            if re.match(pattern, text):
                logger.debug(f"发文字号格式匹配 -> True")
                return True
        
        logger.debug(f"发文字号格式不匹配 -> False")
        return False
    
    def starts_with_urgency_keywords(self, run, context):
        """检查是否以紧急程度关键词开头"""
        text = run['text'].strip()
        patterns = [
            r'^特急\b', r'^加急\b', r'^特提\b', r'^平急\b',
            r'^特\s*急', r'^加\s*急', r'^特\s*提', r'^平\s*急'
        ]
        matched = any(re.search(pattern, text) for pattern in patterns)
        logger.debug(f"紧急程度检查 -> {matched}")
        return matched
    
    def is_official_title_font(self, run, context):
        """检查是否标题字体"""
        font = run['fonts'].get('eastAsia', '')
        
        # 检查字体是否包含"小标宋"或"黑体"
        matched = "小标宋" in font or "黑体" in font
        
        logger.debug(f"标题字体检查 -> {matched}")
        return matched
    
    def is_official_title_size(self, run, context):
        """检查是否标题字号"""
        size_name = docx_utils.get_font_size_name(run['fonts'].get('sz'))
        matched = size_name == "二号"
        logger.debug(f"标题字号检查 -> {matched}")
        return matched
    
    def matches_title_pattern(self, run, context):
        """
        检查标题文本是否符合标题模式
        符合正则表达式则得40分
        """
        text = run['text'].strip()
        
        # 定义标题的正则表达式
        pattern = re.compile(
            r'^(?:[\u4e00-\u9fa5a-zA-Z0-9]+(?:[\s\u3000]+[\u4e00-\u9fa5a-zA-Z0-9]+)*[\s\u3000]+)?'  # 任意前缀
            r'(?:关于)?'  # 可选的"关于"
            r'[《]?(.+?)[》]?'  # 书名号包裹的内容（可选）
            r'(?:的)?'  # 可选的"的"
            r'[\s\u3000]*'  # 空格或全角空格
            r'(?:通知|决定|批复|意见|报告|请示|函|通报|通告|公告|议案|纪要|细则|办法|规定|方案|计划|指示|命令|倡议|倡议书|公示|说明)$'  # 公文类型结尾
        )
        
        matched = pattern.match(text) is not None
        logger.debug(f"标题模式检查 -> {matched}")
        return matched

    def starts_with_attachment(self, run, context):
        """检查是否以附件开头"""
        text = run['text'].strip()
        matched = text.startswith(('附件：', '附件:', '附件1：', '附件1:'))
        logger.debug(f"附件开头检查: '{text}' -> {matched}")
        return matched
    
    def is_first_attachment(self, run, context):
        """检查是否为第一个附件"""
        matched = self.element_positions['attachment'] == -1
        logger.debug(f"第一个附件检查: {matched}")
        return matched
    
    def is_following_attachment(self, run, context):
        """检查是否为后续附件"""
        if self.element_positions['attachment'] == -1:
            return False
            
        if not self.is_numbered_list_item(run, context):
            return False
            
        matched = context['para_index'] > self.element_positions['attachment']
        logger.debug(f"后续附件位置检查: {context['para_index']} > {self.element_positions['attachment']} -> {matched}")
        return matched

    def is_numbered_list_item(self, run, context):
        """检查是否为编号列表项"""
        text = run['text'].strip()
        matched = re.match(r'^\d+[\.．]\s*', text) is not None
        logger.debug(f"编号列表项检查: '{text}' -> {matched}")
        return matched

    def is_standalone_date(self, run, context):
        """检查是否为独立日期的段落（整段内容就是一个日期）"""
        text = run['text'].strip()
        
        # 优化的日期正则表达式，支持多种分隔符
        date_pattern = r'^\d{4}[年\-\.\/\\]\d{1,2}[月\-\.\/\\]\d{1,2}[日号\-\.\/\\]?$'
        
        # 检查文本是否完全匹配日期格式
        matched = re.fullmatch(date_pattern, text) is not None
        logger.debug(f"独立日期检查: '{text}' -> {matched}")
        return matched

    def is_wrapped_in_brackets(self, run, context):
        """检查是否被括号包裹（附注）"""
        text = run['text'].strip()
        
        # 定义左右括号的集合
        left_brackets = ['(', '（']
        right_brackets = [')', '）']
        
        # 检查是否以左括号开头且以右括号结尾
        if len(text) < 2:
            return False
            
        starts_with_bracket = text[0] in left_brackets
        ends_with_bracket = text[-1] in right_brackets
        
        matched = starts_with_bracket and ends_with_bracket
        logger.debug(f"括号包裹检查: '{text}' -> {matched}")
        return matched

    def is_body_position(self, run, context):
        """检查是否在正文位置"""
        if self.element_positions['recipient'] == -1 or context['para_index'] <= self.element_positions['recipient']:
            return False
        
        if self.element_positions['attachment'] != -1 and context['para_index'] >= self.element_positions['attachment']:
            return False
        
        if self.element_positions['issuing_authority_signature'] != -1 and context['para_index'] >= self.element_positions['issuing_authority_signature']:
            return False
        
        if self.element_positions['issuing_authority_signature'] == -1 and self.element_positions['issue_date'] != -1 and context['para_index'] >= self.element_positions['issue_date']:
            return False
        
        return True

    def is_printing_date_format(self, run, context):
        """检查是否符合印发日期格式（2025年7月5日印发）"""
        text = run['text'].strip()
        
        # 定义印发日期格式的正则表达式
        patterns = [
            r'^\d{4}年\d{1,2}月\d{1,2}日印发$',
            r'^\d{4}年\d{1,2}月\d{1,2}号印发$',
            r'^\d{4}年\d{1,2}月\d{1,2}日$',
            r'^\d{4}年\d{1,2}月\d{1,2}号$'
        ]
        
        matched = any(re.match(pattern, text) for pattern in patterns)
        logger.debug(f"印发日期格式检查 -> {matched}")
        return matched

    def has_doc_number_and_signer(self, run, context):
        """
        检查是否包含发文字号+签发人组合
        格式：发文字号 + 空格/制表符 + 签发人：+ 人名
        示例：XX〔2024〕1号   签发人：张三
        """
        text = run['text'].strip()
        
        # 构建匹配模式：发文字号 + 空白字符 + "签发人：" + 人名
        for doc_pattern in self.DOCUMENT_NUMBER_PATTERNS:
            # 匹配：发文字号 + 至少一个空白 + "签发人：" + 中文名字
            pattern = rf'^({doc_pattern})[\s\u3000\t]+签发人[：:][\u4e00-\u9fa5]{{2,4}}$'
            if re.match(pattern, text):
                logger.debug(f"发文字号+签发人格式匹配: '{text}'")
                return True
        
        logger.debug(f"发文字号+签发人格式不匹配: '{text}'")
        return False

    def is_person_name_format(self, run, context):
        """
        检查是否为人名格式
        特征：
        - 2-4个连续中文字符
        - 可能包含顿号、空格分隔多个人名
        - 不包含常见公文关键词
        """
        text = run['text'].strip()
        
        # 公文关键词黑名单
        doc_keywords = [
            '通知', '决定', '批复', '意见', '报告', '请示', '函',
            '通报', '通告', '公告', '议案', '纪要', '办法', '规定',
            '人民政府', '委员会', '办公室', '关于'
        ]
        
        # 如果包含公文关键词，不是人名
        if any(kw in text for kw in doc_keywords):
            logger.debug(f"包含公文关键词，不是人名格式: '{text}'")
            return False
        
        # 如果包含特殊符号（除了顿号、空格），不是人名
        if re.search(r'[：:。！？；""''《》【】()（）]', text):
            logger.debug(f"包含特殊符号，不是人名格式: '{text}'")
            return False
        
        # 分割人名（按顿号或空格）
        names = re.split(r'[、\s\u3000]+', text)
        names = [n.strip() for n in names if n.strip()]
        
        # 检查每个名字是否符合人名格式（2-4个中文字符）
        for name in names:
            if not re.match(r'^[\u4e00-\u9fa5]{2,4}$', name):
                logger.debug(f"不符合人名格式: '{name}'")
                return False
        
        matched = len(names) > 0
        logger.debug(f"人名格式检查: '{text}' -> {matched}")
        return matched

    def has_doc_number_and_name(self, run, context):
        """
        检查是否包含发文字号+人名（无"签发人："前缀）
        格式：发文字号 + 空格/制表符 + 人名
        示例：XX〔2024〕2号   李四
        """
        text = run['text'].strip()
        
        # 构建匹配模式：发文字号 + 空白字符 + 中文名字（2-4个字符）
        for doc_pattern in self.DOCUMENT_NUMBER_PATTERNS:
            pattern = rf'^({doc_pattern})[\s\u3000\t]+[\u4e00-\u9fa5]{{2,4}}$'
            if re.match(pattern, text):
                logger.debug(f"发文字号+人名格式匹配: '{text}'")
                return True
        
        logger.debug(f"发文字号+人名格式不匹配: '{text}'")
        return False

    def __str__(self):
        """返回验证器信息"""
        return f"公文元素评分器 | 支持{len(self.ELEMENT_TYPES)}种元素类型"

# 保留原有创建方式
def create_element_scorer():
    """创建元素验证器实例"""
    return ElementScorer()
