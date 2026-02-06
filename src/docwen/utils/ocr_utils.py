"""
OCR工具模块
提供图片文字识别功能，支持多语言识别

版本: RapidOCR (基于 ONNX Runtime)
特点: 纯本地推理，无需网络，体积小巧

支持的语言:
- chinese: 中文/英文 (默认)
- japanese: 日语
- korean: 韩语
- latin: 拉丁语系 (德语/法语/葡萄牙语/西班牙语/越南语等)
- cyrillic: 西里尔语系 (俄语等)
"""

import os
import logging
from typing import List, Dict, Optional

logger = logging.getLogger(__name__)

# 全局OCR实例（单例模式）
_ocr_instance = None
_current_ocr_language = None  # 当前加载的OCR语言

# OCR语言配置值
OCR_LANGUAGE_AUTO = "auto"
OCR_LANGUAGE_CHINESE = "chinese"
OCR_LANGUAGE_CHINESE_CHT = "chinese_cht"
OCR_LANGUAGE_ENGLISH = "english"
OCR_LANGUAGE_JAPANESE = "japanese"
OCR_LANGUAGE_KOREAN = "korean"
OCR_LANGUAGE_LATIN = "latin"
OCR_LANGUAGE_CYRILLIC = "cyrillic"

# 界面语言到OCR语言的映射
LOCALE_TO_OCR_LANGUAGE = {
    "zh_CN": OCR_LANGUAGE_CHINESE,
    "zh_TW": OCR_LANGUAGE_CHINESE_CHT,  # 繁体中文使用繁体模型
    "en_US": OCR_LANGUAGE_ENGLISH,  # 英文界面使用纯英文模型
    "ja_JP": OCR_LANGUAGE_JAPANESE,
    "ko_KR": OCR_LANGUAGE_KOREAN,
    "de_DE": OCR_LANGUAGE_LATIN,
    "fr_FR": OCR_LANGUAGE_LATIN,
    "pt_BR": OCR_LANGUAGE_LATIN,
    "es_ES": OCR_LANGUAGE_LATIN,
    "vi_VN": OCR_LANGUAGE_LATIN,
    "ru_RU": OCR_LANGUAGE_CYRILLIC,
}

# OCR语言到模型文件的映射
OCR_LANGUAGE_MODELS = {
    OCR_LANGUAGE_CHINESE: {
        "det": "ch_PP-OCRv4_det_infer.onnx",
        "rec": "ch_PP-OCRv4_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",
    },
    OCR_LANGUAGE_CHINESE_CHT: {
        "det": "ch_PP-OCRv4_det_infer.onnx",  # 检测模型共用中文
        "rec": "chinese_cht_PP-OCRv3_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",  # 方向分类共用
    },
    OCR_LANGUAGE_ENGLISH: {
        "det": "ch_PP-OCRv4_det_infer.onnx",  # 检测模型共用中文
        "rec": "en_PP-OCRv4_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",  # 方向分类共用
    },
    OCR_LANGUAGE_JAPANESE: {
        "det": "ch_PP-OCRv4_det_infer.onnx",  # 检测模型共用中文
        "rec": "japan_PP-OCRv4_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",  # 方向分类共用
    },
    OCR_LANGUAGE_KOREAN: {
        "det": "ch_PP-OCRv4_det_infer.onnx",  # 检测模型共用中文
        "rec": "korean_PP-OCRv4_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",  # 方向分类共用
    },
    OCR_LANGUAGE_LATIN: {
        "det": "ch_PP-OCRv4_det_infer.onnx",  # 检测模型共用中文
        "rec": "latin_PP-OCRv3_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",  # 方向分类共用
    },
    OCR_LANGUAGE_CYRILLIC: {
        "det": "ch_PP-OCRv4_det_infer.onnx",  # 检测模型共用中文
        "rec": "cyrillic_PP-OCRv3_rec_infer.onnx",
        "cls": "ch_ppocr_mobile_v2.0_cls_infer.onnx",  # 方向分类共用
    },
}


def get_configured_ocr_language() -> str:
    """
    获取配置的OCR语言
    
    返回:
        str: OCR语言配置值（auto/chinese/japanese/latin/cyrillic）
    """
    try:
        from docwen.config import ConfigManager
        config_manager = ConfigManager()
        return config_manager.get_ocr_language()
    except Exception as e:
        logger.warning(f"读取OCR语言配置失败，使用默认值: {e}")
        return OCR_LANGUAGE_AUTO


def resolve_ocr_language(ocr_language: str = None) -> str:
    """
    解析实际使用的OCR语言
    
    如果配置为 auto，则根据当前界面语言自动选择
    
    参数:
        ocr_language: OCR语言配置值，如果为None则从配置读取
        
    返回:
        str: 实际使用的OCR语言（chinese/japanese/latin/cyrillic）
    """
    if ocr_language is None:
        ocr_language = get_configured_ocr_language()
    
    if ocr_language == OCR_LANGUAGE_AUTO:
        # 获取当前界面语言
        try:
            from docwen.i18n import get_current_locale
            current_locale = get_current_locale()
            resolved = LOCALE_TO_OCR_LANGUAGE.get(current_locale, OCR_LANGUAGE_CHINESE)
            logger.debug(f"OCR语言自动解析: 界面语言 {current_locale} -> OCR语言 {resolved}")
            return resolved
        except Exception as e:
            logger.warning(f"获取界面语言失败，使用中文模型: {e}")
            return OCR_LANGUAGE_CHINESE
    
    return ocr_language


def reset_ocr():
    """
    重置OCR实例
    
    当OCR语言配置变更时调用此函数，下次 get_ocr() 会重新初始化
    """
    global _ocr_instance, _current_ocr_language
    
    if _ocr_instance is not None:
        logger.info("重置OCR实例")
        _ocr_instance = None
        _current_ocr_language = None


def get_ocr(ocr_language: str = None):
    """
    获取OCR实例（单例模式，支持多语言）
    
    首次调用时会初始化RapidOCR，后续调用复用同一实例
    如果语言配置变更，会自动重新初始化
    使用项目内置的ONNX模型，完全离线运行
    
    参数:
        ocr_language: 指定OCR语言，如果为None则从配置读取
    
    返回:
        RapidOCR实例
    """
    global _ocr_instance, _current_ocr_language
    
    # 解析实际使用的OCR语言
    target_language = resolve_ocr_language(ocr_language)
    
    # 检查是否需要重新初始化（语言变更）
    if _ocr_instance is not None and _current_ocr_language != target_language:
        logger.info(f"OCR语言变更: {_current_ocr_language} -> {target_language}，重新初始化")
        _ocr_instance = None
    
    if _ocr_instance is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            
            # 获取项目根目录（使用统一的路径工具）
            from docwen.utils.path_utils import get_project_root
            base_dir = get_project_root()
            
            # 模型目录路径
            model_dir = os.path.join(base_dir, "models", "rapidocr")
            
            # 获取对应语言的模型文件
            models = OCR_LANGUAGE_MODELS.get(target_language, OCR_LANGUAGE_MODELS[OCR_LANGUAGE_CHINESE])
            
            det_model_path = os.path.join(model_dir, models["det"])
            rec_model_path = os.path.join(model_dir, models["rec"])
            cls_model_path = os.path.join(model_dir, models["cls"])
            
            logger.info(f"初始化RapidOCR (语言: {target_language})...")
            logger.info(f"模型目录: {model_dir}")
            
            # 检查模型文件是否存在
            missing_models = []
            if not os.path.exists(det_model_path):
                missing_models.append(f"检测模型: {det_model_path}")
            if not os.path.exists(rec_model_path):
                missing_models.append(f"识别模型: {rec_model_path}")
            if not os.path.exists(cls_model_path):
                missing_models.append(f"方向分类模型: {cls_model_path}")
            
            if missing_models:
                error_msg = f"OCR模型文件缺失 (语言: {target_language})，请下载模型文件到 models/rapidocr/ 目录:\n" + "\n".join(missing_models)
                logger.error(error_msg)
                raise FileNotFoundError(error_msg)
            
            # 记录模型文件大小
            logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            logger.info(f"模型文件检查 (语言: {target_language})")
            for model_name, model_path in [
                ("检测模型", det_model_path),
                ("识别模型", rec_model_path),
                ("方向分类模型", cls_model_path)
            ]:
                file_size = os.path.getsize(model_path)
                logger.info(f"  ✓ {model_name}: {file_size:,} 字节")
            logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            
            # 初始化RapidOCR
            logger.info("开始初始化RapidOCR实例...")
            
            _ocr_instance = RapidOCR(
                det_model_path=det_model_path,
                rec_model_path=rec_model_path,
                cls_model_path=cls_model_path,
            )
            
            _current_ocr_language = target_language
            
            logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            logger.info("✓ RapidOCR初始化成功")
            logger.info(f"  OCR语言: {target_language}")
            logger.info("  推理引擎: ONNX Runtime")
            logger.info("  运行模式: 纯本地离线")
            logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                
        except ImportError as e:
            logger.error(f"RapidOCR导入失败: {e}")
            logger.error("请安装: pip install rapidocr_onnxruntime")
            raise ImportError("RapidOCR未安装，请运行: pip install rapidocr_onnxruntime")
        except Exception as e:
            logger.error(f"RapidOCR初始化失败: {e}")
            raise
    
    return _ocr_instance


def extract_text_simple(image_path: str, cancel_event=None) -> str:
    """
    从图片中提取纯文本（用于Markdown）
    
    参数:
        image_path: 图片文件路径
        cancel_event: 取消事件（可选），用于中断OCR识别
        
    返回:
        str: 识别的文本，多行用换行符分隔
        
    注意:
        - 如果识别失败，返回空字符串
        - GIF格式只识别第一帧
        - 如果cancel_event被设置，立即返回空字符串
    """
    # 检查取消事件
    if cancel_event and cancel_event.is_set():
        logger.info("OCR识别被取消（操作前检查）")
        return ""
    
    if not os.path.exists(image_path):
        logger.error(f"图片文件不存在: {image_path}")
        return ""
    
    try:
        # 检查文件格式
        _, ext = os.path.splitext(image_path)
        ext = ext.lower()
        
        if ext == '.gif':
            logger.warning(f"GIF格式只会识别第一帧: {image_path}")
        
        if ext in ['.heic', '.heif']:
            logger.error(f"HEIC/HEIF格式应该已被转换为PNG: {image_path}")
            return ""
        
        # 执行OCR识别 - RapidOCR 使用 __call__ 方法
        logger.debug(f"开始OCR识别: {image_path}")
        ocr = get_ocr()
        result, elapse = ocr(image_path)
        
        # OCR完成后再次检查取消事件（处理OCR过程中的取消）
        if cancel_event and cancel_event.is_set():
            logger.info("OCR识别被取消（操作后检查）")
            return ""
        
        # RapidOCR 返回格式: [[box, text, confidence], ...] 或 None
        if not result:
            logger.warning(f"OCR未识别到文字: {image_path}")
            return ""
        
        # 提取所有文本行
        texts = []
        for item in result:
            # item 格式: [box, text, confidence]
            text = item[1]
            confidence = item[2]
            
            # 只保留置信度较高的结果
            if confidence > 0.5:
                texts.append(text)
                logger.debug(f"识别文字: {text} (置信度: {confidence:.2f})")
        
        result_text = '\n'.join(texts)
        # elapse 是列表（包含各阶段耗时），需要求和得到总耗时
        total_time = sum(elapse) if isinstance(elapse, list) else elapse
        logger.info(f"OCR识别完成，共识别 {len(texts)} 行文字，耗时 {total_time:.2f}秒")
        return result_text
        
    except Exception as e:
        logger.error(f"OCR识别失败: {e}", exc_info=True)
        return ""


def extract_text_with_sizes(image_path: str, cancel_event=None) -> List[Dict[str, any]]:
    """
    从图片中提取文字和字号信息（用于DOCX）
    
    参数:
        image_path: 图片文件路径
        cancel_event: 取消事件（可选），用于中断OCR识别
        
    返回:
        List[Dict]: 文字块列表，每个字典包含:
            - text: 文字内容
            - font_size: 估算的字号（磅值）
            
    注意:
        - 如果识别失败，返回空列表
        - 字号基于文字高度估算，可能不完全准确
        - 如果cancel_event被设置，立即返回空列表
    """
    # 检查取消事件
    if cancel_event and cancel_event.is_set():
        logger.info("OCR识别被取消（操作前检查）")
        return []
    
    if not os.path.exists(image_path):
        logger.error(f"图片文件不存在: {image_path}")
        return []
    
    try:
        # 检查文件格式
        _, ext = os.path.splitext(image_path)
        ext = ext.lower()
        
        if ext == '.gif':
            logger.warning(f"GIF格式只会识别第一帧: {image_path}")
        
        if ext in ['.heic', '.heif']:
            logger.error(f"HEIC/HEIF格式应该已被转换为PNG: {image_path}")
            return []
        
        # 执行OCR识别 - RapidOCR 使用 __call__ 方法
        logger.debug(f"开始OCR识别（含字号）: {image_path}")
        ocr = get_ocr()
        result, elapse = ocr(image_path)
        
        # OCR完成后再次检查取消事件（处理OCR过程中的取消）
        if cancel_event and cancel_event.is_set():
            logger.info("OCR识别被取消（操作后检查）")
            return []
        
        # RapidOCR 返回格式: [[box, text, confidence], ...] 或 None
        if not result:
            logger.warning(f"OCR未识别到文字: {image_path}")
            return []
        
        # 提取文字和字号
        text_blocks = []
        for item in result:
            # item 格式: [box, text, confidence]
            # box 格式: [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
            bbox = item[0]
            text = item[1]
            confidence = item[2]
            
            # 只保留置信度较高的结果
            if confidence > 0.5:
                # 计算文字高度（像素）
                # bbox 格式: [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
                if isinstance(bbox, (list, tuple)) and len(bbox) >= 4:
                    height = bbox[2][1] - bbox[0][1]
                    
                    # 估算字号
                    font_size = estimate_font_size(height)
                    
                    text_blocks.append({
                        'text': text,
                        'font_size': font_size
                    })
                    logger.debug(f"识别文字: {text} (高度: {height:.0f}px, 字号: {font_size}pt, 置信度: {confidence:.2f})")
                else:
                    # 如果边界框格式不正确，使用默认字号
                    text_blocks.append({
                        'text': text,
                        'font_size': 12
                    })
                    logger.debug(f"识别文字: {text} (边界框格式异常，使用默认字号12pt, 置信度: {confidence:.2f})")
        
        # elapse 是列表（包含各阶段耗时），需要求和得到总耗时
        total_time = sum(elapse) if isinstance(elapse, list) else elapse
        logger.info(f"OCR识别完成，共识别 {len(text_blocks)} 行文字，耗时 {total_time:.2f}秒")
        return text_blocks
        
    except Exception as e:
        logger.error(f"OCR识别失败: {e}", exc_info=True)
        return []


def estimate_font_size(pixel_height: float, image_dpi: int = 96) -> float:
    """
    根据文字像素高度估算字号（磅值）
    
    参数:
        pixel_height: 文字高度（像素）
        image_dpi: 图片DPI（默认96）
        
    返回:
        float: 估算的字号（磅值）
        
    说明:
        - 1磅 = 1/72 英寸
        - 在96 DPI下，1英寸 = 96像素
        - 返回最接近的常用字号
    """
    try:
        # 转换为英寸
        inches = pixel_height / image_dpi
        
        # 转换为磅值
        points = inches * 72
        
        # 常用字号列表
        common_sizes = [9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
        
        # 找到最接近的常用字号
        closest_size = min(common_sizes, key=lambda x: abs(x - points))
        
        logger.debug(f"像素高度 {pixel_height:.0f}px -> {points:.1f}pt -> 字号 {closest_size}pt")
        return closest_size
        
    except Exception as e:
        logger.error(f"字号估算失败: {e}")
        return 12  # 默认返回12磅
