"""
OCR工具模块
提供图片文字识别功能，支持中英文识别

版本: PaddleOCR 3.x (≥ 3.0.1)
要求: PaddlePaddle ≥ 3.0.0
"""

import os
import logging
from typing import List, Dict, Optional

logger = logging.getLogger(__name__)

# 全局OCR实例（单例模式）
_ocr_instance = None

# ============================================================
# 彻底禁用PaddlePaddle和PaddleOCR的所有联网功能
# 这些环境变量必须在导入paddle之前设置
# ============================================================

# 1. 设置模型主目录（使用本地models目录）
models_base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", "models"))
os.environ['HUB_HOME'] = models_base_dir
os.environ['PADDLE_HOME'] = models_base_dir
os.environ['PADDLE_MODEL_HOME'] = os.path.join(models_base_dir, "paddleocr")

# 2. 禁用所有网络下载功能
os.environ['PADDLE_DOWNLOAD_SERVER'] = 'none'  # 禁用下载服务器
os.environ['PADDLE_FORCE_LOCAL'] = '1'  # 强制使用本地模型
os.environ['DISABLE_TELEMETRY'] = '1'  # 禁用遥测数据上传

# 3. 禁用PaddleHub联网
os.environ['SERVER_URL'] = ''  # 清空服务器URL
os.environ['NO_PROXY'] = '*'  # 禁用代理

# 4. 强制离线模式
os.environ['PADDLE_SERVING_OFFLINE'] = '1'  # PaddleServing离线模式

# 5. 禁用API分析和性能监控（可能涉及网络上报）
os.environ['FLAGS_enable_api_profile'] = '0'

# 6. 日志控制
os.environ['PADDLEOCR_LOG_LEVEL'] = 'ERROR'  # 只显示错误级别日志
os.environ['GLOG_v'] = '0'  # 减少详细日志输出

logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
logger.info("OCR离线保护已启用")
logger.info(f"  模型目录: {models_base_dir}")
logger.info(f"  网络下载: 已禁用")
logger.info(f"  强制离线: 是")
logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")


def get_ocr():
    """
    获取OCR实例（单例模式）
    
    首次调用时会初始化PaddleOCR，后续调用复用同一实例
    会自动检测并使用可用的GPU，如果没有GPU则使用CPU
    使用项目内置的OCR模型，无需联网下载
    
    返回:
        PaddleOCR实例
    """
    global _ocr_instance
    
    if _ocr_instance is None:
        try:
            from paddleocr import PaddleOCR
            import paddle
            
            # 检测是否有可用的GPU
            has_gpu = paddle.device.is_compiled_with_cuda() and paddle.device.cuda.device_count() > 0
            
            if has_gpu:
                logger.info("检测到GPU，将使用GPU加速OCR识别")
                device = "gpu"
            else:
                logger.info("未检测到GPU或GPU不可用，将使用CPU进行OCR识别")
                device = "cpu"
            
            # 获取项目根目录（使用统一的路径工具）
            from gongwen_converter.utils.path_utils import get_project_root
            base_dir = get_project_root()
            
            # 模型目录路径
            model_dir = os.path.join(base_dir, "models", "paddleocr")
            
            # 检测模型：PP-OCRv5 server版
            det_model_dir = os.path.join(model_dir, "det", "PP-OCRv5_server_det")
            
            # 识别模型：PP-OCRv5 server版
            rec_model_dir = os.path.join(model_dir, "rec", "PP-OCRv5_server_rec")
            
            # 文本行方向分类模型（新版本名称）
            cls_model_dir = os.path.join(model_dir, "cls", "PP-LCNet_x1_0_textline_ori")
            
            # 文档方向分类模型
            doc_ori_model_dir = os.path.join(model_dir, "doc_ori", "PP-LCNet_x1_0_doc_ori")
            
            logger.info("初始化PaddleOCR...")
            logger.info(f"模型目录: {model_dir}")
            
            # 严格检查模型是否存在，不存在则报错（禁止联网下载）
            missing_models = []
            if not os.path.exists(det_model_dir):
                missing_models.append(f"检测模型: {det_model_dir}")
            if not os.path.exists(rec_model_dir):
                missing_models.append(f"识别模型: {rec_model_dir}")
            if not os.path.exists(cls_model_dir):
                missing_models.append(f"文本行方向分类模型: {cls_model_dir}")
            if not os.path.exists(doc_ori_model_dir):
                missing_models.append(f"文档方向分类模型: {doc_ori_model_dir}")
            
            if missing_models:
                error_msg = "OCR模型文件缺失，请按照 models/下载指南.md 下载模型文件:\n" + "\n".join(missing_models)
                logger.error(error_msg)
                raise FileNotFoundError(error_msg)
            
            # 进一步检查每个模型目录的关键文件
            logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            logger.info("详细检查模型文件完整性")
            
            # 只检查必需的参数文件，不检查可选的模型结构文件
            # 注意：某些PaddleOCR模型不包含inference.pdmodel文件也能正常工作
            required_file = 'inference.pdiparams'  # 参数文件是必需的
            all_models_valid = True
            
            for model_name, model_path in [
                ("检测模型", det_model_dir),
                ("识别模型", rec_model_dir),
                ("文本行方向分类模型", cls_model_dir),
                ("文档方向分类模型", doc_ori_model_dir)
            ]:
                logger.info(f"  检查 {model_name}:")
                logger.info(f"    路径: {model_path}")
                
                # 检查必需的参数文件
                param_file = os.path.join(model_path, required_file)
                if os.path.exists(param_file):
                    file_size = os.path.getsize(param_file)
                    logger.info(f"    ✓ {required_file} ({file_size:,} 字节)")
                else:
                    logger.error(f"    ✗ 缺失必需文件: {required_file}")
                    all_models_valid = False
                
                # 检查可选的模型结构文件（仅记录，不强制要求）
                model_file = os.path.join(model_path, 'inference.pdmodel')
                if os.path.exists(model_file):
                    file_size = os.path.getsize(model_file)
                    logger.debug(f"    ✓ inference.pdmodel ({file_size:,} 字节)")
                else:
                    logger.debug(f"    ℹ inference.pdmodel 不存在（可选文件）")
            
            if not all_models_valid:
                error_msg = "模型参数文件不完整，请重新下载完整的模型文件"
                logger.error(error_msg)
                raise FileNotFoundError(error_msg)
            
            logger.info("✓ 所有必需模型文件检查通过")
            logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            
            # 初始化PaddleOCR 3.x - 使用源码确认的参数
            # 参考：paddleocr/_pipelines/ocr.py 和 paddleocr/_common_args.py
            logger.info("开始初始化PaddleOCR实例...")
            
            ocr_params = {
                # OCR特定参数（使用新参数名）
                'lang': 'ch',
                'text_detection_model_dir': det_model_dir,
                'text_recognition_model_dir': rec_model_dir,
                'textline_orientation_model_dir': cls_model_dir,
                'use_textline_orientation': True,
                # 文档预处理参数
                'use_doc_unwarping': False,  # 禁用文档矫正（不需要）
                # 基础参数
                'device': device,  # 'gpu' 或 'cpu'（自动检测）
                'cpu_threads': 4,
                'enable_mkldnn': False,  # 避免OneDNN错误
            }
            
            # 启用文档方向分类
            ocr_params['doc_orientation_classify_model_dir'] = doc_ori_model_dir
            ocr_params['use_doc_orientation_classify'] = True
            
            try:
                # 尝试初始化PaddleOCR
                _ocr_instance = PaddleOCR(**ocr_params)
                logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                logger.info("✓ PaddleOCR初始化成功")
                logger.info("  设备类型: " + ("GPU" if device == "gpu" else "CPU"))
                logger.info("  离线模式: 已启用")
                logger.info("  模型加载: 本地模型")
                logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                
            except Exception as init_error:
                # 捕获初始化过程中的任何异常
                error_type = type(init_error).__name__
                logger.error("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                logger.error("✗ PaddleOCR初始化失败")
                logger.error(f"  异常类型: {error_type}")
                logger.error(f"  异常消息: {str(init_error)}")
                
                # 检查是否是网络相关错误
                error_msg_lower = str(init_error).lower()
                network_keywords = ['network', 'connection', 'timeout', 'download', 'url', 'http', 'request']
                is_network_error = any(keyword in error_msg_lower for keyword in network_keywords)
                
                if is_network_error:
                    logger.error("  ⚠ 检测到网络相关错误！")
                    logger.error("  说明：PaddleOCR尝试联网但被阻止")
                    logger.error("  解决：请确保所有模型文件已正确下载到本地")
                
                logger.error("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
                raise
            
        except ImportError as e:
            logger.error(f"PaddleOCR导入失败: {e}")
            logger.error("请安装: pip install paddleocr paddlepaddle")
            raise ImportError("PaddleOCR未安装，请运行: pip install paddleocr paddlepaddle")
        except Exception as e:
            logger.error(f"PaddleOCR初始化失败: {e}")
            raise
    
    return _ocr_instance


def extract_text_simple(image_path: str) -> str:
    """
    从图片中提取纯文本（用于Markdown）
    
    参数:
        image_path: 图片文件路径
        
    返回:
        str: 识别的文本，多行用换行符分隔
        
    注意:
        - 如果识别失败，返回空字符串
        - GIF格式只识别第一帧
    """
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
        
        # 执行OCR识别 - PaddleOCR 3.x 使用 predict() 方法
        logger.debug(f"开始OCR识别: {image_path}")
        ocr = get_ocr()
        result = ocr.predict(image_path)  # 3.x 使用 predict() 替代 ocr()
        
        # 3.x 返回字典结构，需要访问 rec_texts 字段
        if not result or len(result) == 0 or 'rec_texts' not in result[0]:
            logger.warning(f"OCR未识别到文字: {image_path}")
            return ""
        
        # 提取所有文本行 - 3.x 结构为 result[0]['rec_texts']
        rec_texts = result[0]['rec_texts']
        rec_scores = result[0].get('rec_scores', [1.0] * len(rec_texts))  # 获取置信度
        
        texts = []
        for text, confidence in zip(rec_texts, rec_scores):
            # 只保留置信度较高的结果
            if confidence > 0.5:
                texts.append(text)
                logger.debug(f"识别文字: {text} (置信度: {confidence:.2f})")
        
        result_text = '\n'.join(texts)
        logger.info(f"OCR识别完成，共识别 {len(texts)} 行文字")
        return result_text
        
    except Exception as e:
        logger.error(f"OCR识别失败: {e}", exc_info=True)
        return ""


def extract_text_with_sizes(image_path: str) -> List[Dict[str, any]]:
    """
    从图片中提取文字和字号信息（用于DOCX）
    
    参数:
        image_path: 图片文件路径
        
    返回:
        List[Dict]: 文字块列表，每个字典包含:
            - text: 文字内容
            - font_size: 估算的字号（磅值）
            
    注意:
        - 如果识别失败，返回空列表
        - 字号基于文字高度估算，可能不完全准确
    """
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
        
        # 执行OCR识别 - PaddleOCR 3.x 使用 predict() 方法
        logger.debug(f"开始OCR识别（含字号）: {image_path}")
        ocr = get_ocr()
        result = ocr.predict(image_path)  # 3.x 使用 predict() 替代 ocr()
        
        # 3.x 返回字典结构，需要访问多个字段
        if not result or len(result) == 0 or 'rec_texts' not in result[0]:
            logger.warning(f"OCR未识别到文字: {image_path}")
            return []
        
        # 提取文字、置信度和边界框
        rec_texts = result[0]['rec_texts']
        rec_scores = result[0].get('rec_scores', [1.0] * len(rec_texts))
        # 在3.x中，边界框信息在 det_boxes 或 rec_boxes 字段中
        det_boxes = result[0].get('det_boxes', result[0].get('rec_boxes', []))
        
        if det_boxes is None or len(det_boxes) == 0 or len(det_boxes) != len(rec_texts):
            # 如果没有边界框信息，使用默认字号
            logger.warning("未获取到边界框信息，使用默认字号")
            text_blocks = []
            for text, confidence in zip(rec_texts, rec_scores):
                if confidence > 0.5:
                    text_blocks.append({
                        'text': text,
                        'font_size': 12  # 默认字号
                    })
            logger.info(f"OCR识别完成，共识别 {len(text_blocks)} 行文字（使用默认字号）")
            return text_blocks
        
        # 提取文字和字号
        text_blocks = []
        for text, confidence, bbox in zip(rec_texts, rec_scores, det_boxes):
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
        
        logger.info(f"OCR识别完成，共识别 {len(text_blocks)} 行文字")
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


# 模块测试
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s"
    )
    
    logger.info("OCR工具模块测试开始")
    
    # 测试文件路径
    test_image = "test.jpg"
    
    if os.path.exists(test_image):
        # 测试纯文本提取
        logger.info("测试纯文本提取...")
        text = extract_text_simple(test_image)
        logger.info(f"识别结果:\n{text}")
        
        # 测试文字+字号提取
        logger.info("测试文字+字号提取...")
        blocks = extract_text_with_sizes(test_image)
        for block in blocks:
            logger.info(f"文字: {block['text']}, 字号: {block['font_size']}pt")
    else:
        logger.warning(f"测试文件不存在: {test_image}")
    
    logger.info("OCR工具模块测试结束")
