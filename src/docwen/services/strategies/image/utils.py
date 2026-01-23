"""
图片策略辅助函数模块

提供图片转换策略所需的辅助函数：
- 格式描述生成
- 多页TIFF检测和页面提取

依赖：
- PIL: 图片处理
- path_utils: 输出路径生成
"""

import os
import re
import logging
import datetime
from typing import Optional, Callable, List, Tuple

from PIL import Image

from docwen.i18n import t

logger = logging.getLogger(__name__)


def get_image_format_description(actual_format: str) -> str:
    """
    根据图片格式生成描述（使用真实格式）
    
    参数:
        actual_format: 图片的真实格式（如 'png', 'jpg', 'bmp'）
        
    返回:
        str: 格式描述，如 'fromPng', 'fromJpg' 等
    """
    format_map = {
        'png': 'fromPng',
        'jpg': 'fromJpg',
        'jpeg': 'fromJpeg',
        'tif': 'fromTif',
        'tiff': 'fromTiff',
        'gif': 'fromGif',
        'webp': 'fromWebp',
        'bmp': 'fromBmp',
        'heic': 'fromHeic',
        'heif': 'fromHeif',
    }
    
    return format_map.get(actual_format.lower() if actual_format else '', 'fromImage')


def is_multipage_tiff(file_path: str) -> bool:
    """
    检测TIFF是否为多页
    
    参数:
        file_path: 图片文件路径
        
    返回:
        bool: 是否为多页TIFF
    """
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    if ext not in ['.tif', '.tiff']:
        return False
    
    try:
        with Image.open(file_path) as img:
            # 尝试获取帧数
            n_frames = getattr(img, 'n_frames', 1)
            return n_frames > 1
    except Exception as e:
        logger.warning(f"检测TIFF页数失败: {e}")
        return False


def extract_tiff_pages(
    file_path: str,
    output_dir: str = None,
    actual_format: str = 'tiff',
    progress_callback: Optional[Callable[[str], None]] = None
) -> List[Tuple[int, str]]:
    """
    提取TIFF每一页为独立的PNG文件，使用标准化命名规则
    
    文件命名格式：{原始基础名}_page{N}_{时间戳}_from{格式}.png
    示例：报告_page1_20250109_123000_fromTiff.png
    
    参数:
        file_path: TIFF文件路径
        output_dir: 输出目录（必需）。拆分的PNG文件将保存到此目录
        actual_format: 实际文件格式（默认'tiff'），用于生成描述信息
        progress_callback: 进度回调函数（可选）
        
    返回:
        list: 文件路径列表 [(页码, 文件路径), ...]
        
    注意:
        - 所有页面共享同一个时间戳，确保同一批次拆分的文件时间戳一致
        - 为避免时间戳冲突，在循环外生成统一时间戳，然后手动构建文件名
    """
    if not output_dir:
        raise ValueError("extract_tiff_pages 需要指定 output_dir 参数")
    
    temp_files = []
    
    try:
        with Image.open(file_path) as img:
            n_frames = getattr(img, 'n_frames', 1)
            logger.info(f"TIFF文件共 {n_frames} 页，开始拆分")
            
            # 步骤1：生成统一时间戳（在循环外）
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            logger.debug(f"生成统一时间戳: {timestamp}")
            
            # 步骤2：提取原始文件名（清理旧时间戳）
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            
            # 移除旧时间戳和描述
            timestamp_pattern = r'(_\d{8}_\d{6})(?:.*)?$'
            match = re.search(timestamp_pattern, base_name)
            if match:
                base_name_clean = base_name[:match.start()]
            else:
                base_name_clean = base_name
            
            logger.debug(f"清理后的基础文件名: {base_name_clean}")
            
            # 步骤3：生成描述
            description = f'from{actual_format.capitalize()}'
            
            # 步骤4：遍历每一页
            for i in range(n_frames):
                img.seek(i)
                page_num = i + 1
                
                # 更新进度
                if progress_callback:
                    progress_callback(t('conversion.progress.extracting_page', page=page_num))
                
                # 手动构建文件名（复用统一时间戳）
                # 格式：{原始名}_page{N}_{时间戳}_{描述}.png
                filename = f"{base_name_clean}_page{page_num}_{timestamp}_{description}.png"
                page_path = os.path.join(output_dir, filename)
                
                logger.debug(f"第 {page_num} 页输出路径: {os.path.basename(page_path)}")
                
                # 转换并保存当前帧
                # RGBA模式转为RGB（白色背景）
                frame = img.copy()
                if frame.mode == 'RGBA':
                    background = Image.new('RGB', frame.size, (255, 255, 255))
                    background.paste(frame, mask=frame.split()[3])
                    frame = background
                    logger.debug(f"第 {page_num} 页: RGBA转RGB（白色背景）")
                elif frame.mode != 'RGB':
                    frame = frame.convert('RGB')
                    logger.debug(f"第 {page_num} 页: {frame.mode}转RGB")
                
                # 保存为PNG格式
                frame.save(page_path)
                temp_files.append((page_num, page_path))
                logger.info(f"✓ 第 {page_num} 页提取成功: {os.path.basename(page_path)}")
        
        logger.info(f"TIFF拆分完成，共提取 {len(temp_files)} 页")
        return temp_files
    
    except Exception as e:
        # 清理已创建的文件
        logger.error(f"TIFF拆分失败，清理已创建的文件")
        for _, temp_path in temp_files:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                    logger.debug(f"已删除: {temp_path}")
            except Exception as cleanup_error:
                logger.warning(f"清理文件失败: {temp_path}, 错误: {cleanup_error}")
        
        logger.error(f"提取TIFF页面失败: {e}", exc_info=True)
        raise
