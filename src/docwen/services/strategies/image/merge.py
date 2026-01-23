"""
图片合并策略模块

将多个图片合并为单个多页TIFF文件。

依赖：
- PIL: 图片处理
"""

import os
import logging
from datetime import datetime
from typing import Dict, Any, Callable, Optional

from PIL import Image

from ..base_strategy import BaseStrategy
from docwen.services.result import ConversionResult
from .. import register_action
from docwen.i18n import t

logger = logging.getLogger(__name__)


@register_action("merge_images_to_tiff")
class MergeImagesToTiffStrategy(BaseStrategy):
    """
    将多个图片合并为单个多页TIFF文件的策略
    
    支持两种透明处理模式：
    1. smart（推荐）：智能判断
       - 如果所有图片都有透明通道，转为RGBA保留透明效果
       - 否则统一转为RGB，透明变白
    2. RGB：统一转为RGB模式，透明背景变白
    
    功能特点：
    - 按批量列表顺序合并（从上到下）
    - 保持原始尺寸
    - 输出到第一个文件所在目录
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行图片合并为TIFF
        
        Args:
            file_path: 第一个图片文件路径（用于确定输出目录）
            options: 合并选项
                - mode: "smart" 或 "RGB"（默认"smart"）
                - file_list: 要合并的文件路径列表（必需）
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 转换结果对象
        """
        try:
            # 获取选项
            mode = options.get('mode', 'smart') if options else 'smart'
            file_list = options.get('file_list', []) if options else []
            
            if not file_list:
                return ConversionResult(
                    success=False,
                    message=t('conversion.messages.no_images_to_merge')
                )
            
            logger.info(f"开始合并 {len(file_list)} 个图片为TIFF，模式: {mode}")
            
            if progress_callback:
                progress_callback(t('conversion.progress.preparing_merge_images', count=len(file_list)))
            
            # 输出目录为选中文件所在目录
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 步骤1：加载所有图片
            images = []
            has_alpha_list = []
            
            for idx, img_path in enumerate(file_list, 1):
                if progress_callback:
                    progress_callback(t('conversion.progress.loading_image_progress', current=idx, total=len(file_list)))
                
                try:
                    img = Image.open(img_path)
                    # 检测是否有透明通道
                    has_alpha = img.mode in ['RGBA', 'LA', 'PA']
                    has_alpha_list.append(has_alpha)
                    
                    images.append(img)
                    logger.debug(f"✓ 加载图片 {idx}: {os.path.basename(img_path)} (模式: {img.mode}, 透明: {has_alpha})")
                except Exception as e:
                    logger.warning(f"跳过无法加载的图片: {img_path}, 错误: {e}")
                    continue
            
            if not images:
                return ConversionResult(
                    success=False,
                    message=t('conversion.messages.no_images_loaded')
                )
            
            if progress_callback:
                progress_callback(t('conversion.progress.determining_color_mode'))
            
            # 步骤2：确定目标颜色模式
            if mode == 'smart':
                # 智能模式：如果所有图片都有透明，用RGBA；否则用RGB
                all_have_alpha = all(has_alpha_list)
                target_mode = 'RGBA' if all_have_alpha else 'RGB'
                logger.info(f"智能模式 - 所有图片透明: {all_have_alpha}, 目标模式: {target_mode}")
            else:
                # RGB模式：统一转RGB
                target_mode = 'RGB'
                logger.info(f"RGB模式 - 目标模式: {target_mode}")
            
            if progress_callback:
                progress_callback(t('conversion.progress.converting_to_mode', mode=target_mode))
            
            # 步骤3：转换所有图片到目标模式
            converted_images = []
            for idx, img in enumerate(images, 1):
                if img.mode != target_mode:
                    if target_mode == 'RGB' and img.mode in ['RGBA', 'LA', 'PA']:
                        # 转RGB：创建白色背景
                        logger.debug(f"图片 {idx}: {img.mode} → RGB (白色背景)")
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'PA':
                            img = img.convert('RGBA')
                        if img.mode in ['RGBA', 'LA']:
                            background.paste(img, mask=img.split()[-1])
                        else:
                            background.paste(img)
                        converted_images.append(background)
                    else:
                        # 其他模式转换
                        logger.debug(f"图片 {idx}: {img.mode} → {target_mode}")
                        converted_images.append(img.convert(target_mode))
                else:
                    logger.debug(f"图片 {idx}: 已是{target_mode}模式")
                    converted_images.append(img)
            
            if progress_callback:
                progress_callback(t('conversion.progress.generating_tiff_file'))
            
            # 步骤4：生成输出文件名（使用选中文件所在目录）
            selected_file = options.get("selected_file", file_path) if options else file_path
            output_dir = get_output_directory(selected_file)
            logger.info(f"输出目录基于: {os.path.basename(selected_file)}")
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"合并TIF_{timestamp}.tif"
            output_path = os.path.join(output_dir, output_filename)
            
            # 步骤5：保存为多页TIFF
            logger.info(f"保存多页TIFF文件: {output_filename}")
            converted_images[0].save(
                output_path,
                save_all=True,
                append_images=converted_images[1:] if len(converted_images) > 1 else [],
                compression='tiff_lzw'  # 使用LZW无损压缩
            )
            
            logger.info(f"✓ 成功合并 {len(converted_images)} 个图片为TIFF: {output_filename}")
            
            if progress_callback:
                progress_callback(t('conversion.progress.merge_completed'))
            
            return ConversionResult(
                success=True,
                output_path=output_path,
                message=t('conversion.messages.merge_images_to_tiff_success', count=len(converted_images))
            )
        
        except Exception as e:
            logger.error(f"合并图片为TIFF失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t('conversion.messages.conversion_failed_with_error', error=str(e)),
                error=e
            )
