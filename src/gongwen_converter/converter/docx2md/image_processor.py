"""
DOCX图片提取模块

从DOCX文档中提取内嵌图片并保存到指定文件夹
"""

import os
import logging
from typing import List, Dict
from gongwen_converter.utils.path_utils import generate_output_path

logger = logging.getLogger(__name__)

# 支持的图片格式（真实的外部图片）
SUPPORTED_IMAGE_FORMATS = {
    'jpeg', 'jpg', 'png', 'gif', 'tiff', 'tif', 'bmp'
}

# 忽略的格式（Office绘图形状）
IGNORED_FORMATS = {
    'emf', 'wmf'
}


def extract_images_from_docx(doc, output_folder: str, original_file_path: str) -> List[Dict]:
    """
    从DOCX文档中提取所有图片
    
    参数:
        doc: Document对象（python-docx）
        output_folder: 图片保存的文件夹路径
        original_file_path: 原始DOCX文件路径（用于生成规范的图片文件名）
        
    返回:
        List[Dict]: 图片信息列表，每个元素包含：
            - para_index: 图片所在的段落索引
            - filename: 图片文件名（如 '报告_image1_20250125_091500_fromDocx.png'）
            - image_path: 图片完整路径
    """
    images_info = []
    image_counter = 1
    
    logger.info(f"开始提取DOCX中的图片到: {output_folder}")
    
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 遍历所有段落
    for para_idx, para in enumerate(doc.paragraphs):
        # 查找段落中的所有drawing元素（图片）
        try:
            drawings = para._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
            
            for drawing in drawings:
                # 查找blip元素（图片引用）
                blips = drawing.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                
                for blip in blips:
                    # 获取图片关系ID
                    embed_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
                    rId = blip.get(embed_attr)
                    
                    if not rId:
                        logger.debug(f"段落 {para_idx + 1} 中的图片没有embed属性，跳过")
                        continue
                    
                    try:
                        # 通过关系ID获取图片部分
                        image_part = doc.part.related_parts[rId]
                        
                        # 获取图片内容类型（如 'image/jpeg'）
                        content_type = image_part.content_type
                        ext = content_type.split('/')[-1] if '/' in content_type else 'png'
                        
                        # 标准化扩展名
                        ext = ext.lower()
                        
                        # 检查格式是否支持
                        if ext in SUPPORTED_IMAGE_FORMATS:
                            # 使用统一的路径生成工具生成规范文件名
                            image_path = generate_output_path(
                                input_path=original_file_path,
                                output_dir=output_folder,
                                section=f"image{image_counter}",
                                add_timestamp=True,
                                description="fromDocx",
                                file_type=ext
                            )
                            filename = os.path.basename(image_path)
                            
                            # 保存图片
                            with open(image_path, 'wb') as f:
                                f.write(image_part.blob)
                            
                            logger.debug(f"提取图片: {filename} (段落 {para_idx + 1})")
                            
                            # 记录图片信息
                            images_info.append({
                                'para_index': para_idx,
                                'filename': filename,
                                'image_path': image_path
                            })
                            
                            image_counter += 1
                            
                        elif ext in IGNORED_FORMATS:
                            logger.debug(f"忽略Office形状格式: {ext} (段落 {para_idx + 1})")
                        else:
                            logger.warning(f"未知图片格式: {ext} (段落 {para_idx + 1})")
                    
                    except KeyError:
                        logger.warning(f"无法找到关系ID为 {rId} 的图片部分")
                    except Exception as e:
                        logger.error(f"提取图片时出错 (段落 {para_idx + 1}): {e}")
        
        except Exception as e:
            logger.error(f"处理段落 {para_idx + 1} 时出错: {e}")
            continue
    
    logger.info(f"图片提取完成，共提取 {len(images_info)} 张图片")
    
    return images_info


# 模块测试
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s"
    )
    
    logger.info("=== DOCX图片提取模块测试 ===\n")
    
    # 测试文件
    test_docx = "test.docx"
    
    if os.path.exists(test_docx):
        from docx import Document
        import tempfile
        
        try:
            doc = Document(test_docx)
            logger.info(f"成功加载文档: {test_docx}")
            
            with tempfile.TemporaryDirectory() as temp_dir:
                output_folder = os.path.join(temp_dir, "test_images")
                
                images_info = extract_images_from_docx(doc, output_folder, test_docx)
                
                logger.info(f"\n提取结果:")
                logger.info(f"  总计: {len(images_info)} 张图片")
                
                for img in images_info:
                    logger.info(f"  - {img['filename']} (段落 {img['para_index'] + 1})")
                    logger.info(f"    路径: {img['image_path']}")
                    logger.info(f"    文件存在: {os.path.exists(img['image_path'])}")
        
        except Exception as e:
            logger.error(f"测试失败: {e}", exc_info=True)
    else:
        logger.warning(f"测试文件不存在: {test_docx}")
    
    logger.info("\n=== 测试结束 ===")
