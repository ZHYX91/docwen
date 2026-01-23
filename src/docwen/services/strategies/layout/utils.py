"""
版式文件处理工具函数

提供版式文件（PDF/OFD/XPS/CAJ）的预处理功能，
包括格式检测、临时文件创建和格式转换。

依赖：
- converter.formats.layout: 版式格式转换（OFD/XPS/CAJ → PDF）
- utils.file_type_utils: 文件格式检测
- utils.workspace_manager: 临时文件管理
"""

import os
import logging
from typing import Optional, Tuple
import threading

logger = logging.getLogger(__name__)


def preprocess_layout_file(
    file_path: str,
    temp_dir: str,
    cancel_event: Optional[threading.Event] = None,
    actual_format: Optional[str] = None
) -> Tuple[str, Optional[str]]:
    """
    预处理版式文件：创建输入副本，并将非标准格式转换为PDF
    
    参数:
        file_path: 原始文件路径
        temp_dir: 临时目录路径，转换后的中间文件将输出到此目录
        cancel_event: 用于取消操作的事件对象
        actual_format: 实际文件格式（可选，如果不提供则自动检测）
        
    返回:
        tuple: (处理后的PDF路径, 中间文件路径或None)
            - 所有格式都会先创建 input.{ext} 副本
            - 如果需要转换（XPS/CAJ/OFD），返回转换后的PDF路径和中间PDF文件路径
            - 如果已是标准格式（PDF），返回副本路径和None
            
    说明:
        - 使用actual_format参数避免重复检测文件格式
        - 所有中间文件都输出到temp_dir，由调用者的上下文管理器统一清理
        - 返回的中间文件路径可用于保留中间文件功能
    """
    # 如果没有提供actual_format，则检测
    if actual_format is None:
        from docwen.utils.file_type_utils import detect_actual_file_format
        actual_format = detect_actual_file_format(file_path)
        logger.debug(f"自动检测版式文件格式: {actual_format}")
    else:
        logger.debug(f"使用传入的文件格式: {actual_format}")
    
    # 步骤1：无论什么格式，都先创建输入副本 input.{ext}
    from docwen.utils.workspace_manager import prepare_input_file
    temp_input = prepare_input_file(file_path, temp_dir, actual_format)
    logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
    
    # 步骤2：如果是标准格式（PDF），直接返回副本路径，无中间文件
    if actual_format == 'pdf':
        logger.debug(f"文件已是PDF格式，返回副本路径")
        return temp_input, None
    
    # 步骤3：需要转换的格式 - CAJ
    if actual_format == 'caj':
        logger.info(f"检测到CAJ格式，从副本转换为PDF: {os.path.basename(temp_input)}")
        
        try:
            from docwen.converter.formats.layout import caj_to_pdf
            
            # 调用转换函数，使用副本作为输入
            converted_path = caj_to_pdf(temp_input, cancel_event, output_dir=temp_dir)
            
            if converted_path:
                logger.info(f"CAJ转PDF成功: {os.path.basename(converted_path)}")
                return converted_path, converted_path  # 返回PDF路径和中间文件路径
            else:
                logger.error("CAJ转PDF失败，返回None")
                raise RuntimeError("CAJ转PDF失败")
                
        except NotImplementedError as e:
            logger.error(f"CAJ转PDF功能尚未实现: {e}")
            raise RuntimeError(f"CAJ转PDF功能尚未实现: {e}")
        except Exception as e:
            logger.error(f"CAJ转PDF失败: {e}")
            raise RuntimeError(f"CAJ转PDF失败: {e}")
    
    # 步骤4：需要转换的格式 - XPS
    elif actual_format == 'xps':
        logger.info(f"检测到XPS格式，从副本转换为PDF: {os.path.basename(temp_input)}")
        
        try:
            from docwen.converter.formats.layout import xps_to_pdf
            
            # 调用转换函数，使用副本作为输入
            converted_path = xps_to_pdf(temp_input, cancel_event, output_dir=temp_dir)
            
            if converted_path:
                logger.info(f"XPS转PDF成功: {os.path.basename(converted_path)}")
                return converted_path, converted_path  # 返回PDF路径和中间文件路径
            else:
                logger.error("XPS转PDF失败，返回None")
                raise RuntimeError("XPS转PDF失败")
                
        except Exception as e:
            logger.error(f"XPS转PDF失败: {e}")
            raise RuntimeError(f"XPS转PDF失败: {e}")
    
    # 步骤5：需要转换的格式 - OFD
    elif actual_format == 'ofd':
        logger.info(f"检测到OFD格式，从副本转换为PDF: {os.path.basename(temp_input)}")
        
        try:
            from docwen.converter.formats.layout import ofd_to_pdf
            
            # 调用转换函数，使用副本作为输入
            converted_path = ofd_to_pdf(temp_input, cancel_event, output_dir=temp_dir)
            
            if converted_path:
                logger.info(f"OFD转PDF成功: {os.path.basename(converted_path)}")
                return converted_path, converted_path  # 返回PDF路径和中间文件路径
            else:
                logger.error("OFD转PDF失败，返回None")
                raise RuntimeError("OFD转PDF失败")
                
        except NotImplementedError as e:
            logger.error(f"OFD转PDF功能尚未实现: {e}")
            raise RuntimeError(f"OFD转PDF功能尚未实现: {e}")
        except Exception as e:
            logger.error(f"OFD转PDF失败: {e}")
            raise RuntimeError(f"OFD转PDF失败: {e}")
    
    # 步骤6：其他不支持的格式
    else:
        logger.warning(f"不支持的版式文件格式: {actual_format}")
        raise RuntimeError(f"不支持的版式文件格式: {actual_format}")


def should_keep_intermediates() -> bool:
    """
    判断是否应该保留中间文件
    
    返回:
        bool: True表示保留中间文件，False表示不保留
    """
    try:
        from docwen.config.config_manager import config_manager
        return config_manager.get_save_intermediate_files()
    except Exception as e:
        logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
        return False
