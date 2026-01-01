"""
MD文件序号处理策略

提供纯Markdown文件的小标题序号处理功能：
- 清除原有序号
- 添加新序号（支持多种方案）
- 规范化序号（先清除后添加）

适用于CLI的 process_md_numbering action 和 Obsidian插件的序号处理命令。
"""

import os
import logging
from typing import Optional, Callable, Dict, Any

from gongwen_converter.services.strategies.base_strategy import BaseStrategy
from gongwen_converter.services.strategies import register_action
from gongwen_converter.services.result import ConversionResult

logger = logging.getLogger(__name__)


@register_action('process_md_numbering')
class MdNumberingStrategy(BaseStrategy):
    """
    MD文件序号处理策略
    
    用于纯MD文件的序号清理和添加，不转换格式。
    
    支持的选项:
        - remove_numbering: bool - 是否清除原有序号
        - add_numbering: bool - 是否新增序号
        - numbering_scheme: str - 序号方案ID
            - 'gongwen_standard': 公文标准（一、（一）1.（1）①）
            - 'hierarchical_standard': 层级数字标准（1 1.1 1.1.1）
            - 'legal_standard': 法律条文标准（第一编 第一章 第一节 第一条）
    
    使用示例:
        strategy = MdNumberingStrategy()
        result = strategy.execute(
            file_path='document.md',
            options={
                'remove_numbering': True,
                'add_numbering': True,
                'numbering_scheme': 'gongwen_standard'
            }
        )
    """
    
    # 策略元信息
    name = "MD序号处理"
    description = "处理Markdown文件的小标题序号（清除/添加/规范化）"
    
    # 支持的源格式和目标格式
    supported_source_formats = ['md', 'markdown']
    supported_target_formats = ['md', 'markdown']
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行MD文件序号处理
        
        Args:
            file_path: MD文件路径
            options: 选项字典
                - remove_numbering: 是否清除原有序号
                - add_numbering: 是否新增序号
                - numbering_scheme: 序号方案ID
            progress_callback: 进度回调
            
        Returns:
            ConversionResult: 处理结果
        """
        options = options or {}
        
        # 获取选项
        remove_numbering = options.get('remove_numbering', False)
        add_numbering = options.get('add_numbering', False)
        numbering_scheme = options.get('numbering_scheme', 'gongwen_standard')
        
        # 验证：至少需要执行一个操作
        if not remove_numbering and not add_numbering:
            return ConversionResult(
                success=False,
                message="请至少指定一个操作：--remove-numbering 或 --add-numbering"
            )
        
        logger.info(f"开始处理MD序号: {file_path}")
        logger.debug(f"选项: remove={remove_numbering}, add={add_numbering}, scheme={numbering_scheme}")
        
        try:
            # 验证文件存在
            if not os.path.exists(file_path):
                return ConversionResult(
                    success=False,
                    message=f"文件不存在: {file_path}"
                )
            
            # 验证文件扩展名
            ext = os.path.splitext(file_path)[1].lower()
            if ext not in ['.md', '.markdown']:
                return ConversionResult(
                    success=False,
                    message=f"不支持的文件格式: {ext}，仅支持 .md/.markdown 文件"
                )
            
            # 读取MD文件
            if progress_callback:
                progress_callback("读取文件...")
            
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            original_length = len(content)
            logger.debug(f"原始内容长度: {original_length} 字符")
            
            # 导入序号处理函数
            from gongwen_converter.utils.heading_numbering import process_md_numbering
            
            # 处理序号
            if progress_callback:
                progress_callback("处理序号...")
            
            processed_content = process_md_numbering(
                content=content,
                remove_existing=remove_numbering,
                add_new=add_numbering,
                scheme_id=numbering_scheme
            )
            
            processed_length = len(processed_content)
            logger.debug(f"处理后内容长度: {processed_length} 字符")
            
            # 写回文件
            if progress_callback:
                progress_callback("保存文件...")
            
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(processed_content)
            
            # 构建成功消息
            operations = []
            if remove_numbering:
                operations.append("去除序号")
            if add_numbering:
                operations.append(f"添加序号({numbering_scheme})")
            
            message = f"序号处理完成: {', '.join(operations)}"
            
            logger.info(f"序号处理完成: {file_path}")
            
            return ConversionResult(
                success=True,
                message=message,
                output_path=file_path,
                metadata={
                    'original_length': original_length,
                    'processed_length': processed_length,
                    'operations': operations,
                    'scheme': numbering_scheme if add_numbering else None
                }
            )
            
        except UnicodeDecodeError as e:
            logger.error(f"文件编码错误: {e}")
            return ConversionResult(
                success=False,
                message=f"文件编码错误，请确保文件为UTF-8编码: {e}"
            )
        
        except PermissionError as e:
            logger.error(f"文件权限错误: {e}")
            return ConversionResult(
                success=False,
                message=f"无法写入文件，请检查文件权限: {e}"
            )
        
        except Exception as e:
            logger.error(f"序号处理失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"序号处理失败: {str(e)}"
            )
    
    def validate_options(self, options: Dict[str, Any]) -> bool:
        """
        验证选项是否有效
        
        Args:
            options: 选项字典
            
        Returns:
            bool: 选项是否有效
        """
        # 验证序号方案
        valid_schemes = ['gongwen_standard', 'hierarchical_standard', 'legal_standard']
        scheme = options.get('numbering_scheme', 'gongwen_standard')
        
        if scheme not in valid_schemes:
            logger.warning(f"无效的序号方案: {scheme}，有效方案: {valid_schemes}")
            return False
        
        return True
