"""
统一文件路由中心 (已重构)
职责：处理命令行用户交互，并调用服务层来执行核心业务逻辑。
"""
import os
import sys
import logging
from typing import Tuple, Optional, List, Type

# 导入服务层
from gongwen_converter.services.strategies.base_strategy import BaseStrategy
from gongwen_converter.services.strategies.md_to_document_strategies import MdToDocxStrategy, MdToDocStrategy
from gongwen_converter.services.strategies.md_to_spreadsheet_strategies import MdToXlsxStrategy, MdToXlsStrategy
from gongwen_converter.services.strategies.document_strategies import DocxToMdStrategy, DocxValidationStrategy
from gongwen_converter.services.strategies.spreadsheet_strategies import SpreadsheetToMarkdownStrategy

# 导入底层支持
from gongwen_converter.converter.formats.office import OfficeSoftwareNotFoundError
from gongwen_converter.template.loader import get_available_templates
from gongwen_converter.utils.file_type_utils import get_file_type_by_signature

# 配置日志
logger = logging.getLogger()

class FileRouter:
    """统一文件路由器 - 重构版"""
    
    def __init__(self, config=None):
        """
        初始化文件路由器
        :param config: 转换器配置字典 (可选)
        """
        self.config = config or {}
        logger.info("初始化文件路由器 - 重构版")
        logger.debug(f"配置文件: {self.config}")
        
        # 获取模板目录
        self.templates_dir = self._get_templates_dir()
        logger.info(f"最终模板目录: {self.templates_dir}")

    def route_file(self, file_path: str, show_success: bool = True, interactive: bool = True) -> Tuple[bool, Optional[str]]:
        """
        统一路由文件处理
        
        参数:
            file_path: 输入文件路径
            show_success: 是否显示成功信息
            interactive: 是否使用交互模式
            
        返回:
            tuple: (处理是否成功, 输出文件路径)
        """
        try:
            logger.info("=" * 60)
            logger.info(f"开始路由文件: {file_path}")
            
            # 验证文件
            if not self._validate_file(file_path):
                logger.error(f"文件验证失败: {file_path}")
                if show_success:
                    self._show_error(f"文件无效: {os.path.basename(file_path)}")
                return False, None
            
            # 获取文件扩展名
            ext = os.path.splitext(file_path)[1].lower()
            logger.info(f"文件类型: {ext}")

            # 检查文件实际类型（防止伪装的docx）
            actual_type = self._get_file_type(file_path)
            if actual_type == 'doc' and ext == '.docx':
                logger.warning(f"文件 {file_path} 后缀为docx但实际是doc文件，按doc文件处理")
                ext = '.doc'  # 将扩展名改为.doc，以便后续按doc处理

            # 处理DOC/WPS文件：转换为DOCX
            temp_docx_path = None
            if ext == '.doc' or ext == '.wps':
                file_type_name = "DOC" if ext == '.doc' else "WPS"
                print(f"检测到 {file_type_name} 文件，尝试转换为 DOCX...")
                try:
                    from gongwen_converter.converter.formats.office import office_to_docx
                    # 从扩展名推断格式
                    detected_format = 'doc' if ext == '.doc' else 'wps'
                    temp_docx_path = office_to_docx(
                        file_path,
                        actual_format=detected_format
                    )
                    file_path = temp_docx_path
                    ext = '.docx'
                    logger.info(f"{file_type_name}文件已转换为临时DOCX: {temp_docx_path}")
                except OfficeSoftwareNotFoundError as e:
                    logger.error(f"Office/WPS软件未找到: {e}")
                    self._show_error(str(e))
                    return False, None
                except Exception as e:
                    logger.error(f"{file_type_name}转DOCX失败: {str(e)}")
                    if show_success:
                        self._show_error(f"处理失败: {file_type_name}文件转换失败({str(e)})")
                    return False, None


            # 交互模式处理
            # 交互模式是CLI的唯一模式
            if ext in ['.xlsx', '.xls', '.et', '.csv']:
                return self._execute_strategy(file_path, show_success, SpreadsheetToMarkdownStrategy)
            elif ext == '.md' or ext == '.txt':
                return self._interactive_process_md(file_path, show_success)
            elif ext == '.docx':
                return self._interactive_process_docx(file_path, show_success)
            else:
                logger.error(f"不支持的文件类型: {ext}")
                self._show_error(f"不支持的文件类型: {ext}")
                return False, None
            
        except Exception as e:
            logger.error(f"文件路由过程中发生错误: {str(e)}", exc_info=True)
            if show_success:
                self._show_error(f"处理失败: {str(e)}")
            return False, None
        finally:
            logger.info("文件路由处理完成")
            logger.info("=" * 60)

    def _get_file_type(self, file_path: str) -> str:
        """
        通过文件签名判断文件实际类型
        检测伪装成DOCX的DOC文件
        
        参数:
            file_path: 文件路径
            
        返回:
            str: 文件实际类型 ('doc', 'docx', 'wps', 或 'unknown')
        """
        return get_file_type_by_signature(file_path)

    def _interactive_process_md(self, file_path: str, show_success: bool) -> Tuple[bool, Optional[str]]:
        """交互式处理MD文件（重构版）。"""
        print("\n" + "=" * 50)
        print(f"处理文件: {os.path.basename(file_path)}")
        print("请选择转换类型:\n1. 转换为 Word 文档 (DOCX/DOC)\n2. 转换为 Excel 报表 (XLSX/XLS)\n0. 取消操作")
        choice = input("请输入选项(1-2, 0取消): ").strip()

        if choice == '1':
            return self._handle_md_to_word_flow(file_path, show_success)
        elif choice == '2':
            return self._handle_md_to_spreadsheet_flow(file_path, show_success)
        else:
            print("操作已取消。")
            return False, None

    def _interactive_process_docx(self, file_path: str, show_success: bool) -> Tuple[bool, Optional[str]]:
        """交互式处理DOCX文件（重构版）。"""
        is_converted = "(由DOC转换)" if "from_doc" in os.path.basename(file_path) else ""
        print("\n" + "=" * 50)
        print(f"处理文件: {os.path.basename(file_path)} {is_converted}")
        print("请选择操作类型:\n1. 错别字检测\n2. 转换为Markdown报告\n0. 取消操作")
        choice = input("请输入选项(1-2, 0取消): ").strip()

        if choice == '1':
            options = {"spell_check_options": self._get_spell_check_option()}
            return self._execute_strategy(file_path, show_success, DocxValidationStrategy, options)
        elif choice == '2':
            return self._execute_strategy(file_path, show_success, DocxToMdStrategy)
        else:
            print("操作已取消。")
            return False, None

    def _handle_md_to_word_flow(self, file_path: str, show_success: bool) -> Tuple[bool, Optional[str]]:
        """处理MD到Word的完整交互流程（格式选择、模板、校对）。"""
        print("\n请选择输出格式:\n1. DOCX (Word文档)\n2. DOC (Word 97-2003文档)")
        format_choice = input("请输入选项(1-2, 默认1): ").strip() or "1"

        strategy_map = {'1': MdToDocxStrategy, '2': MdToDocStrategy}
        strategy_class = strategy_map.get(format_choice, MdToDocxStrategy)

        template_name = self._select_template('docx')
        if not template_name:
            return False, None

        spell_check_options = self._get_spell_check_option()
        
        options = {"template_name": template_name, "spell_check_options": spell_check_options}
        return self._execute_strategy(file_path, show_success, strategy_class, options)

    def _handle_md_to_spreadsheet_flow(self, file_path: str, show_success: bool) -> Tuple[bool, Optional[str]]:
        """处理MD到电子表格的完整交互流程。"""
        print("\n请选择输出格式:\n1. XLSX (Excel文档)\n2. XLS (Excel 97-2003文档)")
        format_choice = input("请输入选项(1-2, 默认1): ").strip() or "1"

        strategy_map = {'1': MdToXlsxStrategy, '2': MdToXlsStrategy}
        strategy_class = strategy_map.get(format_choice, MdToXlsxStrategy)

        template_name = self._select_template('xlsx')
        if not template_name:
            return False, None
        
        options = {"template_name": template_name}
        return self._execute_strategy(file_path, show_success, strategy_class, options)

    def _execute_strategy(self, file_path: str, show_success: bool, strategy_class: Type[BaseStrategy], options: Optional[dict] = None) -> Tuple[bool, Optional[str]]:
        """统一执行策略并处理结果的辅助函数。"""
        # 在执行任何操作前，检查文件是否存在
        if not os.path.exists(file_path):
            self._show_error(f"文件 '{os.path.basename(file_path)}' 已被移动或删除，请重新操作。")
            return False, None

        def progress_callback(message: str):
            print(f"... {message}")

        try:
            strategy = strategy_class()
            result = strategy.execute(
                file_path=file_path,
                options=options,
                progress_callback=progress_callback
            )

            if result.success and show_success:
                self._show_success(result.message or "操作成功！")
                if result.output_path:
                    self._open_output_directory(result.output_path)
            elif not result.success:
                self._show_error(result.message or "操作失败！")
            
            return result.success, result.output_path

        except Exception as e:
            logger.error(f"执行策略 {strategy_class.__name__} 时出错: {e}", exc_info=True)
            self._show_error(f"发生严重错误: {e}")
            return False, None
    
    def _select_template(self, template_type: str) -> Optional[str]:
        """让用户从可用模板列表中选择一个。"""
        templates = self._get_available_templates(template_type)
        if not templates:
            print(f"错误: 没有找到可用的{template_type.upper()}模板。")
            return None
        
        print(f"\n可用{template_type.upper()}模板:")
        for i, template in enumerate(templates, 1):
            print(f"{i}. {template}")
        
        choice = input(f"请选择模板(1-{len(templates)}, 默认1): ").strip() or "1"
        try:
            choice_index = int(choice) - 1
            selected_template = templates[choice_index] if 0 <= choice_index < len(templates) else templates[0]
        except (ValueError, IndexError):
            selected_template = templates[0]
        
        print(f"已选择模板: {selected_template}")
        return selected_template

    def _get_spell_check_option(self) -> int:
        """获取用户选择的错别字检查规则。"""
        print("\n" + "=" * 50)
        print("请选择错别字检查规则(可输入数字组合, 或回车使用默认配置):")
        print(" 1=标点配对, 2=符号校对, 4=自定义, 8=智能校对 (示例: 3 表示 1+2)")
        print("输入0不检查, 直接回车使用配置文件默认值。")
        user_input = input("请选择规则(0-15): ").strip()
        
        if not user_input:
            print("将使用配置文件默认设置。")
            return -1 # 使用-1等特殊值表示默认，由策略内部处理
        try:
            return int(user_input)
        except ValueError:
            print("输入无效，将使用配置文件默认设置。")
            return -1

    def _get_available_templates(self, template_type: str) -> List[str]:
        """
        获取可用模板列表 - 带错误处理
        
        参数:
            template_type: 模板类型 ('docx' 或 'xlsx')
            
        返回:
            list: 可用模板名称列表
        """
        try:
            templates = get_available_templates(template_type)
            logger.info(f"找到 {len(templates)} 个{template_type.upper()}模板")
            return templates
        except Exception as e:
            logger.error(f"获取模板列表失败: {str(e)}")
            return []

    def _get_templates_dir(self) -> str:
        """
        获取模板目录
        
        返回:
            str: 模板目录路径
        """
        try:
            from gongwen_converter.template.loader import TemplateLoader
            loader = TemplateLoader()
            return loader.get_default_template_dir()
        except ImportError:
            logger.error("模板加载器不可用，使用默认目录")
            return "templates"
        except Exception as e:
            logger.error(f"获取模板目录失败: {str(e)}")
            return "templates"

    def _validate_file(self, file_path: str) -> bool:
        """
        统一文件验证
        
        参数:
            file_path: 文件路径
            
        返回:
            bool: 文件是否有效
        """
        logger.debug(f"验证文件: {file_path}")
        
        if not os.path.exists(file_path):
            logger.error(f"文件不存在: {file_path}")
            return False
            
        if not os.path.isfile(file_path):
            logger.error(f"路径不是文件: {file_path}")
            return False
            
        if os.path.getsize(file_path) == 0:
            logger.error(f"空文件: {file_path}")
            return False
            
        logger.debug("文件验证通过")
        return True

    def _show_success(self, message: str):
        """统一显示成功信息"""
        logger.info(f"显示成功消息: {message}")
        print(f"\n{message}\n{'=' * 60}")

    def _show_error(self, message: str):
        """统一显示错误信息"""
        logger.error(f"显示错误消息: {message}")
        print(f"\n错误: {message}\n{'=' * 60}")

    def _open_output_directory(self, file_path: str):
        """统一打开输出目录"""
        output_dir = os.path.dirname(file_path)
        logger.info(f"尝试打开输出目录: {output_dir}")
        
        try:
            if sys.platform == 'win32':
                os.startfile(output_dir)
            elif sys.platform == 'darwin':
                import subprocess
                subprocess.Popen(['open', output_dir])
            else:
                import subprocess
                subprocess.Popen(['xdg-open', output_dir])
        except Exception as e:
            logger.warning(f"无法打开输出目录: {e}")
            print(f"提示: 输出文件位于 {output_dir}")

# 全局单例实例 (延迟加载)
global_router: Optional[FileRouter] = None

def route_file(file_path: str, show_success: bool = True, interactive: bool = True) -> Tuple[bool, Optional[str]]:
    """
    全局文件路由函数 - 使用延迟加载单例模式
    
    参数:
        file_path: 输入文件路径
        show_success: 是否显示成功信息
        interactive: 是否使用交互模式
        
    返回:
        tuple: (处理是否成功, 输出文件路径)
    """
    global global_router
    
    # 延迟加载：在第一次调用时才创建实例，确保配置已加载
    if global_router is None:
        try:
            from gongwen_converter.config.config_manager import config_manager
            # 传递转换配置，以便策略在未来可以访问
            conv_config = config_manager.get_conversion_config_block()
            global_router = FileRouter(config=conv_config)
            logger.info("CLI 文件路由器已延迟初始化并加载了配置。")
        except Exception as e:
            logger.critical(f"延迟初始化CLI路由器失败: {e}", exc_info=True)
            # 即使失败，也创建一个无配置的实例以尝试继续
            global_router = FileRouter()

    return global_router.route_file(file_path, show_success, interactive)
