"""
公文转换器模板加载器模块

本模块负责模板文件的加载、缓存和管理，提供统一的模板访问接口。
主要功能包括：
- 支持DOCX和XLSX两种格式的模板加载
- 智能模板缓存机制，提高加载性能
- 模板文件的模糊匹配和智能查找
- 自动检测模板文件变更并更新缓存
- 支持开发和生产环境的多路径适配

"""

import os
import sys
import logging
import hashlib
import time
import copy
from typing import List
from docx import Document
from openpyxl import load_workbook
from docwen.utils.path_utils import ensure_dir_exists, get_project_root

# 配置日志
logger = logging.getLogger(__name__)

# 全局模板缓存
_TEMPLATE_CACHE = {}
_TEMPLATE_MTIMES = {}  # 记录模板文件的最后修改时间

class TemplateLoader:
    """模板加载器核心类"""
    
    def __init__(self, template_dir=None):
        """
        初始化模板加载器
        :param template_dir: 自定义模板目录（可选）
        """
        self.template_dir = template_dir or self.get_default_template_dir()
        logger.info(f"初始化模板加载器 | 模板目录: {self.template_dir}")
        self._validate_template_dir()
        self._refresh_cache()
    
    def get_default_template_dir(self) -> str:
        """
        获取默认模板目录路径
        """
        # 获取项目根目录
        project_root = get_project_root()
        logger.debug(f"项目根目录: {project_root}")
        
        # 开发环境：项目根目录的templates子目录
        templates_dir = os.path.join(project_root, "templates")
        
        # 如果模板目录存在，直接使用
        if os.path.exists(templates_dir):
            logger.debug(f"使用开发环境模板目录: {templates_dir}")
            return templates_dir
        
        # 打包环境：可执行文件同级的templates目录
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
            templates_dir = os.path.join(exe_dir, "templates")
            if os.path.exists(templates_dir):
                logger.debug(f"使用生产环境模板目录: {templates_dir}")
                return templates_dir
        
        # 回退到用户目录
        user_dir = os.path.expanduser("~")
        templates_dir = os.path.join(user_dir, ".docwen", "templates")
        ensure_dir_exists(templates_dir)
        logger.warning(f"使用用户模板目录: {templates_dir}")
        return templates_dir
    
    def _validate_template_dir(self):
        """验证模板目录是否存在"""
        if not os.path.exists(self.template_dir):
            logger.error(f"模板目录不存在: {self.template_dir}")
            raise FileNotFoundError(f"模板目录不存在: {self.template_dir}")
        
        if not os.path.isdir(self.template_dir):
            logger.error(f"模板路径不是目录: {self.template_dir}")
            raise NotADirectoryError(f"模板路径不是目录: {self.template_dir}")
        
        logger.info(f"模板目录验证通过: {self.template_dir}")
    
    def get_template_path(self, template_type: str, template_name: str = None) -> str:
        """
        获取模板完整路径
        支持自动添加文件扩展名
        
        :param template_type: 模板类型 ('docx' 或 'xlsx')
        :param template_name: 自定义模板名 (可选)
        :return: 模板完整路径
        """
        # 未提供模板名时使用默认模板
        if not template_name:
            template_name = self._get_default_template_name(template_type)
            logger.debug(f"使用默认模板: {template_name}")
        
        # 确保模板名称有扩展名
        if not template_name.lower().endswith(f".{template_type}"):
            template_name += f".{template_type}"
        
        # 构建完整路径
        template_path = os.path.join(self.template_dir, template_name)
        
        # 检查模板是否存在
        if not os.path.exists(template_path):
            logger.error(f"模板文件不存在: {template_path}")
            raise FileNotFoundError(f"模板文件不存在: {template_path}")
        
        logger.info(f"获取模板路径: {template_path}")
        return template_path
    
    def _get_default_template_name(self, template_type: str) -> str:
        """获取默认模板名称"""
        if template_type == "docx":
            return "公文通用"
        elif template_type == "xlsx":
            return "报表通用"
        else:
            logger.error(f"不支持的模板类型: {template_type}")
            raise ValueError(f"不支持的模板类型: {template_type}")
        
    def load_docx_template(self, template_name: str = None) -> Document:
        """
        加载DOCX模板
        
        参数:
            template_name: 自定义模板名 (可选)
            
        返回:
            Document对象（深拷贝）
        """
        try:
            # 获取模板路径
            template_path = self.get_template_path("docx", template_name)
            logger.debug(f"DOCX模板路径: {template_path}")
            
            # 生成缓存键
            cache_key = self._generate_cache_key("docx", template_name)
            logger.debug(f"缓存键: {cache_key}")
            
            # 检查文件是否修改
            current_mtime = os.path.getmtime(template_path)
            last_mtime = _TEMPLATE_MTIMES.get(cache_key, 0)
            logger.debug(f"文件修改时间检查 | 当前: {current_mtime} | 上次: {last_mtime}")
            
            # 文件已修改或缓存不存在时重新加载
            if current_mtime > last_mtime or cache_key not in _TEMPLATE_CACHE:
                if cache_key in _TEMPLATE_CACHE:
                    logger.info(f"模板已修改或缓存过期，重新加载: {template_path}")
                else:
                    logger.info(f"模板未缓存，首次加载: {template_path}")
                    
                # 清除旧缓存
                if cache_key in _TEMPLATE_CACHE:
                    del _TEMPLATE_CACHE[cache_key]
                    logger.debug(f"清除旧缓存: {cache_key}")
                
                # 加载新模板
                logger.debug(f"加载模板文件: {template_path}")
                doc = Document(template_path)
                
                # 更新缓存和修改时间
                _TEMPLATE_CACHE[cache_key] = doc
                _TEMPLATE_MTIMES[cache_key] = current_mtime
                logger.info(f"缓存更新 | 键: {cache_key} | 大小: {os.path.getsize(template_path)}字节")
            else:
                logger.debug(f"使用缓存DOCX模板: {cache_key}")
                doc = _TEMPLATE_CACHE[cache_key]
            
            # 返回模板的深拷贝，避免修改缓存中的原始模板
            logger.debug("创建模板的深拷贝...")
            try:
                # 尝试深拷贝
                doc_copy = copy.deepcopy(doc)
                logger.debug("深拷贝创建成功")
                return doc_copy
            except Exception as copy_error:
                # 深拷贝失败时回退到浅拷贝
                logger.warning(f"深拷贝失败: {str(copy_error)}，使用浅拷贝")
                return copy.copy(doc)
                
        except FileNotFoundError as e:
            logger.error(f"模板文件不存在: {template_path}")
            raise RuntimeError(f"模板文件不存在: {os.path.basename(template_path)}") from e
        except PermissionError as e:
            logger.error(f"无权限访问模板文件: {template_path}")
            raise RuntimeError("无权限访问模板文件") from e
        except Exception as e:
            error_str = str(e).lower()
            # 检测是否为无效的 DOCX 文件（不是有效的 ZIP 压缩包）
            if "not a zip file" in error_str or "package not found" in error_str:
                from docwen.i18n import t
                template_basename = os.path.basename(template_path)
                friendly_msg = t('messages.errors.invalid_docx_template', template_name=template_basename)
                logger.error(f"模板文件无效或损坏: {template_path}")
                raise RuntimeError(friendly_msg) from e
            logger.error(f"加载DOCX模板失败: {str(e)}", exc_info=True)
            raise RuntimeError(f"加载模板失败: {str(e)}") from e
        
    def load_xlsx_template(self, template_name: str = None):
        """
        加载Excel模板
        
        :param template_name: 自定义模板名 (可选)
        :return: Workbook对象
        """
        try:
            # 获取模板路径
            template_path = self.get_template_path("xlsx", template_name)
            logger.debug(f"Excel模板路径: {template_path}")
            
            # 生成缓存键
            cache_key = self._generate_cache_key("xlsx", template_name)
            
            # 检查文件是否修改
            current_mtime = os.path.getmtime(template_path)
            last_mtime = _TEMPLATE_MTIMES.get(cache_key, 0)
            
            # 文件已修改或缓存不存在时重新加载
            if current_mtime > last_mtime or cache_key not in _TEMPLATE_CACHE:
                logger.info(f"模板已修改或缓存过期，重新加载: {template_path}")
                
                # 清除旧缓存
                if cache_key in _TEMPLATE_CACHE:
                    del _TEMPLATE_CACHE[cache_key]
                
                # 加载新模板
                wb = load_workbook(template_path)
                
                # 更新缓存和修改时间
                _TEMPLATE_CACHE[cache_key] = wb
                _TEMPLATE_MTIMES[cache_key] = current_mtime
                logger.debug(f"缓存更新 | 键: {cache_key} | 大小: {os.path.getsize(template_path)}字节")
            else:
                logger.debug(f"使用缓存Excel模板: {cache_key}")
                wb = _TEMPLATE_CACHE[cache_key]
            
            return wb
        except Exception as e:
            logger.error(f"加载Excel模板失败: {str(e)}")
            raise
    
    def _generate_cache_key(self, template_type: str, template_name: str = None) -> str:
        """生成模板缓存键（唯一标识）"""
        if not template_name:
            template_name = self._get_default_template_name(template_type)
        
        # 使用哈希确保键值唯一
        key_str = f"{template_type}:{template_name}".lower()
        return hashlib.md5(key_str.encode()).hexdigest()
    
    def get_available_templates(self, template_type: str, name_only: bool = True) -> List[str]:
        """
        获取指定类型的所有可用模板
        支持只返回名称（不带扩展名）
        
        :param template_type: 模板类型 ('docx' 或 'xlsx')
        :param name_only: 是否只返回名称（不带扩展名）
        :return: 模板名称列表
        """
        logger.info(f"获取可用{template_type.upper()}模板...")
        
        # 检查模板目录是否存在
        if not os.path.exists(self.template_dir):
            logger.warning(f"模板目录不存在: {self.template_dir}")
            return []
        
        # 获取指定类型的模板文件
        extension = ".docx" if template_type == "docx" else ".xlsx"
        templates = []
        
        for f in os.listdir(self.template_dir):
            file_path = os.path.join(self.template_dir, f)
            if os.path.isfile(file_path) and f.lower().endswith(extension):
                # 移除扩展名（如果指定）
                if name_only:
                    name = os.path.splitext(f)[0]
                else:
                    name = f
                templates.append(name)
                logger.debug(f"找到模板: {name}")
        
        # 按字母排序
        templates.sort()
        logger.info(f"找到 {len(templates)} 个{template_type.upper()}模板")
        return templates

    def get_template_list(self, template_type: str) -> List[str]:
        """
        获取指定类型的模板列表（GUI专用）
        
        :param template_type: 模板类型 ('docx' 或 'xlsx')
        :return: 模板名称列表
        """
        return self.get_available_templates(template_type)
    
    def find_template(self, template_type: str, search_term: str) -> str:
        """
        模糊查找模板
        
        :param template_type: 模板类型 ('docx' 或 'xlsx')
        :param search_term: 搜索关键词
        :return: 匹配的模板名称（找不到时返回空字符串）
        """
        # 获取所有模板
        templates = self.get_available_templates(template_type)
        
        # 完全匹配
        if search_term in templates:
            logger.debug(f"完全匹配模板: {search_term}")
            return search_term
        
        # 模糊匹配
        matches = []
        for template in templates:
            if search_term.lower() in template.lower():
                matches.append(template)
        
        # 单个匹配
        if len(matches) == 1:
            logger.debug(f"模糊匹配模板: {matches[0]}")
            return matches[0]
        
        # 多个匹配
        if matches:
            logger.debug(f"找到多个匹配模板: {matches}")
            return matches[0]  # 返回第一个匹配项
        
        # 无匹配
        logger.warning(f"未找到匹配模板: {search_term}")
        return ""
    
    def clear_cache(self, template_type: str = None, template_name: str = None):
        """
        清空模板缓存（可指定类型或特定模板）
        
        :param template_type: 模板类型 ('docx' 或 'xlsx')，可选
        :param template_name: 模板名称，可选
        """
        global _TEMPLATE_CACHE, _TEMPLATE_MTIMES
        
        if template_type and template_name:
            # 清除特定模板缓存
            cache_key = self._generate_cache_key(template_type, template_name)
            if cache_key in _TEMPLATE_CACHE:
                del _TEMPLATE_CACHE[cache_key]
            if cache_key in _TEMPLATE_MTIMES:
                del _TEMPLATE_MTIMES[cache_key]
            logger.info(f"清除特定模板缓存: {template_type}/{template_name}")
        elif template_type:
            # 清除特定类型的所有缓存
            keys_to_remove = [k for k in _TEMPLATE_CACHE.keys() if k.startswith(template_type)]
            for key in keys_to_remove:
                del _TEMPLATE_CACHE[key]
                if key in _TEMPLATE_MTIMES:
                    del _TEMPLATE_MTIMES[key]
            logger.info(f"清除{template_type.upper()}类型所有模板缓存")
        else:
            # 清空所有缓存
            cache_size = len(_TEMPLATE_CACHE)
            _TEMPLATE_CACHE = {}
            _TEMPLATE_MTIMES = {}
            logger.info(f"清空所有模板缓存 | 已移除 {cache_size} 个模板")
    
    def refresh_cache(self):
        """刷新缓存（检查所有模板的修改时间）"""
        logger.info("开始刷新模板缓存...")
        
        # 获取所有模板路径
        docx_templates = self.get_available_templates("docx", name_only=False)
        xlsx_templates = self.get_available_templates("xlsx", name_only=False)
        all_templates = [(name, "docx") for name in docx_templates] + \
                        [(name, "xlsx") for name in xlsx_templates]
        
        refreshed_count = 0
        
        for name, t_type in all_templates:
            try:
                # 获取模板路径
                template_path = self.get_template_path(t_type, name)
                
                # 获取当前修改时间
                current_mtime = os.path.getmtime(template_path)
                
                # 生成缓存键
                cache_key = self._generate_cache_key(t_type, name)
                
                # 检查是否需要刷新
                if cache_key in _TEMPLATE_MTIMES and current_mtime > _TEMPLATE_MTIMES[cache_key]:
                    logger.info(f"模板已修改: {name}.{t_type}")
                    
                    # 清除缓存
                    if cache_key in _TEMPLATE_CACHE:
                        del _TEMPLATE_CACHE[cache_key]
                    
                    # 更新修改时间
                    _TEMPLATE_MTIMES[cache_key] = current_mtime
                    refreshed_count += 1
            except Exception as e:
                logger.error(f"刷新模板缓存失败: {name}.{t_type} - {str(e)}")
        
        logger.info(f"缓存刷新完成 | 更新了 {refreshed_count}/{len(all_templates)} 个模板")
    
    def _refresh_cache(self):
        """内部方法：初始加载时刷新缓存"""
        try:
            self.refresh_cache()
        except Exception as e:
            logger.error(f"初始缓存刷新失败: {str(e)}")

# 模块测试
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(level=logging.DEBUG)
    logger.info("模板加载器测试")
    
    try:
        # 创建加载器实例
        loader = TemplateLoader()
        
        # 测试DOCX加载
        doc1 = loader.load_docx_template()
        print(f"DOCX模板加载成功: {len(doc1.paragraphs)} 段落")
        
        # 测试Excel加载
        wb1 = loader.load_xlsx_template()
        print(f"Excel模板加载成功: {wb1.sheetnames} 工作表")
        
        # 测试获取模板列表
        docx_templates = loader.get_available_templates("docx")
        print(f"可用DOCX模板: {docx_templates}")
        
        # 测试获取模板列表（GUI专用）
        docx_template_list = loader.get_template_list("docx")
        print(f"GUI专用模板列表: {docx_template_list}")
                
        # 测试模糊查找
        search_result = loader.find_template("docx", "通用")
        print(f"模糊查找结果: {search_result}")
        
        # 测试缓存机制
        doc2 = loader.load_docx_template()
        print(f"缓存命中测试: {'相同实例' if id(doc1) == id(doc2) else '不同实例'}")
        
        # 模拟修改文件时间
        template_path = loader.get_template_path("docx", "公文通用")
        original_mtime = os.path.getmtime(template_path)
        os.utime(template_path, (time.time(), time.time() + 10))  # 修改时间
        
        # 再次加载
        doc3 = loader.load_docx_template()
        print(f"修改后加载: {'新实例' if id(doc1) != id(doc3) else '相同实例'}")
        
        # 恢复原始时间
        os.utime(template_path, (original_mtime, original_mtime))
        
        # 测试缓存清除
        loader.clear_cache("docx", "公文通用")
        doc4 = loader.load_docx_template()
        print(f"清除后加载: {'新实例' if id(doc3) != id(doc4) else '相同实例'}")
        
        # 测试刷新缓存
        loader.refresh_cache()
        
    except Exception as e:
        logger.error(f"测试失败: {str(e)}")
