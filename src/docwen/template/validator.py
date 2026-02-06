"""
模板验证器模块
验证模板的完整性和占位符合规性
"""

import os
import logging
from .loader import TemplateLoader  # 使用同包内的loader

# 配置日志
logger = logging.getLogger(__name__)

class TemplateValidator:
    """模板验证器核心类"""
    
    def __init__(self, template_loader=None):
        """
        初始化模板验证器
        :param template_loader: 模板加载器实例（可选）
        """
        self.loader = template_loader or TemplateLoader()
        logger.info("初始化模板验证器")
    
    def validate_template(self, template_path, required_placeholders=None):
        """
        验证模板是否包含必需的占位符
        :param template_path: 模板路径
        :param required_placeholders: 必需的占位符列表
        :return: 验证结果 (True/False), 缺失占位符列表, 详细报告
        """
        # 默认必需的占位符
        if not required_placeholders:
            required_placeholders = ["标题", "正文", "发文机关署名"]
        
        logger.info(f"开始验证模板: {template_path}")
        logger.debug(f"必需占位符: {required_placeholders}")
        
        missing = required_placeholders.copy()  # 复制一份避免修改原列表
        report = {
            "template_path": template_path,
            "required_placeholders": required_placeholders,
            "missing": [],
            "found": [],
            "is_valid": False,
            "suggestions": []
        }
        
        # 检查模板文件类型
        if template_path.endswith(".docx"):
            return self._validate_docx_template(template_path, missing, report)
        elif template_path.endswith(".xlsx"):
            return self._validate_excel_template(template_path, missing, report)
        else:
            logger.error(f"不支持的模板格式: {template_path}")
            report["error"] = f"不支持的模板格式: {os.path.basename(template_path)}"
            return False, missing, report
    
    def _validate_docx_template(self, template_path, missing, report):
        """验证DOCX模板"""
        try:
            # 加载模板
            doc = self.loader.load_docx_template(os.path.basename(template_path))
            found_placeholders = set()
            
            # 1. 检查所有段落中的占位符
            for para in doc.paragraphs:
                self._check_text_for_placeholders(para.text, missing, found_placeholders, report)
            
            # 2. 检查表格中的占位符
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            self._check_text_for_placeholders(para.text, missing, found_placeholders, report)
            
            # 更新报告
            report["found"] = list(found_placeholders)
            report["missing"] = [ph for ph in missing if ph not in found_placeholders]
            report["is_valid"] = len(report["missing"]) == 0
            
            # 生成建议
            if report["missing"]:
                report["suggestions"].append(
                    f"添加缺失占位符: {', '.join(report['missing'])}"
                )
            
            logger.info(f"DOCX模板验证结果: {'通过' if report['is_valid'] else '失败'}")
            return report["is_valid"], report["missing"], report
            
        except Exception as e:
            logger.error(f"验证DOCX模板出错: {str(e)}")
            report["error"] = str(e)
            return False, missing, report
    
    def _validate_excel_template(self, template_path, missing, report):
        """验证Excel模板"""
        try:
            # 加载模板
            wb = self.loader.load_xlsx_template(os.path.basename(template_path))
            found_placeholders = set()
            
            # 只检查第一个工作表
            sheet = wb.active
            
            # 检查所有单元格中的占位符
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        self._check_text_for_placeholders(str(cell.value), missing, found_placeholders, report)
            
            # 更新报告
            report["found"] = list(found_placeholders)
            report["missing"] = [ph for ph in missing if ph not in found_placeholders]
            report["is_valid"] = len(report["missing"]) == 0
            
            # 生成建议
            if report["missing"]:
                report["suggestions"].append(
                    f"在单元格中添加缺失占位符: {', '.join(report['missing'])}"
                )
            
            logger.info(f"Excel模板验证结果: {'通过' if report['is_valid'] else '失败'}")
            return report["is_valid"], report["missing"], report
            
        except Exception as e:
            logger.error(f"验证Excel模板出错: {str(e)}")
            report["error"] = str(e)
            return False, missing, report
    
    def _check_text_for_placeholders(self, text, missing, found_placeholders, report):
        """检查文本中是否包含占位符"""
        if not text:
            return
        
        for ph in missing[:]:  # 遍历副本以便安全删除
            placeholder = f"{{{{{ph}}}}}"
            if placeholder in text:
                found_placeholders.add(ph)
                if ph in missing:
                    missing.remove(ph)
                logger.debug(f"找到占位符: {placeholder}")
