"""
DocWen 源代码汇总脚本
用于将src目录下所有Python文件汇总到doc/源代码文件.md中，便于软著申请。

功能：
1. 递归遍历src目录下的所有.py文件
2. 读取文件内容并统计信息
3. 生成格式化的Markdown文档
4. 输出适合软著申请的源代码文档
"""

import os
import sys
import datetime
from pathlib import Path


class SourceCodeCollector:
    """源代码收集器"""

    def __init__(self, src_dir="src", output_file="doc/源代码文件.md"):
        self.src_dir = Path(src_dir)
        self.output_file = Path(output_file)
        self.files_data = []
        self.total_lines = 0
        self.total_files = 0

    def collect_files(self):
        """收集所有Python文件"""
        print("开始收集源代码文件...")

        for root, dirs, files in os.walk(self.src_dir):
            # 按目录名排序，保持一致的输出顺序
            dirs.sort()
            files.sort()

            for file in files:
                if file.endswith(".py"):
                    file_path = Path(root) / file
                    relative_path = file_path.relative_to(self.src_dir.parent)

                    try:
                        # 读取文件内容
                        with open(file_path, "r", encoding="utf-8") as f:
                            content = f.read()

                        # 统计行数
                        lines = content.count("\n") + 1

                        self.files_data.append({"path": relative_path, "content": content, "lines": lines})

                        self.total_lines += lines
                        self.total_files += 1

                        print(f"  ✓ 已收集: {relative_path} ({lines} 行)")

                    except Exception as e:
                        print(f"  ✗ 读取失败: {relative_path} - {e}")

        print(f"收集完成: {self.total_files} 个文件, {self.total_lines} 行代码")

    def generate_markdown(self):
        """生成Markdown格式的源代码文档"""
        print("生成Markdown文档...")

        # 创建文档头部
        markdown_content = self._generate_header()

        # 添加文件索引
        markdown_content += self._generate_file_index()

        # 添加详细源代码
        markdown_content += self._generate_detailed_code()

        # 写入文件
        self.output_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self.output_file, "w", encoding="utf-8") as f:
            f.write(markdown_content)

        print(f"文档生成完成: {self.output_file}")

    def _generate_header(self):
        """生成文档头部"""
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        header = f"""# DocWen 源代码文档

## 项目概述

- **软件名称**: DocWen
- **简称**: DocWen
- **全称**: 公文与表格格式转换工具
- **版本号**: 从 src/docwen/__init__.py 读取
- **开发语言**: Python
- **总文件数**: {self.total_files} 个
- **总代码行数**: {self.total_lines} 行
- **生成时间**: {current_time}
- **用途**: 软件著作权申请材料

## 软件简介

本地运行的公文格式转换工具，支持 Word ↔ Markdown ↔ Excel 多向转换，内置文本校对系统。软件完全离线运行，数据不出本地，采用智能识别算法（多轮评分）解析公文元素，基于模板生成规范格式文档，并提供自定义规则的文本校对功能。界面友好，操作简单直观。

## 主要功能

1. **Word文档转Markdown** - 智能解析公文格式结构
2. **Markdown转Word文档** - 基于模板生成规范公文
3. **表格文件处理** - Excel/CSV与Markdown互转
4. **图片、版式文件转Markdown** - 提取图片、文字识别（OCR）
5. **多格式互相转换** - 文档、表格、图片、版式文件格式互转
6. **表格汇总** - 将多个表格按行、按列、按单元格汇总
7. **版式文件合并拆分** - 合并多个版式文件为PDF，或拆分PDF为多个文件
8. **图片合并** - 将多个图片合并为多页TIFF文件
9. **文本校对系统** - 4种校对规则（标点、符号、错别字、敏感词）
10. **模板管理系统** - 内置多种常用公文和表格模板

---
"""
        return header

    def _generate_file_index(self):
        """生成文件索引"""
        index_content = "## 源代码文件列表\n\n"

        # 按目录分组
        files_by_dir = {}
        for file_data in self.files_data:
            dir_path = str(file_data["path"].parent)
            if dir_path not in files_by_dir:
                files_by_dir[dir_path] = []
            files_by_dir[dir_path].append(file_data)

        # 按目录结构排序
        sorted_dirs = sorted(files_by_dir.keys())

        for dir_path in sorted_dirs:
            # 处理根目录的特殊情况
            if dir_path == ".":
                display_dir = "根目录"
            else:
                display_dir = dir_path

            index_content += f"### {display_dir}\n\n"

            for file_data in sorted(files_by_dir[dir_path], key=lambda x: x["path"].name):
                filename = file_data["path"].name
                lines = file_data["lines"]
                index_content += f"- `{filename}` - {lines} 行代码\n"

            index_content += "\n"

        index_content += "---\n\n"
        return index_content

    def _generate_detailed_code(self):
        """生成详细源代码部分（优化版：模块文件统一使用5级标题）"""
        detailed_content = "## 详细源代码\n\n"

        # 按目录结构组织文件
        file_tree = self._build_file_tree()

        # 生成目录内容
        detailed_content += self._generate_tree_content(file_tree)

        return detailed_content

    def _build_file_tree(self):
        """构建文件树结构"""
        file_tree = {}

        for file_data in self.files_data:
            path_parts = list(file_data["path"].parts)

            # 跳过src目录，直接从下一级开始
            if path_parts[0] == "src":
                path_parts = path_parts[1:]

            current_level = file_tree

            # 构建目录树
            for i, part in enumerate(path_parts):
                if i == len(path_parts) - 1:  # 文件
                    current_level[part] = file_data
                else:  # 目录
                    if part not in current_level:
                        current_level[part] = {}
                    current_level = current_level[part]

        return file_tree

    def _generate_tree_content(self, tree, level=0):
        """递归生成树形目录内容"""
        content = ""

        # 先处理当前层的文件（模块文件统一使用5级标题）
        files = {k: v for k, v in tree.items() if isinstance(v, dict) and "path" not in v}
        modules = {k: v for k, v in tree.items() if not isinstance(v, dict) or "path" in v}

        # 处理模块文件（5级标题）
        for name, file_data in sorted(modules.items()):
            if isinstance(file_data, dict) and "path" in file_data:  # 确保是文件数据
                content += f"##### {file_data['path']}\n\n"
                content += f"```python\n{file_data['content']}\n```\n\n"

        # 处理目录（3级或4级标题）
        for name, subtree in sorted(files.items()):
            # 确定标题级别
            if level == 0:  # 一级子包
                content += f"### {name}\n\n"
            else:  # 二级及更深子包
                content += f"#### {name}\n\n"

            # 递归处理子目录
            content += self._generate_tree_content(subtree, level + 1)

        return content

    def run(self):
        """运行完整的收集和生成流程"""
        print("=" * 60)
        print("DocWen 源代码汇总工具")
        print("=" * 60)

        # 检查源目录是否存在
        if not self.src_dir.exists():
            print(f"错误: 源目录不存在: {self.src_dir}")
            return False

        # 收集文件
        self.collect_files()

        if self.total_files == 0:
            print("错误: 未找到任何Python文件")
            return False

        # 生成文档
        self.generate_markdown()

        print("=" * 60)
        print(f"汇总完成!")
        print(f"- 输出文件: {self.output_file}")
        print(f"- 总文件数: {self.total_files}")
        print(f"- 总代码行数: {self.total_lines}")
        print("=" * 60)

        return True


def main():
    """主函数"""
    collector = SourceCodeCollector()
    success = collector.run()

    if success:
        print("源代码文档生成成功，可用于软著申请。")
        return 0
    else:
        print("源代码文档生成失败。")
        return 1


if __name__ == "__main__":
    sys.exit(main())
