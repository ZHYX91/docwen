"""
文件类型检测工具模块
通过文件签名判断文件实际类型
支持检测伪装成DOCX的DOC文件
支持WPS文件格式
提供统一的文件类别定义和扩展名映射
增强实际文件格式检测，避免伪装后缀名
支持文本内容检测，识别伪装的MD/TXT文件
"""

import os
import re
import logging
from typing import Dict, List, Set, Tuple, Optional

# 配置日志
logger = logging.getLogger(__name__)

# 文件类别定义常量
FILE_CATEGORIES: Dict[str, List[str]] = {
    'text': ['.md', '.txt'],
    'document': ['.docx', '.doc', '.wps', '.rtf', '.odt'],
    'spreadsheet': ['.xlsx', '.xls', '.et', '.csv', '.ods'],
    'layout': ['.pdf', '.xps', '.caj', '.ofd'],
    'image': ['.tif', '.tiff', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.heic', '.heif', '.webp']
}

# 类别名称映射（用于显示）
CATEGORY_NAMES: Dict[str, str] = {
    'text': '文本类',
    'document': '文档类',
    'spreadsheet': '表格类',
    'layout': '版式类',
    'image': '图片类'
}

# 所有支持的扩展名集合
SUPPORTED_EXTENSIONS: Set[str] = {ext for category in FILE_CATEGORIES.values() for ext in category}

# 实际格式到类别的映射
ACTUAL_FORMAT_TO_CATEGORY: Dict[str, str] = {
    # 文本类
    'txt': 'text', 'md': 'text',
    
    # 文档类
    'doc': 'document', 'docx': 'document', 'wps': 'document', 'rtf': 'document', 'odt': 'document',
    
    # 表格类
    'xls': 'spreadsheet', 'xlsx': 'spreadsheet', 'et': 'spreadsheet', 'csv': 'spreadsheet', 'ods': 'spreadsheet',
    
    # 版式类
    'pdf': 'layout', 'xps': 'layout', 'caj': 'layout', 'ofd': 'layout',
    
    # 图片类
    'jpeg': 'image', 'jpg': 'image', 'png': 'image', 'gif': 'image', 
    'bmp': 'image', 'tiff': 'image', 'tif': 'image', 'heic': 'image', 'heif': 'image', 'webp': 'image'
}


def _is_text_file(file_path: str, sample_size: int = 8192) -> bool:
    """
    检测文件是否为纯文本文件
    
    策略：
    1. 尝试用UTF-8读取前N个字节
    2. 检查是否包含过多的NULL字节或不可打印字符
    3. 如果能正常解码且字符合理，认为是文本
    
    参数:
        file_path: 文件路径
        sample_size: 采样大小（字节），默认8192
        
    返回:
        bool: 是否为纯文本文件
    """
    try:
        with open(file_path, 'rb') as f:
            sample = f.read(sample_size)
        
        if not sample:  # 空文件
            return True
        
        # 检查NULL字节比例（二进制文件通常有很多NULL）
        null_ratio = sample.count(b'\x00') / len(sample)
        if null_ratio > 0.1:  # 超过10%是NULL，不太可能是文本
            logger.debug(f"文件包含过多NULL字节({null_ratio:.1%})，判定为非文本: {file_path}")
            return False
        
        # 尝试UTF-8解码
        try:
            text = sample.decode('utf-8')
        except UnicodeDecodeError:
            # 尝试其他常见编码
            try:
                text = sample.decode('gbk')
            except UnicodeDecodeError:
                logger.debug(f"文件无法用UTF-8或GBK解码，判定为非文本: {file_path}")
                return False
        
        # 检查可打印字符比例
        printable = sum(c.isprintable() or c.isspace() for c in text)
        printable_ratio = printable / len(text) if text else 0
        
        # 如果大部分字符是可打印的，认为是文本
        is_text = printable_ratio > 0.85
        if is_text:
            logger.debug(f"文件可打印字符比例{printable_ratio:.1%}，判定为文本: {file_path}")
        else:
            logger.debug(f"文件可打印字符比例{printable_ratio:.1%}，判定为非文本: {file_path}")
        
        return is_text
        
    except Exception as e:
        logger.debug(f"检测文本文件失败: {e}")
        return False


def _detect_text_format(file_path: str, sample_lines: int = 50) -> str:
    """
    检测纯文本文件的具体格式（Markdown、CSV 或 TXT）
    
    检测策略：
    1. Markdown 检测：
       - 标题：# ## ### 等 Markdown 标题
       - 表格：表格分隔线 |---|---|
       - 代码块：``` 或 ~~~ 标记
       - YAML Front Matter：文件开头的 --- ... ---
    2. CSV 检测：是否具有统一的列分隔符和一致的列数
    3. 默认为 TXT：都不符合时判定为普通文本
    
    参数:
        file_path: 文件路径
        sample_lines: 采样行数，默认50行
        
    返回:
        str: 'md'（Markdown）、'csv'（CSV表格）或 'txt'（普通文本）
    """
    try:
        # 尝试多种编码读取文件
        content = None
        for encoding in ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    lines = [f.readline() for _ in range(sample_lines)]
                    lines = [line for line in lines if line]  # 过滤空行
                content = lines
                logger.debug(f"成功使用 {encoding} 编码读取文件: {file_path}")
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
            except Exception as e:
                logger.debug(f"使用 {encoding} 编码读取失败: {e}")
                continue
        
        if not content:
            logger.debug(f"无法读取文件内容，默认判定为txt: {file_path}")
            return 'txt'
        
        # 1. Markdown 检测：检查多种 Markdown 特征
        markdown_heading_pattern = re.compile(r'^#{1,6}\s+.+')
        # Markdown 表格分隔行模式：|------|------|  或  | --- | --- |
        markdown_table_separator = re.compile(r'^\s*\|[\s\-:]+\|[\s\-:]*')
        # Wiki链接模式：[[链接]] 或 [[链接|别名]]
        wiki_link_pattern = re.compile(r'\[\[.+?\]\]')
        # Markdown链接模式：[文本](链接) 或 ![图片](链接)
        markdown_link_pattern = re.compile(r'!?\[.+?\]\(.+?\)')
        # Markdown参考式链接：[文本][ref]
        markdown_ref_pattern = re.compile(r'\[.+?\]\[.+?\]')
        
        markdown_heading_count = 0
        markdown_table_count = 0
        code_block_markers = 0
        wiki_link_count = 0
        markdown_link_count = 0
        yaml_front_matter = False
        
        # 检测 YAML Front Matter（文件开头的 --- ... ---）
        if len(content) >= 2 and content[0].strip() == '---':
            for i in range(1, min(len(content), 20)):
                if content[i].strip() == '---':
                    yaml_front_matter = True
                    logger.info(f"检测到YAML Front Matter，判定为md格式: {file_path}")
                    break
        
        # 检测 Markdown 标题、表格、代码块、链接
        for line in content:
            stripped = line.strip()
            
            # 标题检测
            if markdown_heading_pattern.match(stripped):
                markdown_heading_count += 1
            
            # 表格分隔线检测
            elif markdown_table_separator.match(stripped):
                # 检查是否包含足够的 - 字符（至少3个）
                if stripped.count('-') >= 3:
                    markdown_table_count += 1
            
            # 代码块标记检测（``` 或 ~~~）
            if stripped.startswith('```') or stripped.startswith('~~~'):
                code_block_markers += 1
            
            # Wiki链接检测
            wiki_link_count += len(wiki_link_pattern.findall(line))
            
            # Markdown链接检测（标准链接和参考式链接）
            markdown_link_count += len(markdown_link_pattern.findall(line))
            markdown_link_count += len(markdown_ref_pattern.findall(line))
        
        # 判定为 Markdown 的条件（任一满足即可）：
        # - 有 YAML Front Matter
        # - 有 Markdown 标题
        # - 有 Markdown 表格
        # - 有至少2个代码块标记（开始+结束）
        # - 有 Wiki 链接
        # - 有 Markdown 链接
        if (yaml_front_matter or 
            markdown_heading_count > 0 or 
            markdown_table_count > 0 or 
            code_block_markers >= 2 or
            wiki_link_count > 0 or
            markdown_link_count > 0):
            
            # 记录详细的检测结果
            features = []
            if yaml_front_matter:
                features.append("YAML Front Matter")
            if markdown_heading_count > 0:
                features.append(f"{markdown_heading_count}个标题")
            if markdown_table_count > 0:
                features.append(f"{markdown_table_count}个表格")
            if code_block_markers >= 2:
                features.append(f"{code_block_markers}个代码块标记")
            if wiki_link_count > 0:
                features.append(f"{wiki_link_count}个Wiki链接")
            if markdown_link_count > 0:
                features.append(f"{markdown_link_count}个Markdown链接")
            
            logger.info(f"检测到Markdown特征（{', '.join(features)}），判定为md格式: {file_path}")
            return 'md'
        
        # 2. CSV 检测：检查是否有统一的列分隔符
        if len(content) >= 2:  # 至少需要2行才能判断表格
            # 尝试常见的分隔符
            delimiters = [',', '\t', ';', '|']
            best_delimiter = None
            best_score = 0
            
            for delimiter in delimiters:
                # 统计每行的列数
                column_counts = []
                for line in content:
                    line = line.strip()
                    if not line:  # 跳过空行
                        continue
                    # 分割并计数
                    columns = line.split(delimiter)
                    if len(columns) > 1:  # 至少要有2列
                        column_counts.append(len(columns))
                
                # 检查列数一致性
                if len(column_counts) >= 2:
                    # 计算列数的标准差（列数越一致，标准差越小）
                    avg_cols = sum(column_counts) / len(column_counts)
                    if avg_cols >= 2:  # 平均至少2列
                        # 检查是否所有行的列数相同或相近
                        unique_counts = set(column_counts)
                        if len(unique_counts) <= 2:  # 列数最多只有2种不同值（允许表头行不同）
                            score = len(column_counts) * avg_cols / len(unique_counts)
                            if score > best_score:
                                best_score = score
                                best_delimiter = delimiter
            
            # 如果找到合适的分隔符，判定为 CSV
            if best_delimiter is not None and best_score >= 4:  # 阈值：至少4分
                delimiter_name = {',': '逗号', '\t': '制表符', ';': '分号', '|': '竖线'}.get(best_delimiter, best_delimiter)
                logger.info(f"检测到表格格式（分隔符：{delimiter_name}），判定为csv格式: {file_path}")
                return 'csv'
        
        # 3. 默认判定为普通文本
        logger.debug(f"未检测到特殊格式标记，判定为txt格式: {file_path}")
        return 'txt'
        
    except Exception as e:
        logger.debug(f"检测文本格式失败: {e}，默认判定为txt")
        return 'txt'


def _detect_ole_type(file_path: str) -> str:
    """
    检测OLE复合文档的类型（DOC/XLS/WPS/ET）
    
    通过解析OLE内部流结构来区分不同的Office格式：
    - Word文档（DOC/WPS）包含 'WordDocument' 流
    - Excel工作簿（XLS/ET）包含 'Workbook' 或 'Book' 流
    
    参数:
        file_path: 文件路径
        
    返回:
        str: 'word'（文档类） 或 'excel'（表格类） 或 'unknown'（无法确定）
    """
    try:
        import olefile
        
        if not olefile.isOleFile(file_path):
            logger.debug(f"不是有效的OLE文件: {file_path}")
            return 'unknown'
        
        with olefile.OleFileIO(file_path) as ole:
            # 获取所有流的列表
            streams = ole.listdir()
            stream_names = ['/'.join(stream) for stream in streams]
            
            logger.debug(f"OLE文件内部流: {stream_names[:10]}")  # 只记录前10个
            
            # 检测Word文档特征流
            word_indicators = ['WordDocument', '1Table', '0Table', 'Data']
            if any(indicator in stream_names for indicator in word_indicators):
                logger.debug(f"检测到Word文档特征流，判定为word类型")
                return 'word'
            
            # 检测Excel工作簿特征流
            excel_indicators = ['Workbook', 'Book', 'BOOK']
            if any(indicator in stream_names for indicator in excel_indicators):
                logger.debug(f"检测到Excel工作簿特征流，判定为excel类型")
                return 'excel'
            
            # 无法确定
            logger.debug(f"无法通过OLE内部流确定类型")
            return 'unknown'
            
    except ImportError:
        logger.warning(f"olefile库未安装，无法解析OLE内部结构")
        logger.info("如需OLE格式精确识别，请安装: pip install olefile")
        return 'unknown'
    except Exception as e:
        logger.debug(f"解析OLE文件失败: {e}")
        return 'unknown'


def detect_actual_file_format(file_path: str) -> str:
    """
    通过文件签名检测文件实际格式
    支持文档、表格、图片、版式等各类文件格式
    
    参数:
        file_path: 文件路径
        
    返回:
        str: 文件实际格式标识符
    """
    try:
        with open(file_path, 'rb') as f:
            header = f.read(16)  # 增加读取长度以支持更多格式
        
        # 图片格式检测
        # JPEG: FF D8 FF
        if header.startswith(b'\xFF\xD8\xFF'):
            logger.debug(f"文件签名检测为JPEG: {file_path}")
            return 'jpeg'
        
        # PNG: 89 50 4E 47 0D 0A 1A 0A
        elif header.startswith(b'\x89PNG\r\n\x1a\n'):
            logger.debug(f"文件签名检测为PNG: {file_path}")
            return 'png'
        
        # GIF: 47 49 46 38
        elif header.startswith(b'GIF8'):
            logger.debug(f"文件签名检测为GIF: {file_path}")
            return 'gif'
        
        # BMP: 42 4D
        elif header.startswith(b'BM'):
            logger.debug(f"文件签名检测为BMP: {file_path}")
            return 'bmp'
        
        # TIFF: 49 49 2A 00 或 4D 4D 00 2A
        elif header.startswith(b'II\x2a\x00') or header.startswith(b'MM\x00\x2a'):
            logger.debug(f"文件签名检测为TIFF: {file_path}")
            return 'tiff'
        
        # HEIC/HEIF: ftypheic
        elif b'ftypheic' in header[:12]:
            logger.debug(f"文件签名检测为HEIC: {file_path}")
            return 'heic'
        
        # WebP: RIFF + WEBP (前4字节是RIFF，偏移8-11字节是WEBP)
        elif header[:4] == b'RIFF' and len(header) >= 12 and header[8:12] == b'WEBP':
            logger.debug(f"文件签名检测为WebP: {file_path}")
            return 'webp'
        
        # 版式文件检测
        # PDF: 25 50 44 46
        elif header.startswith(b'%PDF'):
            logger.debug(f"文件签名检测为PDF: {file_path}")
            return 'pdf'
        
        # 文档格式检测
        # RTF文件签名: {\rtf
        elif header.startswith(b'{\\rtf'):
            logger.debug(f"文件签名检测为RTF: {file_path}")
            return 'rtf'
        
        # DOC/XLS/WPS/ET文件签名: D0 CF 11 E0 A1 B1 1A E1 (OLE格式)
        elif header.startswith(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'):
            # OLE文档和表格共享签名，需要解析内部流来区分
            logger.debug(f"检测到OLE格式文件，解析内部结构: {file_path}")
            
            ole_type = _detect_ole_type(file_path)
            ext = os.path.splitext(file_path)[1].lower()
            
            if ole_type == 'word':
                # 确认为Word文档，根据扩展名区分DOC和WPS
                if ext == '.wps':
                    logger.debug(f"OLE内部流检测为WPS: {file_path}")
                    return 'wps'
                else:
                    logger.debug(f"OLE内部流检测为DOC: {file_path}")
                    return 'doc'
            
            elif ole_type == 'excel':
                # 确认为Excel工作簿，根据扩展名区分XLS和ET
                if ext == '.et':
                    logger.debug(f"OLE内部流检测为ET: {file_path}")
                    return 'et'
                else:
                    logger.debug(f"OLE内部流检测为XLS: {file_path}")
                    return 'xls'
            
            else:
                # 无法通过内部流确定，回退到扩展名判断
                logger.warning(f"OLE内部流无法确定类型，回退到扩展名判断: {file_path}")
                
                if ext in ['.xls', '.xlsx', '.et']:
                    logger.debug(f"扩展名判断为XLS/ET: {file_path}")
                    return 'et' if ext == '.et' else 'xls'
                elif ext == '.wps':
                    logger.debug(f"扩展名判断为WPS: {file_path}")
                    return 'wps'
                else:
                    logger.debug(f"扩展名判断为DOC: {file_path}")
                    return 'doc'
        
        # DOCX/XLSX/ODT/ODS/OFD文件签名: PK 03 04 (ZIP格式)
        elif header.startswith(b'PK\x03\x04'):
            # 检查ZIP包内部结构以区分具体格式
            try:
                import zipfile
                with zipfile.ZipFile(file_path, 'r') as zip_file:
                    file_list = zip_file.namelist()
                    
                    # 1. 检查OFD格式（包含OFD.xml）
                    if 'OFD.xml' in file_list or any(name.startswith('Doc_') for name in file_list):
                        logger.debug(f"文件签名检测为OFD: {file_path}")
                        return 'ofd'
                    
                    # 2. 检查ODT/ODS文件（通过mimetype识别）
                    if 'mimetype' in file_list:
                        with zip_file.open('mimetype') as mimetype_file:
                            mimetype = mimetype_file.read().decode('utf-8').strip()
                            if mimetype == 'application/vnd.oasis.opendocument.text':
                                logger.debug(f"文件签名检测为ODT: {file_path}")
                                return 'odt'
                            elif mimetype == 'application/vnd.oasis.opendocument.spreadsheet':
                                logger.debug(f"文件签名检测为ODS: {file_path}")
                                return 'ods'
                    
                    # 3. 检查XPS文件（包含FixedDocumentSequence.fdseq或Documents/目录）
                    if 'FixedDocumentSequence.fdseq' in file_list or \
                       any(name.startswith('Documents/') and name.endswith('.fpage') for name in file_list):
                        logger.debug(f"文件签名检测为XPS: {file_path}")
                        return 'xps'
                    
                    # 4. 检查XLSX文件（包含xl/目录）
                    if any(name.startswith('xl/') for name in file_list):
                        logger.debug(f"文件签名检测为XLSX: {file_path}")
                        return 'xlsx'
                    
                    # 5. 检查DOCX文件（包含word/目录）
                    if any(name.startswith('word/') for name in file_list):
                        logger.debug(f"文件签名检测为DOCX: {file_path}")
                        return 'docx'
                    
                    # 6. 无法确定，回退到扩展名
                    logger.warning(f"ZIP文件无明显特征，回退到扩展名判断: {file_path}")
                    ext = os.path.splitext(file_path)[1].lower()
                    if ext == '.xlsx':
                        return 'xlsx'
                    elif ext == '.xps':
                        return 'xps'
                    elif ext == '.ods':
                        return 'ods'
                    elif ext == '.odt':
                        return 'odt'
                    elif ext == '.ofd':
                        return 'ofd'
                    else:
                        return 'docx'
                        
            except Exception as e:
                logger.warning(f"ZIP文件详细检测失败: {str(e)}，回退到扩展名")
                ext = os.path.splitext(file_path)[1].lower()
                if ext == '.xlsx':
                    return 'xlsx'
                elif ext == '.xps':
                    return 'xps'
                elif ext == '.ods':
                    return 'ods'
                elif ext == '.odt':
                    return 'odt'
                elif ext == '.ofd':
                    return 'ofd'
                else:
                    return 'docx'
        
        else:
            # 没有明确的二进制文件签名，尝试文本内容检测
            logger.debug(f"未检测到明确的二进制签名，尝试文本内容检测: {file_path}")
            
            if _is_text_file(file_path):
                # 是文本文件，通过内容检测具体格式
                text_format = _detect_text_format(file_path)
                logger.info(f"纯文本文件，内容检测为{text_format.upper()}格式: {file_path}")
                return text_format
            
            # 不是文本文件，回退到扩展名判断
            logger.debug(f"非文本文件，回退到扩展名判断: {file_path}")
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.rtf':
                logger.debug(f"扩展名检测为RTF: {file_path}")
                return 'rtf'
            elif ext == '.doc':
                logger.debug(f"扩展名检测为DOC: {file_path}")
                return 'doc'
            elif ext == '.docx':
                logger.debug(f"扩展名检测为DOCX: {file_path}")
                return 'docx'
            elif ext == '.wps':
                logger.debug(f"扩展名检测为WPS: {file_path}")
                return 'wps'
            elif ext == '.xls':
                logger.debug(f"扩展名检测为XLS: {file_path}")
                return 'xls'
            elif ext == '.xlsx':
                logger.debug(f"扩展名检测为XLSX: {file_path}")
                return 'xlsx'
            elif ext == '.et':
                logger.debug(f"扩展名检测为ET: {file_path}")
                return 'et'
            elif ext == '.odt':
                logger.debug(f"扩展名检测为ODT: {file_path}")
                return 'odt'
            elif ext == '.pdf':
                logger.debug(f"扩展名检测为PDF: {file_path}")
                return 'pdf'
            elif ext == '.xps':
                logger.debug(f"扩展名检测为XPS: {file_path}")
                return 'xps'
            elif ext in ['.jpg', '.jpeg']:
                logger.debug(f"扩展名检测为JPEG: {file_path}")
                return 'jpeg'
            elif ext == '.png':
                logger.debug(f"扩展名检测为PNG: {file_path}")
                return 'png'
            elif ext == '.gif':
                logger.debug(f"扩展名检测为GIF: {file_path}")
                return 'gif'
            elif ext == '.bmp':
                logger.debug(f"扩展名检测为BMP: {file_path}")
                return 'bmp'
            elif ext in ['.tif', '.tiff']:
                logger.debug(f"扩展名检测为TIFF: {file_path}")
                return 'tiff'
            elif ext in ['.heic', '.heif']:
                logger.debug(f"扩展名检测为HEIC: {file_path}")
                return 'heic'
            elif ext == '.txt':
                logger.debug(f"扩展名检测为TXT: {file_path}")
                return 'txt'
            elif ext == '.md':
                logger.debug(f"扩展名检测为MD: {file_path}")
                return 'md'
            else:
                logger.debug(f"未知文件类型: {file_path}")
                return 'unknown'
                
    except Exception as e:
        logger.error(f"检查文件格式失败: {str(e)}")
        # 回退到扩展名
        ext = os.path.splitext(file_path)[1].lower()
        format_mapping = {
            '.rtf': 'rtf', '.doc': 'doc', '.docx': 'docx', '.wps': 'wps',
            '.xls': 'xls', '.xlsx': 'xlsx', '.et': 'et', '.odt': 'odt',
            '.pdf': 'pdf', '.xps': 'xps', '.jpg': 'jpeg', '.jpeg': 'jpeg', '.png': 'png',
            '.gif': 'gif', '.bmp': 'bmp', '.tif': 'tiff', '.tiff': 'tiff',
            '.heic': 'heic', '.heif': 'heic', '.txt': 'txt', '.md': 'md'
        }
        return format_mapping.get(ext, 'unknown')



def get_file_category(file_path: str) -> str:
    """
    根据文件扩展名获取文件类别
    
    参数:
        file_path: 文件路径
        
    返回:
        str: 文件类别
            - "text": 文本类（.md, .txt）
            - "document": 文档类（.docx, .doc, .wps, .rtf）
            - "spreadsheet": 表格类（.xlsx, .xls, .et, .csv）
            - "layout": 版式类（.pdf）
            - "image": 图片类（.tif, .tiff, .jpg, .jpeg, .png, .bmp, .gif, .heic, .heif）
            - "unknown": 未知类型
    """
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    for category, extensions in FILE_CATEGORIES.items():
        if ext in extensions:
            logger.debug(f"文件类别检测: {file_path} -> {category}")
            return category
    
    logger.debug(f"未知文件类别: {file_path}")
    return 'unknown'


def get_category_name(category: str) -> str:
    """
    获取类别显示名称
    
    参数:
        category: 文件类别
        
    返回:
        str: 类别显示名称，如果类别不存在则返回原类别
    """
    return CATEGORY_NAMES.get(category, category)


def is_supported_file(file_path: str) -> bool:
    """
    检查文件是否支持
    
    参数:
        file_path: 文件路径
        
    返回:
        bool: 文件是否支持
    """
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    # 允许无后缀名文件通过，后续会在格式验证中标记为错误
    if not ext:
        logger.debug(f"文件无扩展名，允许通过: {file_path}")
        return True
    
    is_supported = ext in SUPPORTED_EXTENSIONS
    logger.debug(f"文件支持检查: {file_path} -> {is_supported}")
    return is_supported


def get_allowed_types_by_category(category: str) -> str:
    """
    根据类别获取允许的文件类型字符串（用顿号分隔）
    
    参数:
        category: 文件类别
        
    返回:
        str: 允许的文件类型字符串，用顿号分隔
    """
    if category not in FILE_CATEGORIES:
        return '未知类型'
    
    extensions = FILE_CATEGORIES[category]
    # 将扩展名转换为大写格式（去掉点号）
    type_names = [ext[1:].upper() for ext in extensions]
    # 用顿号连接
    return '、'.join(type_names)


def _has_known_signature(file_path: str) -> bool:
    """
    检查文件是否有已知的二进制签名（不包含回退到扩展名的情况）
    
    参数:
        file_path: 文件路径
        
    返回:
        bool: 是否有已知签名
    """
    try:
        with open(file_path, 'rb') as f:
            header = f.read(16)
        
        # 检查所有已知的二进制签名
        known_signatures = [
            # 图片格式
            (b'\xFF\xD8\xFF', 'jpeg'),                    # JPEG
            (b'\x89PNG\r\n\x1a\n', 'png'),                # PNG
            (b'GIF8', 'gif'),                             # GIF
            (b'BM', 'bmp'),                               # BMP
            (b'II\x2a\x00', 'tiff'),                      # TIFF (little-endian)
            (b'MM\x00\x2a', 'tiff'),                      # TIFF (big-endian)
            # 文档格式
            (b'%PDF', 'pdf'),                             # PDF
            (b'{\\rtf', 'rtf'),                           # RTF
            (b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1', 'ole'), # OLE (DOC/XLS/WPS/ET)
            (b'PK\x03\x04', 'zip'),                       # ZIP (DOCX/XLSX/ODT/ODS/XPS/OFD)
        ]
        
        for signature, _ in known_signatures:
            if header.startswith(signature):
                return True
        
        # 检查HEIC/WebP等特殊格式
        if b'ftypheic' in header[:12]:
            return True
        if header[:4] == b'RIFF' and len(header) >= 12 and header[8:12] == b'WEBP':
            return True
        
        return False
        
    except Exception as e:
        logger.debug(f"检查文件签名失败: {e}")
        return False


def validate_file_format(file_path: str) -> dict:
    """
    验证文件格式是否匹配（扩展名与实际格式）
    
    参数:
        file_path: 文件路径
        
    返回:
        dict: 格式验证结果
            - actual_format: 实际格式 (如 'doc')
            - extension_format: 扩展名格式 (如 'docx')
            - is_match: 是否匹配
            - warning_message: 警告消息（不匹配时或无法验证时）
    """
    # 检测实际格式
    actual_format = detect_actual_file_format(file_path)
    
    # 获取扩展名格式
    ext = os.path.splitext(file_path)[1].lower()
    extension_format = ext[1:] if ext else 'unknown'
    
    # 格式映射到显示名称
    format_names = {
        'doc': 'DOC', 'docx': 'DOCX', 'wps': 'WPS', 'rtf': 'RTF', 'odt': 'ODT',
        'xls': 'XLS', 'xlsx': 'XLSX', 'et': 'ET', 'csv': 'CSV',
        'pdf': 'PDF',
        'jpeg': 'JPEG', 'png': 'PNG', 'gif': 'GIF', 'bmp': 'BMP', 'tiff': 'TIFF', 'heic': 'HEIC',
        'txt': 'TXT', 'md': 'MD'
    }
    
    actual_name = format_names.get(actual_format, actual_format.upper())
    ext_name = format_names.get(extension_format, extension_format.upper())
    
    # 定义等价的扩展名组（用于处理 jpg/jpeg、tif/tiff、heic/heif 等情况）
    equivalent_formats = [
        {'jpeg', 'jpg'},      # JPEG 的两种扩展名
        {'tiff', 'tif'},      # TIFF 的两种扩展名
        {'heic', 'heif'}      # HEIC 的两种扩展名
    ]
    
    # 检查是否匹配（考虑等价格式）
    def are_formats_equivalent(fmt1: str, fmt2: str) -> bool:
        """检查两个格式是否等价"""
        if fmt1 == fmt2:
            return True
        for equiv_group in equivalent_formats:
            if fmt1 in equiv_group and fmt2 in equiv_group:
                return True
        return False
    
    is_match = are_formats_equivalent(actual_format, extension_format) or (actual_format == 'unknown')
    
    # 检查是否有已知签名（用于判断是否是回退到扩展名的情况）
    has_signature = _has_known_signature(file_path)
    
    # 检查是否是纯文本文件（文本文件通过内容检测，不是回退）
    is_text_format = actual_format in ['txt', 'md', 'csv']
    
    # 生成警告消息
    warning_message = ""
    if not is_match:
        # 格式不匹配的情况
        if actual_format in ['txt', 'md', 'csv']:
            # 区分各种纯文本格式
            format_display = {
                'txt': 'TXT(纯文本)',
                'md': 'Markdown',
                'csv': 'CSV(表格文本)'
            }
            warning_message = f"⚠️ 实际为 {format_display[actual_format]} 格式"
            logger.warning(f"检测到格式不匹配: {file_path} (扩展名: {ext_name}, 实际: {format_display[actual_format]})")
        else:
            warning_message = f"⚠️ 实际为 {actual_name} 格式"
            logger.warning(f"检测到格式不匹配: {file_path} (扩展名: {ext_name}, 实际: {actual_name})")
    elif is_match and not has_signature and not is_text_format and actual_format != 'unknown':
        # 格式匹配但没有已知签名（回退到扩展名的情况）
        # 不是纯文本格式，也不是未知格式
        warning_message = "⚠️ 无法验证文件实际格式"
        logger.warning(f"无法验证文件格式: {file_path} (扩展名: {ext_name}, 无已知签名)")
    else:
        logger.debug(f"文件格式匹配: {file_path} -> {actual_format}")
    
    return {
        'actual_format': actual_format,
        'extension_format': extension_format,
        'is_match': is_match,
        'warning_message': warning_message
    }


def get_actual_file_category(file_path: str) -> str:
    """
    基于实际文件格式获取文件类别
    避免伪装后缀名问题
    
    参数:
        file_path: 文件路径
        
    返回:
        str: 文件实际类别
    """
    # 检测实际文件格式
    actual_format = detect_actual_file_format(file_path)
    
    # 根据实际格式映射到类别
    if actual_format in ACTUAL_FORMAT_TO_CATEGORY:
        category = ACTUAL_FORMAT_TO_CATEGORY[actual_format]
        logger.debug(f"实际文件类别检测: {file_path} -> {category} (格式: {actual_format})")
        return category
    
    # 如果无法通过实际格式确定，回退到扩展名检测
    logger.debug(f"无法通过实际格式确定类别，回退到扩展名检测: {file_path}")
    return get_file_category(file_path)


def get_file_info(file_path: str) -> dict:
    """
    获取完整的文件信息
    包含实际格式、类别、验证结果等
    
    参数:
        file_path: 文件路径
        
    返回:
        dict: 完整文件信息
            - file_path: 文件路径
            - actual_format: 实际格式
            - actual_category: 实际类别
            - extension: 文件扩展名
            - extension_category: 扩展名对应的类别
            - is_supported: 是否支持
            - is_valid: 格式是否有效
            - warning_message: 警告消息
    """
    # 验证文件格式
    validation_result = validate_file_format(file_path)
    
    # 获取实际类别
    actual_category = get_actual_file_category(file_path)
    
    # 获取扩展名类别
    extension_category = get_file_category(file_path)
    
    # 检查是否支持
    is_supported = is_supported_file(file_path)
    
    # 判断文件是否有效（格式匹配且支持）
    is_valid = validation_result['is_match'] and is_supported
    
    return {
        'file_path': file_path,
        'actual_format': validation_result['actual_format'],
        'actual_category': actual_category,
        'extension': os.path.splitext(file_path)[1].lower(),
        'extension_category': extension_category,
        'is_supported': is_supported,
        'is_valid': is_valid,
        'warning_message': validation_result['warning_message']
    }


def activate_optimal_tab(file_paths: List[str]) -> str:
    """
    根据文件类别统计激活最优选项卡
    统一的选项卡切换逻辑，适用于单文件和批量模式
    
    参数:
        file_paths: 文件路径列表
        
    返回:
        str: 最优选项卡类别
    """
    if not file_paths:
        return 'text'  # 默认返回文本选项卡
    
    # 统计各类别文件数量
    category_count = {
        'text': 0, 'document': 0, 'spreadsheet': 0, 'layout': 0, 'image': 0
    }
    
    for file_path in file_paths:
        # 使用实际文件类别进行统计
        category = get_actual_file_category(file_path)
        if category in category_count:
            category_count[category] += 1
    
    logger.debug(f"文件类别统计: {category_count}")
    
    # 找到文件数量最多的类别
    max_count = max(category_count.values())
    candidates = [cat for cat, count in category_count.items() if count == max_count]
    
    # 如果只有一个候选，直接返回
    if len(candidates) == 1:
        logger.info(f"自动激活选项卡: {candidates[0]} (文件数: {max_count})")
        return candidates[0]
    
    # 多个候选时按优先级排序：文本 > 文档 > 表格 > 版式文件 > 图片
    priority_order = ['text', 'document', 'spreadsheet', 'layout', 'image']
    for category in priority_order:
        if category in candidates:
            logger.info(f"自动激活选项卡: {category} (文件数: {max_count}, 按优先级选择)")
            return category
    
    # 默认返回第一个候选
    logger.info(f"自动激活选项卡: {candidates[0]} (文件数: {max_count}, 默认选择)")
    return candidates[0]
