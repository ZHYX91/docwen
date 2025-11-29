"""
命令行主入口模块
实现完整的命令行交互功能
支持交互式文件处理
使用极简日志系统初始化
"""

import sys
import os
import logging
import traceback
from gongwen_converter.cli.router import route_file

# === 文件转换路由 ===
def convert_file(file_path, show_success=True, interactive=True):
    """文件转换路由"""
    try:
        # 使用新的路由函数（启用交互模式）
        return route_file(file_path, show_success, interactive)
    except Exception as e:
        logging.error(f"文件转换失败: {file_path}", exc_info=True)
        return False, None

# === 路径解析器 ===
def parse_file_path(user_input: str) -> str:
    """
    文件路径解析器
    支持多种拖放格式：
    1. PowerShell格式: & 'd:\测试.md'
    2. 标准拖放格式: "d:\测试.md"
    3. 带空格的未加引号路径: d:\测试.md
    
    返回:
        str: 解析后的文件路径
    """
    # 获取当前日志器（可能是极简或正式）
    logger = logging.getLogger()
    
    logger.debug(f"原始用户输入: {user_input}")
    
    # 1. 处理PowerShell拖放格式 (& 'path')
    if user_input.startswith('& '):
        logger.debug("检测到PowerShell拖放格式")
        # 移除前缀
        path_part = user_input[2:].strip()
        
        # 检查单引号包裹
        if path_part.startswith("'") and path_part.endswith("'"):
            logger.debug("单引号包裹路径")
            return path_part[1:-1]
        # 检查双引号包裹
        elif path_part.startswith('"') and path_part.endswith('"'):
            logger.debug("双引号包裹路径")
            return path_part[1:-1]
        else:
            logger.debug("无引号路径")
            return path_part
    
    # 2. 处理标准拖放格式 (双引号包裹)
    elif user_input.startswith('"') and user_input.endswith('"'):
        logger.debug("检测到标准拖放格式（双引号）")
        return user_input[1:-1]
    
    # 3. 直接检查存在的路径（可能包含空格）
    elif os.path.exists(user_input):
        logger.debug("直接存在的路径")
        return user_input
    
    # 4. 尝试去除两端空格后检查
    trimmed = user_input.strip()
    if os.path.exists(trimmed):
        logger.debug("去除空格后存在的路径")
        return trimmed
    
    # 5. 返回原始输入（让后续逻辑处理）
    logger.debug("无法解析路径格式，返回原始输入")
    return user_input

# === 交互模式 ===
def interactive_mode():
    """交互模式主循环"""
    # 获取当前日志器
    logger = logging.getLogger()
    
    while True:
        logger.info("=" * 60)
        logger.info("公文报表转换工具 - 交互模式")
        logger.info("=" * 60)
        logger.info("请选择操作:")
        logger.info("  q. 退出程序")
        logger.info("-" * 60)
        logger.info("提示: 直接拖入文件进行处理")
        logger.info("=" * 60)
        
        try:
            user_input = input("\n请拖入文件或输入选项: ").strip()
            
            # 命令处理
            if user_input.lower() == 'q':
                logger.info("程序退出")
                print("程序退出")
                return
                
            # 文件处理
            else:
                # 路径解析
                file_path = parse_file_path(user_input)
                logger.debug(f"解析后文件路径: {file_path}")
                
                if os.path.exists(file_path):
                    logger.info(f"处理文件: {file_path}")
                    # 使用交互模式处理文件
                    success, output_path = convert_file(file_path, True, True)
                    if success:
                        logger.info(f"转换成功! 输出文件: {output_path}")
                        print(f"\n转换成功! 输出文件: {os.path.basename(output_path)}\n{'='*60}")
                    else:
                        logger.error(f"转换失败: {file_path}")
                        print(f"\n转换失败，请查看日志获取详细信息\n{'='*60}")
                else:
                    logger.warning(f"文件不存在: {file_path}")
                    print(f"\n错误: 文件不存在 - {file_path}\n{'=' * 60}")
                    
        except (EOFError, KeyboardInterrupt):
            logger.info("程序退出")
            print("程序退出")
            return
        except Exception as e:
            logger.error(f"交互模式错误: {str(e)}", exc_info=True)
            print(f"\n错误: 处理失败 - {str(e)}\n{'=' * 60}")

# === 主程序入口 ===
def main():
    """主程序入口（严格遵循初始化顺序）"""
    # 日志系统已在cli_run.py中预初始化和正式初始化
    logger = logging.getLogger()
    
    try:
        # ==== 阶段1：配置系统 ====
        logger.info("加载配置文件...")
        from gongwen_converter.config.config_manager import config_manager
        # (确保配置系统已初始化)
        _ = config_manager.get_logger_config_block()
        logger.info("配置文件加载成功。")

        # ==== 阶段2：业务逻辑 ====
        # 命令行参数处理
        if len(sys.argv) > 1:
            file_path = parse_file_path(sys.argv[1])
            if os.path.exists(file_path):
                logger.info(f"命令行文件处理: {file_path}")
                # 使用交互模式处理文件
                convert_file(file_path, True, True)
                return
            else:
                logger.error(f"文件不存在: {file_path}")
                print(f"错误: 文件不存在 - {file_path}")
                return
                
        # 进入交互模式
        logger.info("进入交互模式。")
        interactive_mode()
        
    except Exception as e:
        # 使用预初始化日志记录严重错误
        logger.critical(f"主程序崩溃: {str(e)}", exc_info=True)
        print(f"严重错误: {str(e)}")

if __name__ == "__main__":
    main()
