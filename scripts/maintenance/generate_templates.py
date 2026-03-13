"""
多语言模板生成脚本

从空白模板出发，为每种支持的语言生成带有本地化样式和占位符的 DOCX 模板。
生成结果默认输出到 samples/generated_templates/，确认无误后可手动替换 templates/ 中的模板，
或使用 --install 参数直接覆盖。

依赖：
- docwen.i18n: 国际化管理（样式名、占位符、YAML 键名）
- docwen.converter.md2docx: MD → DOCX 转换核心

使用方式：
    # 生成所有语言模板到 samples/generated_templates/
    python scripts/maintenance/generate_templates.py

    # 只生成指定语言
    python scripts/maintenance/generate_templates.py --locale en_US ja_JP

    # 生成后直接覆盖 templates/ 中的同名模板
    python scripts/maintenance/generate_templates.py --install
"""

import argparse
import logging
import os
import shutil
import sys
import tempfile
import time

# 路径配置
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(os.path.dirname(SCRIPT_DIR))

# 确保 src 在 sys.path 中
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s - %(name)s - %(message)s",
)
logger = logging.getLogger(__name__)

# 空白模板源文件（存放在脚本同级目录）
BLANK_TEMPLATE_SRC = os.path.join(SCRIPT_DIR, "空白模板.docx")

# templates 目录（运行时临时放置空白模板的位置）
TEMPLATES_DIR = os.path.join(PROJECT_ROOT, "templates")
BLANK_TEMPLATE_DST = os.path.join(TEMPLATES_DIR, "空白模板.docx")

# 默认输出目录
DEFAULT_OUTPUT_DIR = os.path.join(PROJECT_ROOT, "samples", "generated_templates")

# 语言代码 → 模板文件名（不含扩展名）
# 从 locale TOML 的 [meta].template_name 自动发现
LOCALE_TEMPLATE_NAMES = None  # 延迟初始化


def discover_locale_template_names() -> dict:
    """
    扫描 locales/ 目录，读取每个 TOML 的 meta.template_name。

    新增语言只需在 TOML 里填 template_name，无需修改本脚本。

    返回：
        {locale_code: template_name, ...}
    """
    import tomllib

    locales_dir = os.path.join(PROJECT_ROOT, "src", "docwen", "i18n", "locales")
    result = {}
    for filename in sorted(os.listdir(locales_dir)):
        if not filename.endswith(".toml"):
            continue
        locale_code = filename[:-5]  # 去掉 .toml
        filepath = os.path.join(locales_dir, filename)
        with open(filepath, "rb") as f:
            data = tomllib.load(f)
        template_name = data.get("meta", {}).get("template_name")
        if template_name:
            result[locale_code] = template_name
        else:
            logger.debug("跳过 %s（未定义 meta.template_name）", locale_code)
    logger.info("自动发现 %d 个语言模板: %s", len(result), ", ".join(result))
    return result


def reset_singletons():
    """
    重置 I18nManager 和 StyleNameResolver 状态，确保语言切换生效。

    关键：不销毁 StyleNameResolver 单例，而是就地刷新其内部状态。
    因为 injector.py 在模块加载时通过 from docwen.i18n import style_resolver
    持有了对该对象的引用，销毁后创建新实例不会更新 injector 中的引用。
    """
    import docwen.i18n as i18n_module
    from docwen.i18n import style_resolver
    from docwen.i18n.i18n_manager import I18nManager

    # 重置 I18nManager 单例
    I18nManager._instance = None
    I18nManager._initialized = False
    i18n_module._i18n = None

    # 就地刷新 StyleNameResolver（清除 _i18n 引用 + 缓存）
    style_resolver._i18n = None
    style_resolver.clear_cache()


def set_locale(locale: str):
    """
    在内存中设置语言（不写配置文件）。

    参数：
        locale: 语言代码，如 'en_US'
    """
    from docwen.i18n.i18n_manager import I18nManager

    i18n = I18nManager()
    i18n._locale = locale
    i18n._load_translations()
    logger.info("语言已切换为: %s (%s)", locale, i18n.get_current_locale_name())


def get_locale_placeholders(locale: str) -> dict:
    """
    读取指定语言的占位符和 YAML 键名。

    参数：
        locale: 语言代码

    返回：
        字典，包含 placeholders 和 yaml_keys
    """
    import tomllib

    locale_path = os.path.join(PROJECT_ROOT, "src", "docwen", "i18n", "locales", f"{locale}.toml")
    with open(locale_path, "rb") as f:
        data = tomllib.load(f)

    return {
        "placeholders": data.get("placeholders", {}),
        "yaml_keys": data.get("yaml_keys", {}),
    }


def generate_md_content(locale_data: dict) -> str:
    """
    根据语言的占位符信息生成 MD 内容。

    使用 aliases 字段传递本地化标题占位符，利用 _get_fallback_title 的最高
    优先级回退机制，自动填充所有语言版本的标题键。正文使用本地化的 body 占位符。

    参数：
        locale_data: get_locale_placeholders() 返回的字典

    返回：
        Markdown 字符串
    """
    placeholders = locale_data["placeholders"]

    placeholder_title = placeholders.get("title", "title")
    placeholder_body = placeholders.get("body", "body")

    md_lines = [
        "---",
        f'aliases: "{{{{{placeholder_title}}}}}"',
        "---",
        "",
        f"{{{{{placeholder_body}}}}}",
    ]

    return "\n".join(md_lines)


def convert_one_locale(locale: str, output_path: str) -> bool:
    """
    使用指定语言转换 MD 到 DOCX。

    参数：
        locale: 语言代码
        output_path: 输出 DOCX 路径

    返回：
        是否成功
    """
    from docwen.converter.md2docx.core import convert

    # 重置单例 + 切换语言
    reset_singletons()
    set_locale(locale)

    # 读取该语言的占位符
    locale_data = get_locale_placeholders(locale)
    md_content = generate_md_content(locale_data)
    logger.info("生成 MD 内容:\n%s", md_content)

    # 写临时 MD 文件
    with tempfile.NamedTemporaryFile(mode="w", suffix=".md", delete=False, encoding="utf-8") as tmp:
        tmp.write(md_content)
        tmp_md_path = tmp.name

    try:
        result = convert(
            md_path=tmp_md_path,
            output_path=output_path,
            template_name="空白模板",
            progress_callback=lambda msg: logger.info("[%s] %s", locale, msg),
        )
        return result is not None
    finally:
        # 清理临时 MD 文件
        try:
            os.unlink(tmp_md_path)
        except OSError:
            pass


def setup_blank_template():
    """
    将空白模板从 scripts/maintenance/ 复制到 templates/（供 TemplateLoader 使用）。

    返回：
        bool: 是否需要在结束时清理（即 templates/ 中原本没有空白模板）
    """
    if not os.path.exists(BLANK_TEMPLATE_SRC):
        logger.error("空白模板不存在: %s", BLANK_TEMPLATE_SRC)
        raise FileNotFoundError(f"空白模板不存在: {BLANK_TEMPLATE_SRC}")

    already_exists = os.path.exists(BLANK_TEMPLATE_DST)
    shutil.copy2(BLANK_TEMPLATE_SRC, BLANK_TEMPLATE_DST)
    logger.info("已复制空白模板到: %s", BLANK_TEMPLATE_DST)

    # 等待文件系统同步（OneDrive 可能有延迟）
    for i in range(10):
        if os.path.exists(BLANK_TEMPLATE_DST):
            logger.info("空白模板验证通过 (尝试 %d)", i + 1)
            break
        logger.warning("等待空白模板文件可见... (%d/10)", i + 1)
        time.sleep(1)
    else:
        raise FileNotFoundError(f"复制后空白模板仍不可见: {BLANK_TEMPLATE_DST}")

    return not already_exists


def cleanup_blank_template(should_cleanup: bool):
    """
    清理 templates/ 中的临时空白模板。

    参数：
        should_cleanup: 是否需要清理（仅当之前不存在时才删除）
    """
    if should_cleanup and os.path.exists(BLANK_TEMPLATE_DST):
        try:
            os.remove(BLANK_TEMPLATE_DST)
            logger.info("已清理临时空白模板: %s", BLANK_TEMPLATE_DST)
        except OSError as e:
            logger.warning("清理空白模板失败: %s", e)


def install_templates(output_dir: str):
    """
    将生成的模板从输出目录复制到 templates/ 目录。

    参数：
        output_dir: 生成模板所在的目录
    """
    installed = 0
    for filename in os.listdir(output_dir):
        if filename.endswith(".docx"):
            src = os.path.join(output_dir, filename)
            dst = os.path.join(TEMPLATES_DIR, filename)
            shutil.copy2(src, dst)
            logger.info("已安装模板: %s", filename)
            installed += 1
    logger.info("共安装 %d 个模板到 %s", installed, TEMPLATES_DIR)


def main():
    """主函数：解析参数并执行模板生成"""
    global LOCALE_TEMPLATE_NAMES
    LOCALE_TEMPLATE_NAMES = discover_locale_template_names()

    parser = argparse.ArgumentParser(
        description="从空白模板生成多语言 DOCX 模板",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--locale",
        nargs="+",
        choices=list(LOCALE_TEMPLATE_NAMES.keys()),
        default=list(LOCALE_TEMPLATE_NAMES.keys()),
        help="要生成的语言（默认全部）",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"输出目录（默认: {DEFAULT_OUTPUT_DIR}）",
    )
    parser.add_argument(
        "--install",
        action="store_true",
        help="生成后直接复制到 templates/ 目录覆盖同名模板",
    )

    args = parser.parse_args()
    output_dir = args.output_dir
    locales = args.locale

    # 准备输出目录
    os.makedirs(output_dir, exist_ok=True)

    logger.info("=" * 60)
    logger.info("多语言模板生成")
    logger.info("空白模板: %s", BLANK_TEMPLATE_SRC)
    logger.info("输出目录: %s", output_dir)
    logger.info("目标语言: %s", ", ".join(locales))
    logger.info("=" * 60)

    # 复制空白模板到 templates/
    should_cleanup = setup_blank_template()

    results = {}
    try:
        for locale in locales:
            template_name = LOCALE_TEMPLATE_NAMES[locale]
            output_path = os.path.join(output_dir, f"{template_name}.docx")

            logger.info("-" * 40)
            logger.info("生成模板: %s (%s)", template_name, locale)

            try:
                success = convert_one_locale(locale, output_path)
                results[locale] = success
                if success:
                    file_size = os.path.getsize(output_path)
                    logger.info(
                        "✅ %s 成功 → %s (%d bytes)",
                        locale,
                        template_name,
                        file_size,
                    )
                else:
                    logger.error("❌ %s 失败 (convert 返回 None)", locale)
            except Exception as e:
                results[locale] = False
                logger.error("❌ %s 异常: %s", locale, str(e), exc_info=True)
    finally:
        # 清理临时空白模板
        cleanup_blank_template(should_cleanup)

    # 打印汇总
    logger.info("=" * 60)
    logger.info("生成结果汇总")
    logger.info("=" * 60)
    success_count = sum(1 for v in results.values() if v)
    fail_count = len(results) - success_count

    for locale, success in results.items():
        name = LOCALE_TEMPLATE_NAMES[locale]
        status = "✅ 成功" if success else "❌ 失败"
        logger.info("  %s (%s): %s", locale, name, status)

    logger.info("-" * 40)
    logger.info("成功: %d / %d, 失败: %d", success_count, len(results), fail_count)
    logger.info("输出目录: %s", output_dir)

    # 如果指定 --install，复制到 templates/
    if args.install and success_count > 0:
        logger.info("-" * 40)
        logger.info("正在安装模板到 templates/ ...")
        install_templates(output_dir)

    if fail_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
