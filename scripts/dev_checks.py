"""
开发环境快速自检脚本

验证项目配置文件可解析、GUI 布局计算逻辑正确等基础约束。

使用方式：
    python scripts/dev_checks.py
"""

import logging
import sys
import tomllib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))


def parse_all_toml(directory: Path) -> None:
    for f in sorted(directory.glob("*.toml")):
        tomllib.loads(f.read_text(encoding="utf-8"))


def main() -> int:
    logging.getLogger().setLevel(logging.WARNING)
    logging.getLogger("config_manager").setLevel(logging.WARNING)
    logging.getLogger("safe_log").setLevel(logging.WARNING)

    import docwen.config.schemas.gui
    import docwen.gui.core.window
    import docwen.gui.settings.general_tab
    import docwen.utils.gui_utils

    parse_all_toml(Path("configs"))
    parse_all_toml(Path("src/docwen/i18n/locales"))

    layout = docwen.gui.core.window.MainWindow._compute_grid_layout(
        expand_side_panels=True, triple_mode=True, left_width=400, center_width=460, right_width=300
    )
    assert layout[0]["weight"] == 0
    assert layout[1]["weight"] > 0
    assert layout[2]["weight"] > layout[1]["weight"]
    assert layout[3]["minsize"] == 300
    assert layout[4]["weight"] == 0

    layout2 = docwen.gui.core.window.MainWindow._compute_grid_layout(
        expand_side_panels=False, triple_mode=True, left_width=400, center_width=460, right_width=300
    )
    assert layout2[0]["weight"] == 1
    assert layout2[1]["weight"] == 0
    assert layout2[2]["weight"] == 4
    assert layout2[4]["weight"] == 1

    print("dev checks ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
