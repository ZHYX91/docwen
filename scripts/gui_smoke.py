"""
GUI 冒烟测试脚本

快速验证 GUI 组件能否正常导入和初始化。
使用 --gui 参数可启动实际窗口进行简单渲染测试。

使用方式：
    python scripts/gui_smoke.py          # 仅检查导入
    python scripts/gui_smoke.py --gui    # 启动窗口渲染测试
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import ttkbootstrap as tb

from docwen.config.config_manager import ConfigManager
from docwen.gui.components.action_panel.base import ActionPanelBase
from docwen.gui.components.status_bar import StatusBar


def main():
    config_manager = ConfigManager()
    print("imports ok")

    if "--gui" not in sys.argv:
        print("gui smoke skipped (run with --gui)")
        return

    root = tb.Window(themename=config_manager.get_default_theme())
    root.withdraw()

    container = tb.Frame(root)
    container.pack(fill="both", expand=True)

    status_bar = StatusBar(container)
    status_bar.pack(fill="both", expand=True)

    action_panel = ActionPanelBase(container, config_manager)
    action_panel.pack(fill="x")

    for size in [(480, 720), (800, 720), (1200, 720)]:
        root.geometry(f"{size[0]}x{size[1]}")
        root.update_idletasks()

    root.deiconify()
    import time

    start = time.monotonic()
    while time.monotonic() - start < 1.2:
        root.update()
        root.update_idletasks()
        time.sleep(0.03)
    root.destroy()
    print("gui smoke ok (--gui)")


if __name__ == "__main__":
    main()
