"""
GUI 冒烟测试脚本

快速验证 PySide6 GUI 能否正常导入和初始化。
使用 --gui 参数可启动主窗口进行简单渲染测试。

使用方式：
    python scripts/gui_smoke.py          # 仅检查导入
    python scripts/gui_smoke.py --gui    # 启动窗口渲染测试
"""

import os
import sys

from PySide6.QtCore import QTimer


def main():
    os.environ.setdefault("PYTHONUTF8", "1")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

    from docwen_gui.app import create_main_window, create_qapplication

    app = create_qapplication([])
    window = create_main_window()
    print("imports ok")

    if "--gui" not in sys.argv:
        window.close()
        print("gui smoke skipped (run with --gui)")
        return 0

    window.show()
    for width, height in [(480, 720), (800, 720), (1200, 720)]:
        window.resize(width, height)
        app.processEvents()
    QTimer.singleShot(600, window.close)
    exit_code = app.exec()
    window.close()
    print("gui smoke ok (--gui)")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
