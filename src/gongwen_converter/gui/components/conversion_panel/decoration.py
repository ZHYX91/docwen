"""
装饰性绘图模块

提供面板底部的装饰性斜线条纹绘制功能。

主要功能：
- 在面板底部空白区域绘制菱形网格装饰
- 支持主题切换后自动刷新装饰颜色
- 颜色自动跟随主题变化

依赖：
- ttkbootstrap: 提供主题颜色获取
- tkinter: Canvas绑定
- dpi_utils: 支持DPI缩放

使用方式：
    此模块作为 Mixin 类被 ConversionPanel 继承，不应直接实例化。
"""

import logging
import tkinter as tk

import ttkbootstrap as tb

from gongwen_converter.utils.dpi_utils import scale

logger = logging.getLogger(__name__)


class DecorationMixin:
    """
    装饰性绘图混入类
    
    提供面板底部的装饰性菱形网格绘制功能。
    需要与 ConversionPanelBase 一起使用。
    """
    
    def _create_decoration_area(self):
        """
        创建装饰区域
        
        在面板底部创建Canvas用于绘制装饰性斜线条纹。
        Canvas直接放置在主Frame中，不需要额外的容器。
        """
        logger.debug("创建装饰性斜线条纹区域")
        
        # 创建 Canvas 组件（不设置 bg，让它使用默认背景色，绘制时再动态设置）
        self.decoration_canvas = tk.Canvas(
            self,
            highlightthickness=0  # 无边框
        )
        self.decoration_canvas.grid(row=3, column=0, sticky="nsew", pady=(scale(10), 0))
        
        # 绑定 Configure 事件，窗口大小变化时重绘
        self.decoration_canvas.bind('<Configure>', self._on_decoration_canvas_configure)
        
        # 延迟首次绘制（等待布局完成）
        self.after(100, self._draw_diagonal_stripes)
        
        logger.debug("装饰性斜线条纹区域创建完成")
    
    def _on_decoration_canvas_configure(self, event):
        """
        处理 Canvas 大小变化事件
        
        参数:
            event: Configure 事件对象，包含 width 和 height
        """
        # 重绘斜线条纹
        self._draw_diagonal_stripes()
    
    def _get_theme_color(self) -> str:
        """
        获取当前主题的装饰线条颜色
        
        返回:
            str: 十六进制颜色值（如 #95A5A6）
        """
        try:
            # 获取 ttkbootstrap 的主题颜色
            style = self.winfo_toplevel().style
            colors = style.colors
            
            # 使用 success 色（绿色系，代表功能可用状态）
            # 如需更换颜色，只需修改此处即可
            line_color = colors.success
            
            # 如果是字符串直接返回，否则尝试获取十六进制值
            if isinstance(line_color, str):
                result = line_color
            else:
                # 某些版本可能返回颜色对象
                result = str(line_color)
            
            logger.debug(f"获取主题颜色成功: {result}")
            return result
        except Exception as e:
            logger.warning(f"获取主题颜色失败，使用默认灰色: {e}")
            # 默认使用中性灰色
            return "#95A5A6"
    
    def _get_background_color(self) -> str:
        """
        获取当前主题的背景颜色
        
        返回:
            str: 十六进制颜色值
        """
        try:
            style = self.winfo_toplevel().style
            colors = style.colors
            bg_color = colors.bg
            if isinstance(bg_color, str):
                result = bg_color
            else:
                result = str(bg_color)
            logger.debug(f"获取背景颜色成功: {result}")
            return result
        except Exception as e:
            logger.warning(f"获取背景颜色失败，使用默认白色: {e}")
            return "#FFFFFF"
    
    def _blend_colors(self, bg_color: str, fg_color: str, alpha: float) -> str:
        """
        混合两个颜色（模拟透明度效果）
        
        参数:
            bg_color: 背景颜色（十六进制，如 #FFFFFF）
            fg_color: 前景颜色（十六进制，如 #95A5A6）
            alpha: 前景颜色的不透明度（0.0 - 1.0）
            
        返回:
            str: 混合后的颜色（十六进制）
        """
        try:
            # 解析背景颜色
            bg_color = bg_color.lstrip('#')
            if len(bg_color) == 6:
                bg_r = int(bg_color[0:2], 16)
                bg_g = int(bg_color[2:4], 16)
                bg_b = int(bg_color[4:6], 16)
            else:
                bg_r, bg_g, bg_b = 255, 255, 255
            
            # 解析前景颜色
            fg_color = fg_color.lstrip('#')
            if len(fg_color) == 6:
                fg_r = int(fg_color[0:2], 16)
                fg_g = int(fg_color[2:4], 16)
                fg_b = int(fg_color[4:6], 16)
            else:
                fg_r, fg_g, fg_b = 149, 165, 166
            
            # 混合颜色
            r = int(bg_r * (1 - alpha) + fg_r * alpha)
            g = int(bg_g * (1 - alpha) + fg_g * alpha)
            b = int(bg_b * (1 - alpha) + fg_b * alpha)
            
            # 限制范围
            r = max(0, min(255, r))
            g = max(0, min(255, g))
            b = max(0, min(255, b))
            
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception as e:
            logger.debug(f"颜色混合失败: {e}")
            return "#E8E8E8"  # 默认浅灰色
    
    def _draw_diagonal_stripes(self):
        """
        绘制交叉网格（菱形网格）装饰图案
        
        特点：
        - 两组45度斜线交叉形成菱形网格
        - 线条使用主题 success 颜色
        - 透明度约 20%（通过颜色混合实现）
        - 线条间距 10-12px（DPI自适应）
        - 线条宽度 1px
        - 高度可变时自动适应
        """
        logger.debug("开始绘制菱形网格装饰")
        
        if not hasattr(self, 'decoration_canvas'):
            logger.debug("decoration_canvas 不存在，跳过绘制")
            return
        
        canvas = self.decoration_canvas
        
        # 获取 Canvas 尺寸
        try:
            width = canvas.winfo_width()
            height = canvas.winfo_height()
            logger.debug(f"Canvas 尺寸: {width}x{height}")
        except Exception as e:
            logger.debug(f"获取 Canvas 尺寸失败: {e}")
            return
        
        # 尺寸太小时不绘制
        if width < 10 or height < 10:
            logger.debug(f"Canvas 尺寸太小({width}x{height})，跳过绘制")
            return
        
        # 清空现有内容
        canvas.delete("all")
        
        # 获取颜色
        bg_color = self._get_background_color()
        line_color = self._get_theme_color()
        
        # 设置背景色
        canvas.configure(bg=bg_color)
        
        # 混合颜色以实现透明效果
        blended_color = self._blend_colors(bg_color, line_color, 0.20)
        logger.debug(f"绘制菱形网格: 背景={bg_color}, 线条={line_color}, 混合后={blended_color}")
        
        # 绘制参数
        line_spacing = scale(12)  # 线条间距（支持 DPI 缩放）
        line_width = 1  # 线条宽度
        
        # 计算需要绘制的斜线数量
        diagonal_length = width + height  # 对角线方向的最大覆盖长度
        num_lines = int(diagonal_length / line_spacing) + 2
        
        # === 绘制第一组斜线：从左下到右上（/ 方向）===
        for i in range(num_lines):
            offset = i * line_spacing
            
            # 计算线段的起点和终点
            if offset < width:
                x1 = offset
                y1 = height
            else:
                x1 = 0
                y1 = height - (offset - width)
            
            if offset < height:
                x2 = 0
                y2 = height - offset
            else:
                x2 = offset - height
                y2 = 0
            
            # 绘制线段
            if 0 <= x1 <= width and 0 <= x2 <= width and 0 <= y1 <= height and 0 <= y2 <= height:
                canvas.create_line(
                    x1, y1, x2, y2,
                    fill=blended_color,
                    width=line_width,
                    tags="grid"
                )
        
        # === 绘制第二组斜线：从左上到右下（\ 方向）===
        for i in range(num_lines):
            offset = i * line_spacing
            
            # 计算线段的起点和终点
            if offset < width:
                # 起点在顶边
                x1 = offset
                y1 = 0
            else:
                # 起点在左边
                x1 = 0
                y1 = offset - width
            
            if offset < height:
                # 终点在左边
                x2 = 0
                y2 = offset
            else:
                # 终点在底边
                x2 = offset - height
                y2 = height
            
            # 绘制线段
            if 0 <= x1 <= width and 0 <= x2 <= width and 0 <= y1 <= height and 0 <= y2 <= height:
                canvas.create_line(
                    x1, y1, x2, y2,
                    fill=blended_color,
                    width=line_width,
                    tags="grid"
                )
    
    def refresh_decoration(self):
        """
        刷新装饰区域（主题切换时调用）
        
        当用户切换主题后，调用此方法重新绘制斜线条纹
        以适配新主题的颜色
        """
        if hasattr(self, 'decoration_canvas'):
            self._draw_diagonal_stripes()
            logger.debug("装饰区域已刷新")
